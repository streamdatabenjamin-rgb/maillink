# ========================================
# Gmail Mail Merge Tool - Optimized Batch + Resume (Silent Mode)
# ========================================
import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
import tempfile
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool (Optimized Batch + Resume)")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/gmail.compose",
]

CLIENT_CONFIG = {
    "web": {
        "client_id": st.secrets["gmail"]["client_id"],
        "client_secret": st.secrets["gmail"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [st.secrets["gmail"]["redirect_uri"]],
    }
}

# ========================================
# Constants
# ========================================
DONE_FILE = os.path.join(tempfile.gettempdir(), "mailmerge_done.json")
BATCH_SIZE_DEFAULT = 50

# ========================================
# Cached Gmail label lookup
# ========================================
@st.cache_data(show_spinner=False)
def cached_labels(service):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        return {l["name"].lower(): l["id"] for l in labels}
    except Exception:
        return {}

# ========================================
# Recovery Logic
# ========================================
if os.path.exists(DONE_FILE) and not st.session_state.get("done", False):
    try:
        with open(DONE_FILE, "r") as f:
            done_info = json.load(f)
        file_path = done_info.get("file")
        if file_path and os.path.exists(file_path):
            st.success("✅ Previous mail merge completed successfully.")
            st.download_button(
                "⬇️ Download Updated CSV",
                data=open(file_path, "rb"),
                file_name=os.path.basename(file_path),
                mime="text/csv",
            )
            if st.button("🔁 Reset for New Run"):
                os.remove(DONE_FILE)
                st.session_state.clear()
                st.experimental_rerun()
            st.stop()
    except Exception:
        pass

# ========================================
# Helpers
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def convert_bold(text):
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\\1</b>", text)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\\2" style="color:#1a73e8; text-decoration:underline;" target="_blank">\\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"""
    <html><body style="font-family: Verdana, sans-serif; font-size: 14px; line-height: 1.6;">
        {text}
    </body></html>
    """

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = cached_labels(service)
        if label_name.lower() in labels:
            return labels[label_name.lower()]
        created = service.users().labels().create(
            userId="me",
            body={"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
        ).execute()
        return created["id"]
    except Exception:
        return None

def send_email_backup(service, csv_path):
    try:
        user_email = service.users().getProfile(userId="me").execute()["emailAddress"]
        msg = MIMEMultipart()
        msg["To"] = user_email
        msg["From"] = user_email
        msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        msg.attach(MIMEText("Attached is the backup CSV for your mail merge run.", "plain"))
        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        st.info(f"📧 Backup CSV emailed to {user_email}")
    except Exception as e:
        st.warning(f"⚠️ Could not send backup email: {e}")

# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
else:
    code = st.experimental_get_query_params().get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        flow.fetch_token(code=code[0])
        creds = flow.credentials
        st.session_state["creds"] = creds.to_json()
        st.rerun()
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Session Setup
# ========================================
if "sending" not in st.session_state:
    st.session_state["sending"] = False
if "done" not in st.session_state:
    st.session_state["done"] = False

# ========================================
# MAIN UI
# ========================================
if not st.session_state["sending"]:
    st.header("📤 Upload Recipient List")
    uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.lower().endswith("csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        for col in ["ThreadId", "RfcMessageId", "Status"]:
            if col not in df.columns:
                df[col] = ""

        pending = df[df["Status"] != "Sent"]
        if not pending.empty:
            df = pending.reset_index(drop=True)

        st.dataframe(df.head())
        subject_template = st.text_input("Subject", "Hello {Name}")
        body_template = st.text_area(
            "Body",
            """Dear {Name},\n\nWelcome to **Mail Merge App** demo.\n\nThanks,  \n**Your Company**""",
            height=250,
        )
        label_name = st.text_input("Gmail label", "Mail Merge Sent")
        delay = st.slider("Delay (seconds)", 10, 75, 20)
        workers = st.slider("Parallel workers", 1, 5, 3)
        send_mode = st.radio("Choose mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        if st.button("🚀 Send Emails / Save Drafts"):
            st.session_state.update({
                "sending": True,
                "df": df,
                "subject_template": subject_template,
                "body_template": body_template,
                "label_name": label_name,
                "delay": delay,
                "send_mode": send_mode,
                "workers": workers,
            })
            st.rerun()

# ========================================
# Sending Mode
# ========================================
if st.session_state["sending"]:
    df = st.session_state["df"]
    subject_template = st.session_state["subject_template"]
    body_template = st.session_state["body_template"]
    label_name = st.session_state["label_name"]
    delay = st.session_state["delay"]
    send_mode = st.session_state["send_mode"]
    workers = st.session_state["workers"]

    st.markdown("<h3>📨 Sending emails... please wait.</h3>", unsafe_allow_html=True)
    progress = st.progress(0)
    status_box = st.empty()

    label_id = None
    if send_mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    total = len(df)
    sent_count, skipped, errors = 0, [], []

    def send_one(idx, row):
        nonlocal sent_count
        to_addr = extract_email(str(row.get("Email", "")).strip())
        if not to_addr:
            df.loc[idx, "Status"] = "Skipped"
            return (None, "Skipped", None)
        try:
            subject = subject_template.format(**row)
            body_html = convert_bold(body_template.format(**row))
            message = MIMEText(body_html, "html")
            message["To"] = to_addr
            message["Subject"] = subject
            raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
            msg_body = {"raw": raw}

            if send_mode == "💾 Save as Draft":
                draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                return (draft.get("message", {}).get("threadId", ""), draft.get("message", {}).get("id", ""), "Draft")
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                return (sent_msg.get("threadId", ""), sent_msg.get("id", ""), "Sent")
        except Exception as e:
            return (None, str(e), "Error")

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(send_one, idx, row): idx for idx, row in df.iterrows()}
        for i, future in enumerate(as_completed(futures)):
            idx = futures[future]
            pct = int(((i + 1) / total) * 100)
            progress.progress(pct)
            try:
                thread, msg_id, status = future.result()
                df.loc[idx, "ThreadId"] = thread
                df.loc[idx, "RfcMessageId"] = msg_id
                df.loc[idx, "Status"] = status
                if status == "Sent":
                    sent_count += 1
            except Exception as e:
                errors.append((idx, str(e)))
            status_box.info(f"Processed {i + 1}/{total}")

    # Save CSV
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
    file_name = f"Updated_{safe_label}_{timestamp}.csv"
    file_path = os.path.join(tempfile.gettempdir(), file_name)
    df.to_csv(file_path, index=False, encoding='utf-8-sig')

    try:
        send_email_backup(service, file_path)
    except Exception as e:
        st.warning(f"Backup email failed: {e}")

    with open(DONE_FILE, "w") as f:
        json.dump({"done_time": str(datetime.now()), "file": file_path}, f)

    st.session_state["sending"] = False
    st.session_state["done"] = True
    st.session_state["summary"] = {"sent": sent_count, "errors": errors, "skipped": skipped}
    st.rerun()

# ========================================
# Completion
# ========================================
if st.session_state["done"]:
    summary = st.session_state.get("summary", {})
    st.success(f"✅ Process completed. Sent: {summary.get('sent', 0)}")
    if summary.get("errors"):
        st.error(f"❌ {len(summary['errors'])} errors occurred.")
    if summary.get("skipped"):
        st.warning(f"⚠️ Skipped some invalid emails.")
    with open(DONE_FILE, "r") as f:
        done_info = json.load(f)
    file_path = done_info.get("file")
    if file_path and os.path.exists(file_path):
        st.download_button(
            "⬇️ Download Updated CSV",
            data=open(file_path, "rb"),
            file_name=os.path.basename(file_path),
            mime="text/csv",
        )
    if st.button("🔁 New Run / Reset"):
        if os.path.exists(DONE_FILE):
            os.remove(DONE_FILE)
        st.session_state.clear()
        st.experimental_rerun()
