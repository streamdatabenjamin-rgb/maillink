import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
import math
import threading
from datetime import datetime, timedelta
import pytz
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool (with Follow-up Replies + Draft Save)")

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
# Helpers
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None


class SafeDict(dict):
    def __missing__(self, key):
        return ""


def safe_format(template: str, row: pd.Series):
    try:
        return template.format_map(SafeDict(**row.to_dict()))
    except Exception:
        return template


def convert_bold(text):
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\2" style="color:#1a73e8; text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"""
    <html>
        <body style="font-family: Verdana, sans-serif; font-size: 14px; line-height: 1.6;">
            {text}
        </body>
    </html>
    """


def with_backoff(fn, max_retries=5, initial_delay=1.0, factor=2.0):
    delay = initial_delay
    for attempt in range(1, max_retries + 1):
        try:
            return fn()
        except Exception as e:
            if attempt == max_retries:
                raise
            time.sleep(delay)
            delay *= factor


# ========================================
# Gmail Label Helper
# ========================================
def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        label_obj = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        created_label = service.users().labels().create(userId="me", body=label_obj).execute()
        return created_label["id"]
    except Exception as e:
        st.warning(f"Could not get/create label: {e}")
        return None


# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

creds = None
if st.session_state.get("creds"):
    try:
        creds = Credentials.from_authorized_user_info(
            json.loads(st.session_state["creds"]), scopes=SCOPES
        )
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            st.session_state["creds"] = creds.to_json()
    except Exception:
        st.warning("Session credentials invalid — please re-authorize.")
        st.session_state["creds"] = None
        st.experimental_rerun()

if not creds:
    code = st.experimental_get_query_params().get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        flow.fetch_token(code=code[0])
        creds = flow.credentials
        st.session_state["creds"] = creds.to_json()
        st.experimental_rerun()
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(
            prompt="consent", access_type="offline", include_granted_scopes="true"
        )
        st.markdown(
            f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account."
        )
        st.stop()

creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
if creds.expired and creds.refresh_token:
    creds.refresh(Request())
    st.session_state["creds"] = creds.to_json()
service = build("gmail", "v1", credentials=creds)

# ========================================
# Upload Recipients
# ========================================
st.header("📤 Upload Recipient List")
st.info("⚠️ Upload maximum of **70–80 contacts** for smooth operation and to protect your Gmail account.")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])
if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())
    st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")

    df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="recipient_editor_inline")

    st.header("✍️ Compose Your Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area(
        "Body (supports **bold**, [link](https://example.com), and line breaks)",
        """Dear {Name},

Welcome to our **Mail Merge App** demo.

You can add links like [Visit Google](https://google.com)
and preserve formatting exactly.

Thanks,  
**Your Company**""",
        height=250,
    )

    st.subheader("👁️ Preview Email")
    if not df.empty:
        recipient_options = df["Email"].astype(str).tolist()
        selected_email = st.selectbox("Select recipient to preview", recipient_options)
        try:
            preview_row = df[df["Email"] == selected_email].iloc[0]
            preview_subject = safe_format(subject_template, preview_row)
            preview_body = safe_format(body_template, preview_row)
            preview_html = convert_bold(preview_body)

            st.markdown(
                f'<span style="font-family: Verdana, sans-serif; font-size:16px;"><b>Subject:</b> {preview_subject}</span>',
                unsafe_allow_html=True
            )
            st.markdown("---")
            st.markdown(preview_html, unsafe_allow_html=True)
        except KeyError as e:
            st.error(f"⚠️ Missing column in data: {e}")

    st.header("🏷️ Label & Timing Options")
    label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")
    delay = st.slider("Delay between emails (seconds)", 20, 75, 20, 1)

    send_mode = st.radio(
        "Choose sending mode",
        ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"]
    )

    def send_email_backup(service, csv_path):
        try:
            user_profile = service.users().getProfile(userId="me").execute()
            user_email = user_profile.get("emailAddress")
            msg = MIMEMultipart()
            msg["To"] = user_email
            msg["From"] = user_email
            msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            body = MIMEText("Attached is the backup CSV file for your mail merge run.", "plain")
            msg.attach(body)
            with open(csv_path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
            msg.attach(part)
            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
            service.users().messages().send(userId="me", body={"raw": raw}).execute()
            st.info(f"📧 Backup CSV emailed to your Gmail inbox ({user_email}).")
            return user_email
        except Exception as e:
            st.warning(f"⚠️ Could not send backup email: {e}")
            return None

    # Background collector
    def collect_ids_and_send_csv(service, df, file_path, label_name, user_email):
        updated = False
        for idx, row in df.iterrows():
            if not row.get("ThreadId") or not row.get("RfcMessageId"):
                try:
                    msg_detail = service.users().messages().get(
                        userId="me",
                        id=row.get("SentMsgId", ""),
                        format="metadata",
                        metadataHeaders=["Message-ID"],
                    ).execute()
                    headers = msg_detail.get("payload", {}).get("headers", [])
                    for h in headers:
                        if h.get("name", "").lower() == "message-id":
                            df.loc[idx, "RfcMessageId"] = h.get("value")
                            break
                    df.loc[idx, "ThreadId"] = msg_detail.get("threadId", "")
                    updated = True
                except Exception:
                    continue
            time.sleep(random.uniform(2, 4))
        if updated:
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
                updated_path = os.path.join("/tmp", f"Updated_{safe_label}_{timestamp}.csv")
                df.to_csv(updated_path, index=False)

                msg = MIMEMultipart()
                msg["To"] = user_email
                msg["From"] = user_email
                msg["Subject"] = f"📁 Updated Mail Merge CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                body = MIMEText("Attached is your updated CSV with Gmail ThreadId and Message-ID.", "plain")
                msg.attach(body)
                with open(updated_path, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(updated_path))
                part["Content-Disposition"] = f'attachment; filename="{os.path.basename(updated_path)}"'
                msg.attach(part)
                raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
                service.users().messages().send(userId="me", body={"raw": raw}).execute()
            except Exception as e:
                print(f"Background update failed: {e}")

    # ========================================
    # 🚀 Send Emails / Save Drafts
    # ========================================
    if st.button("🚀 Send Emails / Save Drafts"):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        errors, skipped = [], []
        progress_bar = st.progress(0)
        status_box = st.empty()

        if "ThreadId" not in df.columns:
            df["ThreadId"] = None
        if "RfcMessageId" not in df.columns:
            df["RfcMessageId"] = None
        df["SentMsgId"] = None

        for idx, row in df.iterrows():
            to_addr = extract_email(str(row.get("Email", "")).strip())
            if not to_addr:
                skipped.append(row.get("Email"))
                continue
            try:
                subject = safe_format(subject_template, row)
                body_html = convert_bold(safe_format(body_template, row))
                message = MIMEText(body_html, "html")
                profile = service.users().getProfile(userId="me").execute()
                message["From"] = profile.get("emailAddress")
                message["To"] = to_addr
                message["Subject"] = subject

                raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                msg_body = {"raw": raw}

                if send_mode == "💾 Save as Draft":
                    sent_msg = with_backoff(lambda: service.users().drafts().create(userId="me", body={"message": msg_body}).execute())
                    sent_msg = sent_msg.get("message", {})
                else:
                    sent_msg = with_backoff(lambda: service.users().messages().send(userId="me", body=msg_body).execute())

                df.loc[idx, "SentMsgId"] = sent_msg.get("id", "")
                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                sent_count += 1
                progress = math.floor((sent_count / len(df)) * 100)
                progress_bar.progress(min(progress, 100))
                status_box.markdown(f"✅ Sent: {sent_count}/{len(df)} — {to_addr}")

                if delay > 0:
                    time.sleep(random.uniform(delay * 0.9, delay * 1.1))

            except Exception as e:
                errors.append((to_addr, str(e)))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"Updated_{safe_label}_{timestamp}.csv"
        file_path = os.path.join("/tmp", file_name)
        df.to_csv(file_path, index=False)
        st.success("✅ Preliminary CSV saved successfully.")
        user_email = send_email_backup(service, file_path)

        if user_email:
            bg_thread = threading.Thread(
                target=collect_ids_and_send_csv,
                args=(service, df.copy(), file_path, label_name, user_email),
                daemon=True,
            )
            bg_thread.start()
            st.success("🕓 Background job started to fetch Gmail IDs and send updated CSV to your inbox.")

        if skipped:
            st.warning(f"⚠️ Skipped invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Some errors occurred: {errors}")
