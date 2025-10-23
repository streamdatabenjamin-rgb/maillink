# ========================================
# Gmail Mail Merge Tool - Batch Send Version (Resumable, Preview + Bulk Labels)
# ========================================
import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Page Setup
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

DONE_FILE = "/tmp/mailmerge_done.csv"
BATCH_SIZE_DEFAULT = 50

# ========================================
# Helpers
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value):
    if not value:
        return None
    m = EMAIL_REGEX.search(str(value))
    return m.group(0) if m else None

def convert_bold(text):
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"\[(.*?)\]\((https?://[^\s)]+)\)", r'<a href="\2" target="_blank">\1</a>', text)
    return text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")

def fetch_message_id_header(service, msg_id):
    try:
        msg = service.users().messages().get(userId="me", id=msg_id, format="metadata", metadataHeaders=["Message-ID"]).execute()
        headers = msg.get("payload", {}).get("headers", [])
        for h in headers:
            if h["name"].lower() == "message-id":
                return h["value"]
    except Exception:
        pass
    return ""

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        new_label = {"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"}
        created = service.users().labels().create(userId="me", body=new_label).execute()
        return created["id"]
    except Exception:
        return None

def generate_preview(subject_template, body_template, row):
    try:
        subject = subject_template.format(**row)
        body_html = convert_bold(body_template.format(**row))
        return subject, body_html
    except Exception:
        return subject_template, convert_bold(body_template)

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
    st.info(f"⚙️ Default batch size: **{BATCH_SIZE_DEFAULT}** emails per run.")
    uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        # Load previous batch if exists
        if os.path.exists(DONE_FILE):
            prev_df = pd.read_csv(DONE_FILE)
            df = pd.merge(df, prev_df, how="outer", on=df.columns.tolist(), suffixes=("", "_old"))
            for col in ["Status", "ThreadId", "RfcMessageId"]:
                if col+"_old" in df.columns:
                    df[col] = df[col].combine_first(df[col+"_old"])
                    df.drop(columns=[col+"_old"], inplace=True)

        for col in ["Status", "ThreadId", "RfcMessageId"]:
            if col not in df.columns:
                df[col] = ""

        st.dataframe(df.head())
        st.info("Include 'ThreadId' and 'RfcMessageId' for follow-ups if needed.")

        subject_template = st.text_input("Subject", "Hello {Name}")
        body_template = st.text_area("Body", "Dear {Name},\n\nWelcome!\n\nThanks,\n**Team**", height=250)
        label_name = st.text_input("Gmail Label", "Mail Merge Sent")
        delay = st.slider("Delay (seconds)", 20, 75, 20)
        send_mode = st.radio("Choose mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        # ===== Live Preview Mail =====
        if not df.empty:
            first_row = df.iloc[0].to_dict()
            preview_container = st.container()
            preview_subject, preview_body = generate_preview(subject_template, body_template, first_row)
            preview_container.markdown(f"### 📬 Preview Email (First Row)")
            preview_container.markdown(f"**Subject:** {preview_subject}")
            preview_container.markdown(f"**Body:**\n{preview_body}", unsafe_allow_html=True)

        if st.button("🚀 Send Emails / Save Drafts"):
            st.session_state.update({
                "sending": True,
                "df": df,
                "subject_template": subject_template,
                "body_template": body_template,
                "label_name": label_name,
                "delay": delay,
                "send_mode": send_mode
            })
            st.rerun()

# ========================================
# SENDING MODE
# ========================================
if st.session_state["sending"]:
    df = st.session_state["df"]
    subject_template = st.session_state["subject_template"]
    body_template = st.session_state["body_template"]
    label_name = st.session_state["label_name"]
    delay = st.session_state["delay"]
    send_mode = st.session_state["send_mode"]

    st.markdown("### 📨 Sending emails... please wait.")
    progress = st.progress(0)
    status_box = st.empty()

    unsent_df = df[df["Status"] != "Sent"].copy()
    total_unsent = len(unsent_df)
    total = len(df)

    if total_unsent == 0:
        st.success("✅ All rows already sent. Nothing to do.")
        st.session_state["sending"] = False
        st.stop()

    if send_mode != "💾 Save as Draft":
        st.info(f"📦 Sending in batches of {BATCH_SIZE_DEFAULT} (Remaining: {total_unsent})")

    label_id = None
    if send_mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    sent_count, skipped, errors = 0, [], []
    batch_limit = BATCH_SIZE_DEFAULT if send_mode != "💾 Save as Draft" else total_unsent
    sent_msg_ids = []

    for idx, row in unsent_df.iloc[:batch_limit].iterrows():
        pct = int(((df["Status"]=="Sent").sum() + 1) / total * 100)
        progress.progress(min(pct, 100))
        status_box.info(f"Processing {sent_count + 1}/{batch_limit}")

        to_addr = extract_email(row.get("Email", ""))
        if not to_addr:
            skipped.append(row.get("Email"))
            df.loc[idx, "Status"] = "Skipped"
            continue

        try:
            subject = subject_template.format(**row)
            body_html = f"<html><body>{convert_bold(body_template.format(**row))}</body></html>"
            message = MIMEText(body_html, "html")
            message["To"] = to_addr
            message["Subject"] = subject

            msg_body = {}
            if send_mode == "↩️ Follow-up (Reply)":
                thread_id = str(row.get("ThreadId", "")).strip()
                rfc_id = str(row.get("RfcMessageId", "")).strip()
                if thread_id and rfc_id:
                    message["In-Reply-To"] = rfc_id
                    message["References"] = rfc_id
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                    msg_body = {"raw": raw, "threadId": thread_id}
                else:
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                    msg_body = {"raw": raw}
            else:
                raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                msg_body = {"raw": raw}

            if send_mode == "💾 Save as Draft":
                draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                df.loc[idx, "ThreadId"] = draft.get("message", {}).get("threadId", "")
                df.loc[idx, "RfcMessageId"] = draft.get("message", {}).get("id", "")
                df.loc[idx, "Status"] = "Draft"
                st.info(f"📝 Draft saved for {to_addr}")
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                df.loc[idx, "RfcMessageId"] = fetch_message_id_header(service, sent_msg.get("id", "")) or sent_msg.get("id", "")
                df.loc[idx, "Status"] = "Sent"
                sent_msg_ids.append(sent_msg["id"])
                st.info(f"✅ Sent to {to_addr}")

            sent_count += 1
            if send_mode != "💾 Save as Draft":
                time.sleep(random.uniform(delay * 0.9, delay * 1.1))
        except Exception as e:
            df.loc[idx, "Status"] = f"Error: {e}"
            errors.append((to_addr, str(e)))

    # ===== Bulk Label Application =====
    if send_mode == "🆕 New Email" and label_id and sent_msg_ids:
        for msg_id in sent_msg_ids:
            try:
                service.users().messages().modify(
                    userId="me",
                    id=msg_id,
                    body={"addLabelIds": [label_id]}
                ).execute()
            except Exception:
                pass

    # Save updated CSV (accumulative)
    df.to_csv(DONE_FILE, index=False)
    st.success(f"✅ Batch complete. Sent {sent_count} emails.")

    if send_mode != "💾 Save as Draft" and total_unsent > BATCH_SIZE_DEFAULT:
        st.info("💡 Re-upload this updated CSV to continue the next batch.")

    st.download_button(
        "⬇️ Download Updated CSV",
        data=open(DONE_FILE, "rb"),
        file_name=os.path.basename(DONE_FILE),
        mime="text/csv",
    )

    st.session_state["sending"] = False
