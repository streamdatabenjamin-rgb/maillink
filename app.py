import streamlit as st
import pandas as pd
import base64
import time
import random
import re
import json
import os
from datetime import datetime
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="centered")
st.title("📧 Gmail Mail Merge")
st.write("Send personalized emails via Gmail API with custom templates and CSV upload.")

# ========================================
# Authentication Setup
# ========================================
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

if "credentials" not in st.session_state:
    st.session_state.credentials = None

def authenticate_gmail():
    if "client_secret.json" not in os.listdir():
        st.error("❌ Missing 'client_secret.json'. Please upload your OAuth credentials file.")
        return None

    flow = Flow.from_client_secrets_file("client_secret.json", scopes=SCOPES, redirect_uri="urn:ietf:wg:oauth:2.0:oob")
    auth_url, _ = flow.authorization_url(prompt="consent")
    st.markdown(f"[🔗 Click here to Authorize Gmail Access]({auth_url})")

    code = st.text_input("Enter Authorization Code:")
    if st.button("Submit Code"):
        try:
            flow.fetch_token(code=code)
            creds = flow.credentials
            st.session_state.credentials = creds.to_json()
            st.success("✅ Gmail successfully authenticated.")
            st.rerun()
        except Exception as e:
            st.error(f"Authentication failed: {e}")

authenticate_gmail()
if not st.session_state.credentials:
    st.stop()

# ========================================
# Gmail API Setup
# ========================================
creds = Credentials.from_authorized_user_info(json.loads(st.session_state.credentials), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Helper Functions
# ========================================
def extract_email(text):
    match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text)
    return match.group(0) if match else None

def convert_bold(text):
    return re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)

def send_email_backup(service, file_path):
    """Send backup CSV to the authenticated Gmail."""
    try:
        message = MIMEText("Backup of updated mail merge CSV file attached.")
        message["To"] = "me"
        message["Subject"] = "Mail Merge Backup CSV"
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
    except Exception as e:
        st.warning(f"⚠️ Backup email failed: {e}")

def get_or_create_label(service, label_name):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        new_label = service.users().labels().create(userId="me", body={"name": label_name}).execute()
        return new_label["id"]
    except Exception:
        return None

# ========================================
# CSV Upload
# ========================================
uploaded_file = st.file_uploader("📂 Upload CSV File", type=["csv"])
if uploaded_file is None:
    st.stop()

df = pd.read_csv(uploaded_file)
st.write("📋 Data Preview:", df.head())

# ========================================
# Email Template and Options
# ========================================
subject_template = st.text_input("✉️ Subject Template")
body_template = st.text_area("📝 Email Body Template (Use **bold** for highlights)", height=200)
send_mode = st.radio("Send Mode", ["📤 Send", "💾 Save as Draft", "↩️ Follow-up (Reply)"], horizontal=True)
label_name = st.text_input("🏷️ Gmail Label (for organization)", value="MailMerge")
delay = st.slider("⏱️ Delay between emails (seconds)", 1, 120, 60)

# ========================================
# Safe Mail Sending (Resistant to UI Interruptions)
# ========================================
if "send_in_progress" not in st.session_state:
    st.session_state["send_in_progress"] = False
if "send_mode" not in st.session_state:
    st.session_state["send_mode"] = None
if "df_cache" not in st.session_state:
    st.session_state["df_cache"] = None
if "label_cache" not in st.session_state:
    st.session_state["label_cache"] = None
if "delay_cache" not in st.session_state:
    st.session_state["delay_cache"] = None
if "subject_cache" not in st.session_state:
    st.session_state["subject_cache"] = None
if "body_cache" not in st.session_state:
    st.session_state["body_cache"] = None

# When the user clicks "Send"
if st.button("🚀 Send Emails / Save Drafts"):
    st.session_state["send_in_progress"] = True
    st.session_state["df_cache"] = df.copy()
    st.session_state["label_cache"] = label_name
    st.session_state["delay_cache"] = delay
    st.session_state["send_mode"] = send_mode
    st.session_state["subject_cache"] = subject_template
    st.session_state["body_cache"] = body_template
    st.rerun()

# Continue sending even if UI is touched
if st.session_state["send_in_progress"]:
    df = st.session_state["df_cache"]
    label_name = st.session_state["label_cache"]
    delay = st.session_state["delay_cache"]
    send_mode = st.session_state["send_mode"]
    subject_template = st.session_state["subject_cache"]
    body_template = st.session_state["body_cache"]

    with st.spinner("📨 Sending emails... Please wait (don’t touch the app)."):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        skipped, errors = [], []

        if "ThreadId" not in df.columns:
            df["ThreadId"] = None
        if "RfcMessageId" not in df.columns:
            df["RfcMessageId"] = None

        for idx, row in df.iterrows():
            to_addr = extract_email(str(row.get("Email", "")).strip())
            if not to_addr:
                skipped.append(row.get("Email"))
                continue

            try:
                subject = subject_template.format(**row)
                body_html = convert_bold(body_template.format(**row))
                message = MIMEText(body_html, "html")
                message["To"] = to_addr
                message["Subject"] = subject

                msg_body = {}
                if send_mode == "↩️ Follow-up (Reply)" and "ThreadId" in row and "RfcMessageId" in row:
                    thread_id = str(row["ThreadId"]).strip()
                    rfc_id = str(row["RfcMessageId"]).strip()
                    if thread_id and thread_id.lower() != "nan" and rfc_id:
                        message["In-Reply-To"] = rfc_id
                        message["References"] = rfc_id
                        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                        msg_body = {"raw": raw, "threadId": thread_id}
                    else:
                        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                        msg_body = {"raw": raw}
                else:
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                    msg_body = {"raw": raw}

                if send_mode == "💾 Save as Draft":
                    draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                    sent_msg = draft.get("message", {})
                else:
                    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                if delay > 0:
                    time.sleep(random.uniform(delay * 0.9, delay * 1.1))

                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                df.loc[idx, "RfcMessageId"] = sent_msg.get("id", "")
                sent_count += 1

            except Exception as e:
                errors.append((to_addr, str(e)))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"Updated_{safe_label}_{timestamp}.csv"
        file_path = os.path.join("/tmp", file_name)
        df.to_csv(file_path, index=False)
        st.success(f"✅ {sent_count} emails processed successfully.")

        st.download_button(
            "⬇️ Download Updated CSV",
            data=open(file_path, "rb"),
            file_name=file_name,
            mime="text/csv",
        )

        send_email_backup(service, file_path)

    st.session_state["send_in_progress"] = False
