import streamlit as st
import pandas as pd
import base64
import time
import re
import json
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/gmail.labels"]

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
# Email Utilities
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    """Extracts the first valid email from a string, or None if not found."""
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def create_message(to, subject, body, is_html=True):
    """Create a MIME message with optional HTML support."""
    if is_html:
        message = MIMEText(body, "html")
    else:
        message = MIMEText(body)
    message["to"] = to
    message["subject"] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": raw}

# ========================================
# Gmail Label Helpers
# ========================================
def get_or_create_label(service, label_name="MailMerge"):
    """Get Gmail label by name or create it if it doesn't exist."""
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for lbl in labels:
            if lbl["name"].lower() == label_name.lower():
                return lbl["id"]

        # Create new label if not found
        label_obj = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        new_label = service.users().labels().create(userId="me", body=label_obj).execute()
        return new_label["id"]

    except HttpError as error:
        st.error(f"Error creating label: {error}")
        return None

def send_email(service, to, subject, body, label_id=None):
    """Send an email and apply a label."""
    message = create_message(to, subject, body, is_html=True)
    sent_message = service.users().messages().send(userId="me", body=message).execute()

    if label_id:
        try:
            service.users().messages().modify(
                userId="me",
                id=sent_message["id"],
                body={"addLabelIds": [label_id]},
            ).execute()
        except HttpError as e:
            st.warning(f"Could not label email to {to}: {e}")
    return sent_message

# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
else:
    code = st.query_params.get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        flow.fetch_token(code=code)
        creds = flow.credentials
        st.session_state["creds"] = creds.to_json()
        st.rerun()  # ✅ Fixed: replaces deprecated experimental_rerun
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

# Build Gmail API client
creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Upload Recipients
# ========================================
st.header("📤 Upload Recipient List")
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())

    # ========================================
    # Email Template
    # ========================================
    st.header("✍️ Compose Your Email")
    st.markdown("You can use **HTML** for formatting (bold, links, lists, etc.). Example:")
    st.code("""
<b>Hello {Name},</b><br><br>
We’d love to connect with you!<br>
<a href="https://yourwebsite.com">Visit our site</a><br><br>
Best,<br>Your Team
    """, language="html")

    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area("Body (HTML supported)", "<b>Dear {Name},</b><br><br>This is a test mail.<br><br>Regards,<br>Your Company")

    # Label and Delay
    st.header("🏷️ Gmail Label & ⏱️ Delay")
    label_name = st.text_input("Enter Gmail label name to tag sent emails", "MailMerge")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)

    if st.button("🚀 Send Emails"):
        sent_count = 0
        skipped = []
        errors = []

        label_id = get_or_create_label(service, label_name)

        for idx, row in df.iterrows():
            to_addr_raw = str(row.get("Email", "")).strip()
            to_addr = extract_email(to_addr_raw)

            if not to_addr:
                skipped.append(to_addr_raw)
                continue

            try:
                subject = subject_template.format(**row)
                body = body_template.format(**row)
            except KeyError as e:
                st.error(f"Missing placeholder in data: {e}")
                continue

            try:
                send_email(service, to_addr, subject, body, label_id=label_id)
                sent_count += 1
                time.sleep(delay)
            except Exception as e:
                errors.append((to_addr, str(e)))

        # Summary
        st.success(f"✅ Successfully sent {sent_count} emails.")
        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to send {len(errors)} emails: {errors}")
