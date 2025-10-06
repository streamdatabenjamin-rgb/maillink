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
from googleapiclient.errors import HttpError

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.labels"
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
# Email Helpers
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def create_message(to, subject, body, is_html=True):
    message = MIMEText(body, "html" if is_html else "plain")
    message["to"] = to
    message["subject"] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": raw}

def get_or_create_label(service, label_name="MailMerge"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]

        label_body = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show"
        }
        label = service.users().labels().create(userId="me", body=label_body).execute()
        return label["id"]
    except HttpError as e:
        st.error(f"⚠️ Failed to get/create label: {e}")
        return None

def send_email(service, to, subject, body, label_name="MailMerge"):
    message = create_message(to, subject, body, is_html=True)
    try:
        sent_message = service.users().messages().send(userId="me", body=message).execute()
        label_id = get_or_create_label(service, label_name)
        if label_id:
            service.users().messages().modify(
                userId="me",
                id=sent_message["id"],
                body={"addLabelIds": [label_id]}
            ).execute()
        return sent_message
    except HttpError as e:
        st.error(f"❌ Error sending email: {e}")
        return None

# ========================================
# Robust OAuth Flow (fixed rerun issue)
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None
if "creds_scopes" not in st.session_state:
    st.session_state["creds_scopes"] = None
if "rerun_flag" not in st.session_state:
    st.session_state["rerun_flag"] = False

def needs_reauth():
    if not st.session_state.get("creds"):
        return True
    if st.session_state.get("creds_scopes") != SCOPES:
        return True
    return False

if needs_reauth():
    st.session_state["creds"] = None
    st.session_state["creds_scopes"] = SCOPES

    flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
    flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]

    query_params = st.experimental_get_query_params()
    code_list = query_params.get("code", None)

    if code_list:
        try:
            flow.fetch_token(code=code_list[0])
            creds = flow.credentials
            st.session_state["creds"] = creds.to_json()
            st.session_state["creds_scopes"] = SCOPES
            # trigger a soft rerun by toggling a flag
            st.session_state["rerun_flag"] = not st.session_state["rerun_flag"]
            st.stop()
        except Exception as e:
            st.error(f"⚠️ OAuth token exchange failed: {e}")
            st.stop()
    else:
        auth_url, _ = flow.authorization_url(prompt="consent")
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()
else:
    creds = Credentials.from_authorized_user_info(
        json.loads(st.session_state["creds"]), SCOPES
    )

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
    subject_template = st.text_input("Subject", "Hello {Name}")

    default_body = """<p><b>Dear {Name},</b></p>
    <p>We’d like to invite you to explore our latest <a href="https://phoenixxit.com" target="_blank" style="color:#007BFF;">Phoenixx IT Properties</a>.</p>
    <p>Thank you for your continued support.</p>
    <p>Best Regards,<br><b>Team Phoenixx IT</b></p>"""

    body_template = st.text_area("Body (HTML supported)", default_body, height=250)

    st.markdown("""
    💡 **Tips for formatting:**  
    - Use `<b>Bold</b>`, `<i>Italic</i>`, `<u>Underline</u>`  
    - Add links: `<a href="https://yourlink.com">Click Here</a>`  
    - Change colors: `<span style="color:red;">Text</span>`  
    """)

    st.markdown("### 📄 Email Preview")
    try:
        preview_html = body_template.format(**{col: f"Sample_{col}" for col in df.columns})
        st.markdown(preview_html, unsafe_allow_html=True)
    except Exception:
        st.warning("⚠️ Some placeholders might not match your CSV column names.")

    # ========================================
    # Sending Options
    # ========================================
    st.header("⏱️ Sending Options")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)
    label_name = st.text_input("📛 Gmail Label Name", value="MailMerge")

    # ========================================
    # Send Emails
    # ========================================
    if st.button("🚀 Send Emails"):
        sent_count = 0
        skipped = []
        errors = []

        for idx, row in df.iterrows():
            to_addr_raw = str(row.get("Email", "")).strip()
            to_addr = extract_email(to_addr_raw)

            if not to_addr:
                skipped.append(to_addr_raw)
                continue

            subject = subject_template.format(**row)
            try:
                body = body_template.format(**row)
            except Exception:
                body = body_template

            try:
                send_email(service, to_addr, subject, body, label_name=label_name)
                sent_count += 1
                st.write(f"✅ Sent to: {to_addr}")
                time.sleep(delay)
            except Exception as e:
                errors.append((to_addr, str(e)))

        st.success(f"✅ Successfully sent {sent_count} emails.")
        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to send {len(errors)} emails: {errors}")
