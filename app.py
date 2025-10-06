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
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

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
# Smart Email Extractor
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    """Extracts the first valid email from a string, or None if not found."""
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

# ========================================
# Gmail Helpers
# ========================================
def create_message(to, subject, body, is_html=True):
    """Create email message (supports HTML or plain text)."""
    message = MIMEText(body, "html" if is_html else "plain")
    message["to"] = to
    message["subject"] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": raw}

def send_email(service, to, subject, body):
    """Send the email using Gmail API."""
    message = create_message(to, subject, body, is_html=True)
    return service.users().messages().send(userId="me", body=message).execute()

# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(
        json.loads(st.session_state["creds"]), SCOPES
    )
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
        auth_url, _ = flow.authorization_url(prompt="consent")
        st.markdown(
            f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account."
        )
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
    # Email Template (with HTML)
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

    # ========================================
    # Live HTML Preview
    # ========================================
    st.markdown("### 📄 Email Preview")
    try:
        preview_html = body_template.format(**{col: f"Sample_{col}" for col in df.columns})
        st.markdown(preview_html, unsafe_allow_html=True)
    except Exception:
        st.warning("⚠️ Some placeholders might not match your CSV column names.")

    # ========================================
    # Delay Option
    # ========================================
    st.header("⏱️ Sending Options")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)

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
                body = body_template  # fallback if placeholder mismatch

            try:
                send_email(service, to_addr, subject, body)
                sent_count += 1
                st.write(f"✅ Sent to: {to_addr}")
                time.sleep(delay)
            except Exception as e:
                errors.append((to_addr, str(e)))

        # ========================================
        # Final Summary
        # ========================================
        st.success(f"✅ Successfully sent {sent_count} emails.")
        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to send {len(errors)} emails: {errors}")
