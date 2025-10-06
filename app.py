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

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = ["https://www.googleapis.com/auth/gmail.send",
          "https://www.googleapis.com/auth/gmail.modify",
          "https://www.googleapis.com/auth/gmail.labels"]

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

def convert_bold(text):
    if not text:
        return ""
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = text.replace("\n", "<br>")
    return text

def get_or_create_label(service, label_name):
    if not label_name:
        return None
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    for label in labels:
        if label["name"].lower() == label_name.lower():
            return label["id"]
    label_obj = {"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"}
    created_label = service.users().labels().create(userId="me", body=label_obj).execute()
    return created_label["id"]

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

service = build("gmail", "v1", credentials=creds)

# ========================================
# Upload Recipients
# ========================================
st.header("📤 Upload Recipient List")
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith("csv") else pd.read_excel(uploaded_file)
    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())

    # ========================================
    # Email Template
    # ========================================
    st.header("✍️ Compose Your Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area("Body (use **bold** for emphasis)",
                                 "Dear {Name},\n\nThis is a **test mail**.\n\nRegards,\nYour Company",
                                 height=200)

    # ========================================
    # Email Preview
    # ========================================
    st.subheader("👁️ Preview Your Email")
    recipient_options = df["Email"].astype(str).tolist()
    selected_email = st.selectbox("Select recipient to preview", recipient_options)
    preview_row = df[df["Email"] == selected_email].iloc[0]
    st.markdown(f"**Subject:** {subject_template.format(**preview_row)}")
    st.markdown("---")
    st.markdown("**Email Body Preview:**")
    st.markdown(convert_bold(body_template.format(**preview_row)), unsafe_allow_html=True)

    # ========================================
    # Timing & Optional Label
    # ========================================
    st.header("⏱️ Timing & Label Options")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)
    use_label = st.checkbox("Apply Gmail label to sent emails?")
    label_name = ""
    if use_label:
        label_name = st.text_input("Enter label name")

    # ========================================
    # Send Emails
    # ========================================
    if st.button("🚀 Send Emails"):
        label_id = get_or_create_label(service, label_name) if use_label and label_name.strip() else None
        sent_count, skipped, errors = 0, [], []

        with st.spinner("📨 Sending emails..."):
            for _, row in df.iterrows():
                to_addr = extract_email(str(row.get("Email", "")).strip())
                if not to_addr:
                    skipped.append(str(row.get("Email", "")))
                    continue
                try:
                    message = MIMEText(convert_bold(body_template.format(**row)), "html")
                    message["to"] = to_addr
                    message["subject"] = subject_template.format(**row)
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                    msg_body = {"raw": raw}
                    if label_id:
                        msg_body["labelIds"] = [label_id]
                    service.users().messages().send(userId="me", body=msg_body).execute()
                    sent_count += 1
                    time.sleep(delay)
                except Exception as e:
                    errors.append((to_addr, str(e)))

        st.success(f"✅ Successfully sent {sent_count} emails.")
        if skipped:
            st.warning(f"⚠️ Skipped invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to send {len(errors)} emails: {errors}")
