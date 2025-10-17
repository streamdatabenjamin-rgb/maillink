# csv download for all three mode + intractive list previewer
import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
from datetime import datetime, timedelta
import pytz
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

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
# Smart Email Extractor
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

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
# Bold + Link Converter (Verdana)
# ========================================
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
        auth_url, _ = flow.authorization_url(
            prompt="consent", access_type="offline", include_granted_scopes="true"
        )
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
st.info("⚠️ Upload maximum of **70–80 contacts** for smooth operation and to protect your Gmail account.")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    # ✅ FIX: Safe file reading with encoding fallback
    try:
        if uploaded_file.name.endswith("csv"):
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(uploaded_file, encoding="ISO-8859-1")
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"⚠️ Could not read file: {e}")
        st.stop()

    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())
    st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")

    # ========================================
    # Interactive Recipient List Previewer
    # ========================================
    st.markdown("#### 🧮 Manage Recipients (Optional Delete)")

    if "active_rows" not in st.session_state:
        st.session_state.active_rows = df.index.tolist()

    df_display = df.loc[st.session_state.active_rows].reset_index(drop=True)
    delete_indices = []

    for i, row in df_display.iterrows():
        cols = st.columns([0.05, 0.25, 0.7])
        with cols[0]:
            checked = st.checkbox("", key=f"chk_{i}")
            if checked:
                delete_indices.append(i)
        with cols[1]:
            st.markdown(f"**{row.get('Name', '—')}**")
        with cols[2]:
            st.markdown(f"{row.get('Email', '—')}")

    if delete_indices:
        if st.button("🗑️ Delete Selected"):
            remaining = [i for j, i in enumerate(st.session_state.active_rows) if j not in delete_indices]
            st.session_state.active_rows = remaining
            df = df.loc[remaining].reset_index(drop=True)
            st.success(f"✅ Removed {len(delete_indices)} contact(s).")
            st.experimental_rerun()
    else:
        st.caption("☑️ Use checkboxes to remove contacts before sending.")

    # ========================================
    # Rest of your Email Compose, Preview, Send, and CSV logic remains unchanged
    # ========================================
    # (all existing subject/body templates, send modes, delays, ETA, Gmail send/draft logic)
    # --- keep everything below this line identical to your previous working version ---

else:
    st.info("👆 Upload a CSV or Excel file to get started.")
