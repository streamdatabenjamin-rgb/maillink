import streamlit as st
import pandas as pd
import base64
import time
import datetime
import re
import random
import json
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="📧 Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge System")
st.markdown("Upload your CSV and send personalized emails safely with Google Gmail API.")

# ========================================
# Gmail Authentication
# ========================================
if "credentials" not in st.session_state:
    st.session_state["credentials"] = None
if "ui_locked" not in st.session_state:
    st.session_state["ui_locked"] = False

def authenticate_gmail():
    flow = Flow.from_client_secrets_file(
        "client_secret.json",
        scopes=["https://www.googleapis.com/auth/gmail.modify"]
    )
    flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
    auth_url, _ = flow.authorization_url(prompt="consent")
    st.write("🔑 [Authorize Gmail Access]({})".format(auth_url))
    auth_code = st.text_input("Enter the authorization code:")
    if auth_code:
        try:
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            st.session_state["credentials"] = creds
            st.success("✅ Gmail Authentication Successful!")
        except Exception as e:
            st.error(f"❌ Authentication Failed: {e}")

if not st.session_state["credentials"]:
    authenticate_gmail()
else:
    creds = st.session_state["credentials"]
    service = build("gmail", "v1", credentials=creds)

# ========================================
# File Upload
# ========================================
uploaded_file = st.file_uploader("📂 Upload CSV file", type=["csv"], help="Upload maximum of 70–80 rows for smooth performance.")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.success(f"✅ File uploaded successfully! ({len(df)} rows loaded)")

    # Manual deletion option (unsubscribe)
    st.subheader("🗑️ Remove Unsubscribed Rows (Optional)")
    st.caption("Check rows to delete from this run.")
    check_delete = st.checkbox("Enable Manual Row Deletion")
    if check_delete:
        delete_indices = []
        for i in range(len(df)):
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(df.iloc[i].to_dict())
            with col2:
                if st.checkbox(f"Delete Row {i+1}", key=f"del_{i}"):
                    delete_indices.append(i)
        if st.button("🧹 Delete Selected Rows"):
            df.drop(delete_indices, inplace=True)
            df.reset_index(drop=True, inplace=True)
            st.success(f"Deleted {len(delete_indices)} rows. Remaining: {len(df)}")

    st.dataframe(df)

# ========================================
# Mail Settings
# ========================================
st.subheader("⚙️ Mail Settings")
mode = st.radio("Select Mode:", ["New Email", "Follow-up (Reply)", "Save as Draft"], horizontal=True)
delay = st.number_input("⏱️ Delay (seconds between emails)", min_value=30, max_value=600, value=30, step=5, help="Minimum 30 seconds for safety.")
if "list_size" not in st.session_state:
    st.session_state["list_size"] = 0

if uploaded_file:
    st.session_state["list_size"] = len(df)
    total_seconds = len(df) * delay
    eta = datetime.datetime.now() + datetime.timedelta(seconds=total_seconds)
    st.info(f"📅 Estimated Completion Time: **{eta.strftime('%I:%M %p')}**")

subject = st.text_input("📧 Email Subject")
body = st.text_area("📝 Email Body (supports placeholders like {Name})")

# ========================================
# Helper Functions
# ========================================
def create_message(to, subject, body_text):
    msg = MIMEText(body_text, "html")
    msg["to"] = to
    msg["subject"] = subject
    return {"raw": base64.urlsafe_b64encode(msg.as_bytes()).decode()}

def send_message(service, user_id, message):
    return service.users().messages().send(userId=user_id, body=message).execute()

def save_draft(service, user_id, message):
    create_draft = {"message": message}
    return service.users().drafts().create(userId=user_id, body=create_draft).execute()

# ========================================
# UI Lock Helper
# ========================================
def lock_ui():
    st.session_state["ui_locked"] = True
    st.markdown("""
        <style>
        button, input, select, textarea, [data-testid="stFileUploadDropzone"], div[data-baseweb="radio"] {
            pointer-events: none !important;
            opacity: 0.5 !important;
        }
        </style>
    """, unsafe_allow_html=True)
    st.info("🔒 UI locked. Please wait while processing...")

def unlock_ui():
    st.session_state["ui_locked"] = False
    st.markdown("""
        <style>
        button, input, select, textarea, [data-testid="stFileUploadDropzone"], div[data-baseweb="radio"] {
            pointer-events: auto !important;
            opacity: 1 !important;
        }
        </style>
    """, unsafe_allow_html=True)

# Apply lock style dynamically
if st.session_state["ui_locked"]:
    lock_ui()

# ========================================
# Send / Draft Button
# ========================================
if uploaded_file and subject and body:
    send_clicked = st.button("🚀 Send Emails / Save Drafts", key="send_main")

    if send_clicked:
        try:
            lock_ui()
            sent_count, error_count = 0, 0

            with st.spinner("📨 Processing... Please wait and do not close this tab."):
                for idx, row in df.iterrows():
                    try:
                        personalized_body = body
                        for col in df.columns:
                            personalized_body = personalized_body.replace(f"{{{col}}}", str(row[col]))
                        to = row.get("Email") or row.get("email")
                        if not to:
                            continue
                        message = create_message(to, subject, personalized_body)

                        if mode == "New Email":
                            send_message(service, "me", message)
                        elif mode == "Follow-up (Reply)":
                            send_message(service, "me", message)
                        elif mode == "Save as Draft":
                            save_draft(service, "me", message)

                        sent_count += 1
                        time.sleep(delay)

                    except Exception as e:
                        error_count += 1
                        st.warning(f"⚠️ Error sending to {row.get('Email', 'Unknown')}: {e}")
                        unlock_ui()  # Option B – unlock early on any failure
                        break

            st.success(f"✅ Process Completed: {sent_count} sent successfully, {error_count} failed.")
            unlock_ui()

        except Exception as e:
            st.error(f"❌ Unexpected Error: {e}")
            unlock_ui()

else:
    st.warning("⚠️ Please upload a CSV and fill in Subject & Body before sending.")
