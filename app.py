import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
from datetime import datetime, timedelta
import pytz
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

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        label_obj = {"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"}
        created_label = service.users().labels().create(userId="me", body=label_obj).execute()
        return created_label["id"]
    except Exception as e:
        st.warning(f"Could not get/create label: {e}")
        return None

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
        st.experimental_rerun()
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

service = build("gmail", "v1", credentials=Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES))

# ========================================
# Upload Page
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
    
    # Delete unsubscribed
    if "Unsubscribed" in df.columns:
        if st.button("🗑️ Delete unsubscribed rows"):
            df = df[df["Unsubscribed"].astype(str).str.lower() != "yes"]
            st.success("✅ Unsubscribed rows deleted.")
    
    # Compose email
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area("Body", "Dear {Name},\n\nWelcome!\n\nThanks,\nYour Company", height=250)

    # Preview
    if not df.empty:
        recipient_options = df["Email"].astype(str).tolist()
        selected_email = st.selectbox("Select recipient to preview", recipient_options)
        preview_row = df[df["Email"] == selected_email].iloc[0]
        st.markdown(f"**Subject:** {subject_template.format(**preview_row)}")
        st.markdown(convert_bold(body_template.format(**preview_row)), unsafe_allow_html=True)

    # Label and delay
    label_name = st.text_input("Gmail label", value="Mail Merge Sent")
    delay = st.slider("Delay between emails (seconds)", 20, 75, 20, 1)
    eta_ready = st.button("🕒 Calculate ETA")
    if eta_ready:
        total_contacts = len(df)
        min_total = total_contacts * delay * 0.9
        max_total = total_contacts * delay * 1.1
        now = datetime.now(pytz.timezone("Asia/Kolkata"))
        st.success(f"⏳ ETA: {now + timedelta(seconds=min_total)} – {now + timedelta(seconds=max_total)}")

    send_mode = st.radio("Sending mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])
    if st.button("🚀 Send Emails / Save Drafts"):
        st.session_state["sending_mode"] = True
        st.session_state["send_payload"] = {
            "df": df.to_dict(),
            "subject_template": subject_template,
            "body_template": body_template,
            "label_name": label_name,
            "delay": delay,
            "send_mode": send_mode,
        }
        st.experimental_rerun()

# ========================================
# Sending Progress Page
# ========================================
if st.session_state.get("sending_mode"):
    st.header("📬 Sending in Progress...")
    data = st.session_state["send_payload"]
    df = pd.DataFrame(data["df"])
    subject_template = data["subject_template"]
    body_template = data["body_template"]
    label_name = data["label_name"]
    delay = data["delay"]
    send_mode = data["send_mode"]

    progress_bar = st.progress(0)
    status_text = st.empty()

    label_id = get_or_create_label(service, label_name)
    sent_count, errors, skipped = 0, [], []

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

            raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
            msg_body = {"raw": raw}

            if send_mode == "💾 Save as Draft":
                draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                sent_msg = draft.get("message", {})
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

            if send_mode == "🆕 New Email" and label_id and sent_msg.get("id"):
                service.users().messages().modify(userId="me", id=sent_msg["id"], body={"addLabelIds": [label_id]}).execute()

            # Delay ±10%
            time.sleep(random.uniform(delay * 0.9, delay * 1.1))

            df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
            df.loc[idx, "RfcMessageId"] = sent_msg.get("id", "")
            sent_count += 1
            progress_bar.progress(sent_count / len(df))
            status_text.text(f"Sent {sent_count}/{len(df)}: {to_addr}")

        except Exception as e:
            errors.append((to_addr, str(e)))

    st.success(f"✅ Completed! Sent {sent_count}/{len(df)} emails.")
    if skipped:
        st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
    if errors:
        st.error(f"❌ Failed {len(errors)} emails: {errors}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"MailMerge_Sent_{timestamp}.csv"
    file_path = os.path.join("/tmp", file_name)
    df.to_csv(file_path, index=False)
    st.download_button("⬇️ Download Sent CSV", data=open(file_path, "rb"), file_name=file_name, mime="text/csv")
    send_email_backup(service, file_path)

    if st.button("⬅ Back to Mail Merge"):
        st.session_state["sending_mode"] = False
        st.experimental_rerun()
