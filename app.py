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

# Build Gmail API client
creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Upload Recipients
# ========================================
st.header("📤 Upload Recipient List")
st.info("⚠️ Upload maximum of **70–80 contacts** for smooth operation and to protect your Gmail account.")

# 🔁 Backup CSV recovery
if "last_saved_csv" in st.session_state:
    st.info("📁 Backup from previous session available:")
    st.download_button(
        "⬇️ Download Last Saved CSV",
        data=open(st.session_state["last_saved_csv"], "rb"),
        file_name=st.session_state["last_saved_name"],
        mime="text/csv",
    )

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.write("✅ Preview of uploaded data:")

    # Add 'Unsubscribed' column if missing
    if 'Unsubscribed' not in df.columns:
        df['Unsubscribed'] = False

    # Interactive editor
    st.info("📌 Mark 'Unsubscribed' for any email to remove before sending.")
    df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="data_editor")

    # Remove unsubscribed
    if st.button("🗑️ Remove Unsubscribed Emails"):
        before_count = len(df)
        df = df[~df['Unsubscribed']]
        removed_count = before_count - len(df)
        st.success(f"✅ Removed {removed_count} unsubscribed/invalid emails")

    st.info("📌 Include 'ThreadId' and 'RfcMessageId' for follow-ups if needed.")

    # ========================================
    # Compose Email
    # ========================================
    st.header("✍️ Compose Your Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area(
        "Body (supports **bold**, [link](https://example.com), line breaks)",
        """Dear {Name},

Welcome to our **Mail Merge App** demo.

You can add links like [Visit Google](https://google.com)
and preserve formatting exactly.

Thanks,  
**Your Company**""",
        height=250,
    )

    # ========================================
    # Preview Email
    # ========================================
    st.subheader("👁️ Preview Email")
    if not df.empty:
        recipient_options = df["Email"].astype(str).tolist()
        selected_email = st.selectbox("Select recipient to preview", recipient_options)
        try:
            preview_row = df[df["Email"] == selected_email].iloc[0]
            preview_subject = subject_template.format(**preview_row)
            preview_body = body_template.format(**preview_row)
            preview_html = convert_bold(preview_body)
            st.markdown(f'<span style="font-family: Verdana; font-size:16px;"><b>Subject:</b> {preview_subject}</span>', unsafe_allow_html=True)
            st.markdown("---")
            st.markdown(preview_html, unsafe_allow_html=True)
        except KeyError as e:
            st.error(f"⚠️ Missing column in data: {e}")

    # ========================================
    # Label & Timing
    # ========================================
    st.header("🏷️ Label & Timing Options")
    label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")
    delay = st.slider("Delay between emails (seconds)", min_value=30, max_value=300, value=30, step=5, help="Minimum 30s delay recommended.")

    eta_ready = st.button("🕒 Ready to Send / Calculate ETA")
    if eta_ready:
        try:
            total_contacts = len(df)
            min_delay = delay * 0.9
            max_delay = delay * 1.1
            min_total = total_contacts * min_delay
            max_total = total_contacts * max_delay
            local_tz = pytz.timezone("Asia/Kolkata")
            now = datetime.now(local_tz)
            eta_min = now + timedelta(seconds=min_total)
            eta_max = now + timedelta(seconds=max_total)
            st.success(f"📋 Total: {total_contacts}\n⏳ Duration: {min_total/60:.1f}–{max_total/60:.1f} min\n🕒 ETA: {eta_min.strftime('%I:%M %p')}–{eta_max.strftime('%I:%M %p')}")
        except Exception as e:
            st.warning(f"ETA calculation failed: {e}")

    send_mode = st.radio("Choose sending mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

    # ========================================
    # Backup CSV Helper
    # ========================================
    def send_email_backup(service, csv_path):
        try:
            user_profile = service.users().getProfile(userId="me").execute()
            user_email = user_profile.get("emailAddress")
            msg = MIMEMultipart()
            msg["To"] = user_email
            msg["From"] = user_email
            msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            body = MIMEText("Attached is the backup CSV for this mail merge run.\nYou can re-upload for follow-ups.", "plain")
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
    # Send / Draft Processing
    # ========================================
    if st.button("🚀 Send Emails / Save Drafts"):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        skipped, errors = [], []

        if "ThreadId" not in df.columns:
            df["ThreadId"] = None
        if "RfcMessageId" not in df.columns:
            df["RfcMessageId"] = None

        with st.spinner("📨 Processing emails..."):
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
                    # Follow-up mode
                    if send_mode == "↩️ Follow-up (Reply)" and row.get("ThreadId") and row.get("RfcMessageId"):
                        thread_id = str(row["ThreadId"]).strip()
                        rfc_id = str(row["RfcMessageId"]).strip()
                        if thread_id.lower() != "nan" and rfc_id:
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
                        st.info(f"📝 Draft saved for {to_addr}")
                    else:
                        sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                    # Label application
                    if send_mode == "🆕 New Email" and label_id and sent_msg.get("id"):
                        try:
                            service.users().messages().modify(userId="me", id=sent_msg["id"], body={"addLabelIds": [label_id]}).execute()
                        except Exception:
                            st.warning(f"⚠️ Could not apply label to {to_addr}")

                    # Delay ±10%
                    if delay > 0:
                        time.sleep(random.uniform(delay * 0.9, delay * 1.1))

                    # Fetch Message-ID for CSV
                    message_id_header = None
                    for _ in range(5):
                        time.sleep(random.uniform(2, 4))
                        try:
                            msg_detail = service.users().messages().get(
                                userId="me", id=sent_msg.get("id", ""), format="metadata", metadataHeaders=["Message-ID"]
                            ).execute()
                            headers = msg_detail.get("payload", {}).get("headers", [])
                            for h in headers:
                                if h.get("name", "").lower() == "message-id":
                                    message_id_header = h.get("value")
                                    break
                            if message_id_header:
                                break
                        except Exception:
                            continue

                    df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                    df.loc[idx, "RfcMessageId"] = message_id_header or ""
                    sent_count += 1

                except Exception as e:
                    errors.append((to_addr, str(e)))

        # ========================================
        # Summary + Backup CSV
        # ========================================
        if send_mode == "💾 Save as Draft":
            st.success(f"📝 Saved {sent_count} draft(s) to Gmail Drafts.")
        else:
            st.success(f"✅ Successfully processed {sent_count} emails.")

        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to process {len(errors)} emails: {errors}")

        # Save CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"Updated_{safe_label}_{timestamp}.csv"
        file_path = os.path.join("/tmp", file_name)
        df.to_csv(file_path, index=False)
        st.session_state["last_saved_csv"] = file_path
        st.session_state["last_saved_name"] = file_name
        st.success("✅ Updated CSV auto-saved safely on server.")
        st.download_button("⬇️ Download Updated CSV", data=open(file_path, "rb"), file_name=file_name, mime="text/csv")

        # Email backup
        send_email_backup(service, file_path)
