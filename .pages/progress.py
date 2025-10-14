import streamlit as st
import pandas as pd
import base64, time, os, random, re, json
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

st.set_page_config(page_title="Mail Merge Progress", layout="wide")
st.title("📨 Sending Progress")

if "mailmerge_params" not in st.session_state or "active_csv" not in st.session_state:
    st.warning("⚠️ No active mail merge found. Please return to main page.")
    st.page_link("app.py", label="⬅️ Back to Home")
    st.stop()

params = st.session_state["mailmerge_params"]
df = pd.read_csv(st.session_state["active_csv"])

# Gmail setup
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/gmail.compose",
]
creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# Utilities
def convert_bold(text):
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\2" style="color:#1a73e8;text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f'<html><body style="font-family: Verdana; font-size:14px;">{text}</body></html>'

def get_or_create_label(service, name):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for lbl in labels:
            if lbl["name"].lower() == name.lower():
                return lbl["id"]
        created = service.users().labels().create(userId="me", body={"name": name}).execute()
        return created["id"]
    except Exception:
        return None

# --- UI placeholders ---
progress_text = st.empty()
progress_bar = st.progress(0)
log_area = st.empty()

delay = params["delay"]
label_name = params["label_name"]
send_mode = params["send_mode"]
subject_template = params["subject_template"]
body_template = params["body_template"]

label_id = get_or_create_label(service, label_name)
sent_count, failed = 0, []

progress_text.info(f"🚀 Starting {send_mode} process... please wait.")

for i, row in df.iterrows():
    try:
        to_addr = str(row.get("Email", "")).strip()
        if not to_addr or "@" not in to_addr:
            continue

        subject = subject_template.format(**row)
        body_html = convert_bold(body_template.format(**row))

        # --- Prepare email ---
        message = MIMEText(body_html, "html")
        message["To"], message["Subject"] = to_addr
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

        msg_body = {"raw": raw}

        if send_mode.startswith("🆕"):  # New Mail
            sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
            if label_id:
                service.users().messages().modify(
                    userId="me",
                    id=sent_msg["id"],
                    body={"addLabelIds": [label_id]},
                ).execute()

        elif send_mode.startswith("↩️"):  # Follow-up (Reply)
            thread_id = row.get("ThreadId")
            if pd.isna(thread_id) or not str(thread_id).strip():
                log_area.write(f"⚠️ Skipping {to_addr}: missing ThreadId.")
                continue

            msg_body["threadId"] = str(thread_id)
            sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

        elif send_mode.startswith("💾"):  # Save Draft
            draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
            df.loc[i, "DraftId"] = draft.get("id", "")

        # --- Store message IDs ---
        if send_mode != "💾 Save as Draft":
            df.loc[i, "ThreadId"] = sent_msg.get("threadId", "")
            df.loc[i, "RfcMessageId"] = sent_msg.get("id", "")

        sent_count += 1
        progress_bar.progress((i + 1) / len(df))
        log_area.write(f"✅ {send_mode} → {to_addr}")

        if delay > 0 and send_mode != "💾 Save as Draft":
            time.sleep(random.uniform(delay * 0.9, delay * 1.1))

    except Exception as e:
        failed.append((to_addr, str(e)))
        log_area.write(f"❌ Failed {to_addr}: {e}")

# --- Save updated CSV ---
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
file_name = f"Updated_{safe_label}_{timestamp}.csv"
file_path = os.path.join("/tmp", file_name)
df.to_csv(file_path, index=False)
st.session_state["last_saved_csv"] = file_path
st.session_state["last_saved_name"] = file_name

st.success(f"✅ Done! {send_mode} completed for {sent_count}/{len(df)}.")
st.download_button("⬇️ Download Updated CSV", data=open(file_path, "rb"), file_name=file_name, mime="text/csv")

# --- Send backup CSV ---
try:
    user_profile = service.users().getProfile(userId="me").execute()
    user_email = user_profile.get("emailAddress")
    msg = MIMEMultipart()
    msg["To"] = user_email
    msg["From"] = user_email
    msg["Subject"] = f"📁 Mail Merge Backup - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    msg.attach(MIMEText("Attached is your backup CSV file.", "plain"))
    with open(file_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
    part["Content-Disposition"] = f'attachment; filename="{os.path.basename(file_path)}"'
    msg.attach(part)
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()
    st.info(f"📧 Backup CSV emailed to {user_email}")
except Exception as e:
    st.warning(f"⚠️ Could not send backup email: {e}")

if failed:
    st.error(f"❌ {len(failed)} failed:\n{failed}")

st.page_link("app.py", label="⬅️ Back to Mail Merge Home")

