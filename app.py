# ========================================
# Gmail Mail Merge Tool - Stable Multi-Tasking Version---works well woth 200 datas too
# ========================================
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
st.title("📧 Gmail Mail Merge Tool (Stable Multi-tasking Version)")

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
st.info("⚠️ Upload maximum of 70–80 contacts recommended for smooth operation.")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        try:
            df = pd.read_csv(uploaded_file, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="ISO-8859-1")
        except pd.errors.EmptyDataError:
            st.error("❌ Uploaded CSV appears empty or corrupted.")
            st.stop()
        except pd.errors.ParserError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=None, engine="python")
    else:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception:
            st.error("❌ Unable to read Excel file.")
            st.stop()

    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())
    st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")

    # ========================================
    # Email Template
    # ========================================
    st.header("✍️ Compose Your Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area(
        "Body (supports **bold**, [link](https://example.com), and line breaks)",
        """Dear {Name},\n\nWelcome to our **Mail Merge App** demo.\n\nThanks,\n**Your Company**""",
        height=250,
    )

    # ========================================
    # Preview Section
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

            st.markdown(
                f'<span style="font-family: Verdana, sans-serif; font-size:16px;"><b>Subject:</b> {preview_subject}</span>',
                unsafe_allow_html=True
            )
            st.markdown("---")
            st.markdown(preview_html, unsafe_allow_html=True)
        except KeyError as e:
            st.error(f"⚠️ Missing column in data: {e}")

    # ========================================
    # Label & Timing Options
    # ========================================
    st.header("🏷️ Label & Timing Options")
    label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")
    delay = st.slider(
        "Delay between emails (seconds)",
        min_value=20,
        max_value=75,
        value=20,
        step=1
    )

    # ETA
    eta_ready = st.button("🕒 Ready to Send / Calculate ETA")
    if eta_ready:
        total_contacts = len(df)
        total_seconds = total_contacts * delay
        total_minutes = total_seconds / 60
        local_tz = pytz.timezone("Asia/Kolkata")
        now_local = datetime.now(local_tz)
        eta_end = now_local + timedelta(seconds=total_seconds)
        st.success(f"📋 Total Recipients: {total_contacts}\n⏳ Estimated Duration: {total_minutes:.1f} min\n🕒 ETA: {now_local.strftime('%I:%M %p')} – {eta_end.strftime('%I:%M %p')}")

    # ========================================
    # Send Mode
    # ========================================
    send_mode = st.radio(
        "Choose sending mode",
        ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"]
    )

    # ========================================
    # Main Send/Draft Button
    # ========================================
    if st.button("🚀 Send Emails / Save Drafts"):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        skipped, errors = [], []

        progress_bar = st.progress(0)
        status_text = st.empty()

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

                raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                msg_body = {"raw": raw}

                if send_mode == "💾 Save as Draft":
                    draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                    sent_msg = draft.get("message", {})
                else:
                    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                # Delay with random jitter
                time.sleep(random.uniform(delay * 0.9, delay * 1.1))

                # Fetch Message-ID every 10 emails only
                message_id_header = ""
                if sent_count % 10 == 0 and send_mode != "💾 Save as Draft":
                    try:
                        msg_detail = service.users().messages().get(
                            userId="me",
                            id=sent_msg.get("id", ""),
                            format="metadata",
                            metadataHeaders=["Message-ID"],
                        ).execute()
                        headers = msg_detail.get("payload", {}).get("headers", [])
                        for h in headers:
                            if h.get("name", "").lower() == "message-id":
                                message_id_header = h.get("value")
                                break
                    except Exception:
                        pass

                # Apply label for new emails
                if send_mode == "🆕 New Email" and label_id and sent_msg.get("id"):
                    try:
                        service.users().messages().modify(
                            userId="me",
                            id=sent_msg["id"],
                            body={"addLabelIds": [label_id]},
                        ).execute()
                    except Exception:
                        pass

                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                df.loc[idx, "RfcMessageId"] = message_id_header or ""
                sent_count += 1

            except Exception as e:
                err_msg = str(e)
                # Handle Gmail rate limits
                if "Rate Limit" in err_msg or "quota" in err_msg.lower():
                    st.warning("⚠️ Gmail limit reached. Pausing for 10 minutes...")
                    time.sleep(10 * 60)
                    continue
                errors.append((to_addr, err_msg))
                continue

            # Update progress bar
            progress = int((idx + 1) / len(df) * 100)
            progress_bar.progress(progress)
            status_text.text(f"📤 Sending {idx+1}/{len(df)} | ✅ Sent: {sent_count} | ⚠️ Skipped: {len(skipped)} | ❌ Failed: {len(errors)}")

        # Summary
        if send_mode == "💾 Save as Draft":
            st.success(f"📝 Saved {sent_count} draft(s).")
        else:
            st.success(f"✅ Successfully processed {sent_count} emails.")

        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to process {len(errors)} emails: {errors}")

        # Manual CSV Download
        csv = df.to_csv(index=False).encode("utf-8")
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"{safe_label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        st.download_button(
            "⬇️ Download Updated CSV",
            csv,
            file_name,
            "text/csv",
            key="manual_download_final"
        )
