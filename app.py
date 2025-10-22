# ========================================
# Gmail Mail Merge Tool - Batch Send Version (app.py)
# ========================================
# Drop-in Streamlit app that sends emails in configurable batches (default 50)
# - Preserves ThreadId, RfcMessageId, and Status columns
# - Allows resumable batch sending by re-uploading the CSV after each batch
# - Stores final CSV in session_state for reliable download after reruns

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
from googleapiclient.errors import HttpError

# ----------------------------------------
# Page setup
# ----------------------------------------
st.set_page_config(page_title="Gmail Mail Merge - Batch Sender", layout="wide")
st.title("📧 Gmail Mail Merge Tool — Batch Send Mode")

# ----------------------------------------
# Config / constants
# ----------------------------------------
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

EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    value = str(value).strip()
    match = EMAIL_REGEX.search(value)
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


# ----------------------------------------
# OAuth flow (same pattern as your working version)
# ----------------------------------------
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
        auth_url, _ = flow.authorization_url(
            prompt="consent", access_type="offline", include_granted_scopes="true"
        )
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

# build service
creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ----------------------------------------
# Upload CSV / Excel
# ----------------------------------------
st.header("📤 Upload Recipient List")
st.info("Recommended: use batches of ~50. Ensure your file has an 'Email' column.")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if not uploaded_file:
    st.info("Upload a CSV/XLSX to begin. You can re-upload after each batch to continue where you left off.")
    st.stop()

# read file
try:
    if uploaded_file.name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="ISO-8859-1")
        except pd.errors.EmptyDataError:
            st.error("Uploaded CSV appears empty or corrupted.")
            st.stop()
        except pd.errors.ParserError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Could not read uploaded file: {e}")
    st.stop()

st.write("✅ Preview of uploaded data:")
st.dataframe(df.head())

# ensure helper columns
for col in ["Status", "ThreadId", "RfcMessageId"]:
    if col not in df.columns:
        df[col] = ""

# summary
total = len(df)
processed = (df['Status'] == 'Sent').sum() if 'Status' in df.columns else (df['ThreadId'] != '').sum()
remaining = total - processed
st.markdown(f"**Rows:** {total} — **Processed:** {processed} — **Remaining:** {remaining}")

# ----------------------------------------
# Email template
# ----------------------------------------
st.header("✍️ Compose Your Email")
subject_template = st.text_input("Subject", "Hello {Name}")
body_template = st.text_area(
    "Body (supports **bold**, [link](https://example.com), and line breaks)",
    """Dear {Name},\n\nWelcome to our **Mail Merge App** demo.\n\nThanks,\n**Your Company**""",
    height=250,
)

st.subheader("👁️ Preview Email")
if not df.empty and 'Email' in df.columns:
    sample_email = df['Email'].astype(str).iloc[0]
    try:
        preview_row = df.iloc[0]
        preview_subject = subject_template.format(**preview_row)
        preview_body = body_template.format(**preview_row)
        preview_html = convert_bold(preview_body)
        st.markdown(f'<span style="font-family: Verdana, sans-serif; font-size:16px;"><b>Subject:</b> {preview_subject}</span>', unsafe_allow_html=True)
        st.markdown('---')
        st.markdown(preview_html, unsafe_allow_html=True)
    except Exception:
        st.info("Preview not available for the first row due to missing fields.")
else:
    st.error("Make sure your CSV has an 'Email' column.")
    st.stop()

# ----------------------------------------
# Batch controls
# ----------------------------------------
st.header("⚙️ Batch Settings")
BATCH_SIZE = st.number_input("Batch size (rows per run)", min_value=5, max_value=500, value=50, step=5)
delay = st.slider("Delay between emails (seconds)", min_value=5, max_value=120, value=20, step=1)
label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")
send_mode = st.radio("Choose sending mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"]) 

# determine start/end indexes by scanning Status/ThreadId
if (df['Status'] == 'Sent').any():
    # use Status marker if present
    first_unprocessed = df[df['Status'] != 'Sent'].index.min()
else:
    # fallback: find first empty ThreadId / empty RfcMessageId
    mask = (df['ThreadId'] == '') & (df['RfcMessageId'] == '')
    if mask.any():
        first_unprocessed = df[mask].index.min()
    else:
        first_unprocessed = None

if first_unprocessed is None:
    st.success("🎉 All rows look processed (no candidates for sending).")
else:
    start_idx = int(first_unprocessed)
    end_idx = min(start_idx + BATCH_SIZE, total)  # inclusive stop at end_idx-1
    st.info(f"This run will process rows {start_idx+1} → {end_idx} (total {end_idx-start_idx}).")

    # send button
    if st.button("🚀 Send This Batch"):
        label_id = None
        try:
            # try to create/get label
            labels = service.users().labels().list(userId='me').execute().get('labels', [])
            for lab in labels:
                if lab.get('name', '').lower() == label_name.lower():
                    label_id = lab.get('id')
                    break
            if not label_id:
                label_obj = {
                    'name': label_name,
                    'labelListVisibility': 'labelShow',
                    'messageListVisibility': 'show',
                }
                created = service.users().labels().create(userId='me', body=label_obj).execute()
                label_id = created.get('id')
        except Exception as e:
            st.warning(f"Could not ensure label: {e}")

        sub_df = df.iloc[start_idx:end_idx].copy()
        progress_bar = st.progress(0)
        status_text = st.empty()
        sent_count = 0
        errors = []
        start_time = datetime.now()

        for local_i, (global_i, row) in enumerate(sub_df.iterrows(), start=1):
            to_addr = extract_email(str(row.get('Email', '')).strip())
            if not to_addr:
                df.at[global_i, 'Status'] = 'Invalid Email'
                errors.append((global_i, 'Invalid Email'))
                continue

            try:
                subject = subject_template.format(**row)
                body_html = convert_bold(body_template.format(**row))
                message = MIMEText(body_html, 'html')
                message['To'] = to_addr
                message['Subject'] = subject

                raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
                msg_body = {'raw': raw}

                # actual send / draft
                if send_mode == '💾 Save as Draft':
                    sent_msg = service.users().drafts().create(userId='me', body={'message': msg_body}).execute().get('message', {})
                else:
                    try:
                        sent_msg = service.users().messages().send(userId='me', body=msg_body).execute()
                    except HttpError as e:
                        status_code = getattr(e.resp, 'status', None)
                        if status_code in [429, 403]:
                            # exponential backoff (capped)
                            backoff = min(600, 30 * (2 ** (sent_count // 20)))
                            st.warning(f"⚠️ Gmail limit hit (HTTP {status_code}). Waiting {int(backoff/60)} minutes...")
                            time.sleep(backoff)
                            # try again once
                            try:
                                sent_msg = service.users().messages().send(userId='me', body=msg_body).execute()
                            except Exception as e2:
                                df.at[global_i, 'Status'] = f'Error:{str(e2)}'
                                errors.append((global_i, str(e2)))
                                continue
                        else:
                            raise

                # apply label (for new emails)
                if send_mode == '🆕 New Email' and label_id and sent_msg.get('id'):
                    try:
                        service.users().messages().modify(userId='me', id=sent_msg['id'], body={'addLabelIds': [label_id]}).execute()
                    except Exception:
                        pass

                # store ids and status
                df.at[global_i, 'ThreadId'] = sent_msg.get('threadId', '') or df.at[global_i, 'ThreadId']

                # try to store RFC Message-ID metadata (best-effort)
                try:
                    if send_mode != '💾 Save as Draft' and sent_msg.get('id'):
                        msg_detail = service.users().messages().get(userId='me', id=sent_msg.get('id'), format='metadata', metadataHeaders=['Message-ID']).execute()
                        headers = msg_detail.get('payload', {}).get('headers', [])
                        for h in headers:
                            if h.get('name', '').lower() == 'message-id':
                                df.at[global_i, 'RfcMessageId'] = h.get('value')
                                break
                except Exception:
                    # not critical
                    pass

                df.at[global_i, 'Status'] = 'Sent'
                sent_count += 1

                # update small UI every 3 emails
                if local_i % 3 == 0 or local_i == len(sub_df):
                    elapsed = (datetime.now() - start_time).total_seconds()
                    remaining = max(0, (end_idx - start_idx - local_i) * delay)
                    progress_bar.progress(int(local_i / len(sub_df) * 100))
                    status_text.text(f"📤 Batch {start_idx+1}-{end_idx} | Sent: {sent_count}/{len(sub_df)} | Errors: {len(errors)} | ETA (approx): {int(remaining/60)} min")

                # polite randomized delay
                time.sleep(random.uniform(delay * 0.9, delay * 1.1))

            except Exception as e:
                df.at[global_i, 'Status'] = f'Error:{str(e)}'
                errors.append((global_i, str(e)))
                continue

        # batch finished
        st.success(f"✅ Batch finished — Sent: {sent_count} | Errors: {len(errors)}")
        if errors:
            st.error(f"Failed rows: {errors}")

        # save updated CSV bytes to session_state so download survives rerun
        csv_bytes = df.to_csv(index=False).encode('utf-8')
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"{safe_label}_after_{end_idx}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        st.session_state['csv_bytes'] = csv_bytes
        st.session_state['file_name'] = file_name

# show download button if any CSV is ready
if 'csv_bytes' in st.session_state:
    st.download_button(
        "⬇️ Download Updated CSV (use this to continue next batch)",
        st.session_state['csv_bytes'],
        st.session_state['file_name'],
        "text/csv",
        key='download_updated_csv'
    )

# optional: quick guidance
st.markdown('''
**Usage tips:**
- Upload the same CSV repeatedly. Each run processes the next unprocessed batch (rows with Status != 'Sent').
- Use a conservative delay (20–30s) to reduce Gmail quota issues.
- If you hit rate limits, wait 10–30 minutes before retrying — the app uses exponential backoff when Gmail returns 429/403.
- Keep batch sizes small (50 recommended) for reliability in Streamlit.
''')

# End of app
