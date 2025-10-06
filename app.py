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
st.title("📧 Gmail Mail Merge Tool (No Label)")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
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
    """Extracts the first valid email from a string, or None if not found."""
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def convert_bold(text):
    """
    Convert **bold** syntax to <b>bold</b> while escaping other HTML.
    Also converts newlines to <br> for HTML rendering.
    """
    if not text:
        return ""
    # Escape HTML special chars first
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    # Convert **bold** to <b>...</b>
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    # Convert line breaks to <br>
    text = text.replace("\n", "<br>")
    return text

class SafeDict(dict):
    def __missing__(self, key):
        return ""

def compute_textarea_height(text: str, min_h=120, max_h=800):
    """Estimate textarea height based on number of lines (simple heuristic)."""
    lines = text.count("\n") + 1
    height = min(max_h, max(min_h, 20 + lines * 24))
    return height

# ========================================
# OAuth Flow (store creds in session_state)
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    try:
        creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
    except Exception as e:
        st.warning("Invalid saved credentials — please re-authorize.")
        st.session_state["creds"] = None
        st.experimental_rerun()
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

# Build Gmail API client
creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Upload Recipients
# ========================================
st.header("📤 Upload Recipient List (CSV / Excel)")
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

df = None
if uploaded_file:
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.write("✅ Preview of uploaded data:")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Could not read the uploaded file: {e}")

# ========================================
# Layout: Compose (left) + Preview (right)
# ========================================
st.header("✍️ Compose Your Email (plain text + **bold** only)")
col1, col2 = st.columns([2, 1])

with col1:
    subject_template = st.text_input("Subject", "Hello {Name}")
    # Keep / restore body in session_state so height can be computed reliably
    if "body_template" not in st.session_state:
        st.session_state["body_template"] = "Dear {Name},\n\nThis is a **test mail**.\n\nRegards,\nYour Company"
    # compute dynamic height from saved body
    dynamic_height = compute_textarea_height(st.session_state["body_template"])
    body_template = st.text_area("Body (use **bold** for emphasis)", value=st.session_state["body_template"], height=dynamic_height)
    st.session_state["body_template"] = body_template  # save current content

    st.markdown("**Options**")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)

    st.markdown("---")

    # send button
    send_clicked = st.button("🚀 Send Emails")

with col2:
    st.markdown("### Preview")
    if df is not None and len(df) > 0:
        # allow selecting the preview row index
        max_idx = len(df) - 1
        preview_idx = st.number_input("Preview row index (0-based)", min_value=0, max_value=max_idx, value=0, step=1)
        # prepare data for preview using the selected row
        row = df.iloc[preview_idx]
        row_dict = {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}
        row_map = SafeDict(row_dict)
        try:
            preview_subject_raw = subject_template.format_map(row_map)
        except Exception:
            # fallback: safe formatting
            preview_subject_raw = subject_template
        try:
            preview_body_raw = body_template.format_map(row_map)
        except Exception:
            preview_body_raw = body_template
        # convert bold marks to HTML for preview
        preview_subject_html = convert_bold(preview_subject_raw)
        preview_body_html = convert_bold(preview_body_raw)

        st.markdown("**To:** " + (extract_email(str(row.get("Email", ""))) or "N/A"))
        st.markdown("**Subject (preview):**")
        st.markdown(preview_subject_html, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("**Body (preview):**")
        st.markdown(preview_body_html, unsafe_allow_html=True)
    else:
        # No data uploaded -> use placeholders for preview
        try:
            preview_subject_html = convert_bold(subject_template)
            preview_body_html = convert_bold(body_template)
            st.markdown("No recipients uploaded — previewing templates with placeholders:")
            st.markdown("**Subject (preview):**")
            st.markdown(preview_subject_html, unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("**Body (preview):**")
            st.markdown(preview_body_html, unsafe_allow_html=True)
        except Exception:
            st.info("Edit the subject/body to preview.")

# ========================================
# Send Emails (no label feature)
# ========================================
if send_clicked:
    if df is None:
        st.error("Please upload a CSV or Excel file with an `Email` column before sending.")
    else:
        sent_count = 0
        skipped = []
        errors = []

        with st.spinner("📨 Sending emails..."):
            for idx, row in df.iterrows():
                raw_email_field = str(row.get("Email", "")).strip()
                to_addr = extract_email(raw_email_field)
                if not to_addr:
                    skipped.append(raw_email_field)
                    continue

                # prepare mapping for safe formatting
                row_map = SafeDict({k: ("" if pd.isna(v) else str(v)) for k, v in row.items()})

                try:
                    # safe format subject & body: missing placeholders become empty strings
                    try:
                        subject = subject_template.format_map(row_map)
                    except Exception:
                        subject = subject_template

                    try:
                        body_text = body_template.format_map(row_map)
                    except Exception:
                        body_text = body_template

                    html_body = convert_bold(body_text)

                    # Build and send HTML email (only bold allowed via convert_bold)
                    message = MIMEText(html_body, "html")
                    message["to"] = to_addr
                    message["subject"] = subject
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                    msg_body = {"raw": raw}

                    service.users().messages().send(userId="me", body=msg_body).execute()

                    sent_count += 1
                    time.sleep(delay)
                except HttpError as he:
                    # Gmail API errors
                    errors.append((to_addr, f"HttpError: {he}"))
                except Exception as e:
                    errors.append((to_addr, str(e)))

        # Summary
        st.success(f"✅ Successfully sent {sent_count} emails.")
        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid/missing emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to send {len(errors)} emails. Examples: {errors[:5]}")
