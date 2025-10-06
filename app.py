# app.py
import streamlit as st
import pandas as pd
import base64
import time
import re
import json
from datetime import datetime
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

# ============================
# Streamlit page setup
# ============================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge (HTML + Labeling + Delay + OAuth)")

# ============================
# Gmail API / OAuth config
# ============================
SCOPES = ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/gmail.labels", "https://www.googleapis.com/auth/gmail.modify"]

CLIENT_CONFIG = {
    "web": {
        "client_id": st.secrets["gmail"]["client_id"],
        "client_secret": st.secrets["gmail"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [st.secrets["gmail"]["redirect_uri"]],
    }
}

# ============================
# Helpers
# ============================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def create_message(to, subject, body, is_html=True):
    """Create email message (supports HTML)."""
    mime = MIMEText(body, "html" if is_html else "plain")
    mime["to"] = to
    mime["subject"] = subject
    raw = base64.urlsafe_b64encode(mime.as_bytes()).decode("utf-8")
    return {"raw": raw}

def get_or_create_label(service, label_name="MailMerge"):
    """Get a label ID by name or create it if not present."""
    try:
        labels_resp = service.users().labels().list(userId="me").execute()
        labels = labels_resp.get("labels", [])
        for lbl in labels:
            if lbl.get("name", "").lower() == label_name.lower():
                return lbl["id"]
        # Create label
        label_obj = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        new_label = service.users().labels().create(userId="me", body=label_obj).execute()
        return new_label["id"]
    except HttpError as e:
        st.warning(f"Could not get/create label '{label_name}': {e}")
        return None

def send_email_and_label(service, to, subject, body, label_id=None):
    """Send message and apply label if provided."""
    message = create_message(to, subject, body, is_html=True)
    try:
        sent = service.users().messages().send(userId="me", body=message).execute()
        if label_id:
            # Modify the sent message to add the label
            try:
                service.users().messages().modify(userId="me", id=sent["id"], body={"addLabelIds": [label_id]}).execute()
            except HttpError as e:
                st.warning(f"Could not label message {sent.get('id')}: {e}")
        return sent
    except HttpError as e:
        raise

# ============================
# OAuth Flow and Credentials
# ============================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

# If we already stored creds in session, construct Credentials
if st.session_state["creds"]:
    try:
        creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
        # Refresh if needed
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            st.session_state["creds"] = creds.to_json()
    except Exception as e:
        st.error(f"Error loading credentials: {e}")
        st.session_state["creds"] = None
        st.stop()
else:
    # Check for return 'code' from OAuth redirect
    query_params = st.experimental_get_query_params()
    code = query_params.get("code")
    if code:
        try:
            flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
            flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
            flow.fetch_token(code=code[0])
            creds = flow.credentials
            st.session_state["creds"] = creds.to_json()
            # Clean URL params and rerun so the app continues authenticated
            st.experimental_set_query_params()
            st.experimental_rerun()
        except Exception as e:
            st.error(f"OAuth token exchange failed: {e}")
            st.stop()
    else:
        # Start auth flow
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send and label emails from your Gmail account.")
        st.stop()

# Build the Gmail service client
try:
    creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
except Exception as e:
    st.error(f"Could not build Gmail service: {e}")
    st.stop()

# ============================
# UI: Upload recipients
# ============================
st.header("📤 Upload Recipient List")
uploaded_file = st.file_uploader("Upload CSV or Excel file (must contain an 'Email' column)", type=["csv", "xlsx"])

df = None
if uploaded_file:
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.success("File uploaded")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Error reading file: {e}")

# ============================
# UI: Email compose & options
# ============================
if df is not None:
    st.header("✍️ Compose Your Email (HTML supported)")
    subject_template = st.text_input("Subject", "Hello {Name}")

    default_body = """<p><b>Dear {Name},</b></p>
<p>We’re excited to share our latest updates. Visit <a href="https://phoenixxit.com" target="_blank">our website</a>.</p>
<p>Best regards,<br><b>Team Phoenixx IT</b></p>"""

    body_template = st.text_area("Body (HTML)", default_body, height=260)
    st.markdown("""
    **Formatting tips:** use `<b>`, `<i>`, `<u>`, `<a href="...">link</a>`, `<ul>/<li>`, inline styles like `<span style="color:#e60000">text</span>`.
    """)

    st.markdown("### 📄 Preview (sample data)")
    try:
        sample_context = {col: f"Sample_{col}" for col in df.columns}
        preview_html = body_template.format(**sample_context)
        st.markdown(preview_html, unsafe_allow_html=True)
    except Exception:
        st.warning("Some placeholders may not match your CSV column names. Preview shown without replacements.")
        st.markdown(body_template, unsafe_allow_html=True)

    st.header("⏱️ Sending Options")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)
    st.write("Tip: Use a small delay to reduce hitting rate limits.")

    st.header("🏷️ Label Sent Emails")
    default_label = f"MailMerge_{datetime.now().strftime('%Y%m%d')}"
    label_name = st.text_input("Label name to apply to sent emails", default_label)

    # Button to send
    if st.button("🚀 Send Emails"):
        if "Email" not in df.columns and "email" not in [c.lower() for c in df.columns]:
            st.error("CSV must contain an 'Email' column (case-insensitive).")
        else:
            sent_count = 0
            skipped = []
            errors = []

            # Prepare label id (create if needed)
            label_id = get_or_create_label(service, label_name)

            for idx, row in df.iterrows():
                # try column keys case-insensitively for Email
                if "Email" in df.columns:
                    raw_email = row.get("Email", "")
                else:
                    # fallback: find first column that case-insensitive matches 'email'
                    email_col = next((c for c in df.columns if c.lower() == "email"), None)
                    raw_email = row.get(email_col, "") if email_col else ""

                to_addr = extract_email(str(raw_email).strip())
                if not to_addr:
                    skipped.append(str(raw_email))
                    continue

                # Render subject and body safely; fallback on KeyError
                try:
                    subject = subject_template.format(**row)
                except Exception:
                    subject = subject_template
                try:
                    body = body_template.format(**row)
                except Exception:
                    # If formatting fails (missing keys), try with simple replacements or leave as-is
                    try:
                        # attempt using column names as strings
                        simple_ctx = {col: str(row.get(col, "")) for col in df.columns}
                        body = body_template.format(**simple_ctx)
                    except Exception:
                        body = body_template

                try:
                    send_email_and_label(service, to_addr, subject, body, label_id=label_id)
                    sent_count += 1
                    st.info(f"Sent to: {to_addr}")
                    time.sleep(delay)
                except Exception as e:
                    errors.append((to_addr, str(e)))
                    st.error(f"Failed to send to {to_addr}: {e}")

            st.success(f"✅ Sent: {sent_count}")
            if skipped:
                st.warning(f"⚠️ Skipped {len(skipped)} invalid addresses: {skipped[:10]}{'... (more)' if len(skipped)>10 else ''}")
            if errors:
                st.error(f"❌ Failed to send {len(errors)} messages. See first errors: {errors[:5]}")

    # Optional: show labels available (for user awareness)
    try:
        lbls = service.users().labels().list(userId="me").execute().get("labels", [])
        lbl_names = [l["name"] for l in lbls]
        st.markdown("**Existing labels in your Gmail:**")
        st.write(", ".join(lbl_names))
    except Exception:
        pass

else:
    st.info("Upload a CSV or Excel file to get started.")
