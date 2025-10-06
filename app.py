# app.py
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
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from typing import Optional, Dict, Any
from collections import defaultdict

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(page_title="Gmail Mail Merge — Production", layout="wide")
st.title("📧 Gmail Mail Merge — Production")

# ---------------------------
# Settings / Constants
# ---------------------------
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

# ---------------------------
# Utility helpers
# ---------------------------
def extract_email(value: Optional[str]) -> Optional[str]:
    """Extract the first valid email from a string, or None if not found."""
    if value is None:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def create_message(to: str, subject: str, body: str) -> Dict[str, str]:
    """Return dict with 'raw' message encoded for Gmail API."""
    message = MIMEText(body)
    message["to"] = to
    message["subject"] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": raw}

def exponential_backoff_send(service, body, max_retries=4, base_delay=1.0):
    """Send a Gmail message with simple exponential backoff for transient errors."""
    attempt = 0
    while True:
        try:
            return service.users().messages().send(userId="me", body=body).execute()
        except HttpError as e:
            attempt += 1
            # treat 5xx and rate-limit-like errors as retryable
            status = getattr(e, "status_code", None)
            # older googleapiclient might not set status_code; fall back to parsing
            reason = ""
            try:
                reason = e.error_details if hasattr(e, "error_details") else str(e)
            except Exception:
                reason = str(e)
            if attempt > max_retries:
                raise
            sleep_time = base_delay * (2 ** (attempt - 1))
            st.info(f"Transient error (attempt {attempt}/{max_retries}). Retrying in {sleep_time:.1f}s...")
            time.sleep(sleep_time)
        except Exception:
            # Non-HttpError (e.g., network). Allow retries as well.
            attempt += 1
            if attempt > max_retries:
                raise
            sleep_time = base_delay * (2 ** (attempt - 1))
            st.info(f"Error sending (attempt {attempt}/{max_retries}). Retrying in {sleep_time:.1f}s...")
            time.sleep(sleep_time)

class SafeDict(dict):
    """Used to safely format templates: missing keys return an empty string instead of KeyError."""
    def __missing__(self, key):
        return ""

def safe_format(template: str, row: Dict[str, Any]) -> str:
    """Safely format template using SafeDict, converting all keys to strings."""
    # Convert keys to strings to support integer column names etc.
    mapping = {str(k): ("" if pd.isna(v) else v) for k, v in row.items()}
    return template.format_map(SafeDict(mapping))

# ---------------------------
# OAuth / Credentials handling
# ---------------------------
if "creds" not in st.session_state:
    st.session_state.creds = None  # JSON string of credentials

# CLIENT_CONFIG must be present in st.secrets under 'gmail'
if "gmail" not in st.secrets:
    st.error("Please set Gmail OAuth credentials in Streamlit secrets (gmail: client_id, client_secret, redirect_uri).")
    st.stop()

CLIENT_CONFIG = {
    "web": {
        "client_id": st.secrets["gmail"]["client_id"],
        "client_secret": st.secrets["gmail"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [st.secrets["gmail"]["redirect_uri"]],
    }
}

def ensure_credentials() -> Optional[Credentials]:
    """Ensure we have valid credentials. If not authorized, present auth link and stop."""
    # If we already have credentials stored
    if st.session_state.creds:
        try:
            creds = Credentials.from_authorized_user_info(json.loads(st.session_state.creds), SCOPES)
        except Exception:
            # If stored creds are corrupted, reset
            st.session_state.creds = None
            return None

        # Refresh if expired and refresh token available
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state.creds = creds.to_json()
            except Exception as e:
                # If refresh fails, clear creds and require re-auth
                st.warning("Session refresh failed, please authorize again.")
                st.session_state.creds = None
                return None
        return creds

    # If code present in query params -> exchange for token
    code = st.experimental_get_query_params().get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        try:
            flow.fetch_token(code=code[0])
        except Exception as e:
            st.error(f"Failed to fetch token: {e}")
            st.stop()
        creds = flow.credentials
        st.session_state.creds = creds.to_json()
        # Redirect to clean URL (remove code param) by re-running
        st.experimental_set_query_params()
        st.experimental_rerun()
        return None

    # No credentials and no code -> show auth link
    flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
    flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
    auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
    st.markdown(
        "### 🔑 Authorize Gmail"
    )
    st.markdown(
        f"Please [authorize this app]({auth_url}) to send emails on your behalf. After authorizing you will be redirected back; the app will finish setup automatically."
    )
    st.stop()
    return None

creds = ensure_credentials()
if not creds:
    st.stop()

# Build Gmail API client
try:
    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
except Exception as e:
    st.error(f"Failed to build Gmail service: {e}")
    st.stop()

# ---------------------------
# Main UI: upload + template
# ---------------------------
st.header("📤 1) Upload Recipients")
uploaded_file = st.file_uploader("Upload CSV or Excel file (first column containing emails or a column named 'Email')", type=["csv", "xlsx", "xls"])

df = None
if uploaded_file:
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Could not read uploaded file: {e}")
        st.stop()

if df is not None:
    st.markdown("**Preview of uploaded data**")
    st.dataframe(df.head())

    # Ensure Email column normalized: try to find a column with emails if Email not explicitly present
    if "Email" not in df.columns:
        # attempt auto-detect email in any column
        detected = None
        for col in df.columns:
            sample_vals = df[col].astype(str).head(20).tolist()
            if any(EMAIL_REGEX.search(v) for v in sample_vals):
                detected = col
                break
        if detected:
            st.warning(f"No 'Email' column found. Detected '{detected}' likely contains emails. Renaming to 'Email'.")
            df = df.rename(columns={detected: "Email"})
        else:
            st.warning("No column named 'Email' detected and no email-like column found. You must include an email column.")
            st.stop()

    st.header("✍️ 2) Compose Email Template")
    subject_template = st.text_input("Subject (use {ColumnName} placeholders)", value="Hello {Name}")
    body_template = st.text_area("Body (use {ColumnName} placeholders)", value="Dear {Name},\n\nThis is a test mail.\n\nRegards,\nYour Company", height=200)

    st.header("⚙️ 3) Sending Options")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)
    dry_run = st.checkbox("Preview / Dry Run (don't actually send emails)", value=True)
    show_preview = st.checkbox("Show preview of first 5 personalized emails", value=True)
    max_batch = st.number_input("Max recipients to process in this run (0 = all)", min_value=0, value=0, step=1, help="Use this to batch large lists.")
    confirm_send = st.checkbox("I confirm I want to send these emails (enable to activate Send button)", value=False)

    # Rate limit reminder
    st.info("⚠️ Reminder: Gmail sending limits apply (e.g., 500/day for regular Gmail accounts). If you're sending large volumes, use a Google Workspace account or a dedicated sending service.")

    # Preview block
    if show_preview:
        st.subheader("Preview (first 5 rows)")
        preview_rows = df.head(5).to_dict(orient="records")
        for r in preview_rows:
            to_field = extract_email(r.get("Email", ""))
            subj = safe_format(subject_template, r)
            body = safe_format(body_template, r)
            st.markdown(f"**To:** {to_field}  \n**Subject:** {subj}  \n**Body:**\n```\n{body}\n```")

    # Send / Run
    st.header("🚀 4) Send Emails")
    if st.button("Send Emails", disabled=not confirm_send):
        # Determine recipients slice
        total_recipients = len(df)
        if max_batch and max_batch > 0:
            df_to_process = df.head(max_batch)
        else:
            df_to_process = df

        if total_recipients > 1000:
            st.warning("Large recipient lists can take a while and may trigger Gmail limits. Consider batching or using a dedicated transactional email provider for high-volume sends.")

        if dry_run:
            st.success("Dry run enabled — emails will not be sent. Showing what WOULD be sent.")
        # Prepare logging
        results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        # iterate rows
        rows = df_to_process.to_dict(orient="records")
        for i, row in enumerate(rows, start=1):
            # Update progress/status early so user sees activity
            progress_bar.progress(int((i-1)/max(1, len(rows)) * 100))
            to_raw = row.get("Email", "")
            to_addr = extract_email(to_raw)
            if not to_addr:
                results.append({"Email": to_raw, "Status": "Skipped", "Error": "Invalid or missing email"})
                continue

            # Safe format subject/body — missing placeholder keys become empty strings
            try:
                subject = safe_format(subject_template, row)
                body = safe_format(body_template, row)
            except Exception as e:
                results.append({"Email": to_addr, "Status": "Skipped", "Error": f"Template error: {e}"})
                continue

            status_text.text(f"Processing {i}/{len(rows)} — sending to {to_addr} ...")

            if dry_run:
                # In dry run just record that we would have sent
                results.append({"Email": to_addr, "Status": "DryRun", "Error": "" , "Subject": subject, "Body": body})
                # small delay so UI is usable
                time.sleep(min(0.05, delay))
            else:
                # Attempt to send with backoff handling
                try:
                    message_body = create_message(to_addr, subject, body)
                    exponential_backoff_send(service, message_body, max_retries=4, base_delay=1.0)
                    results.append({"Email": to_addr, "Status": "Sent", "Error": "", "Subject": subject})
                    # respectful delay between sends
                    time.sleep(delay)
                except HttpError as he:
                    results.append({"Email": to_addr, "Status": "Failed", "Error": f"HttpError: {he}"})
                except Exception as e:
                    results.append({"Email": to_addr, "Status": "Failed", "Error": str(e)})

            # update progress bar
            progress_bar.progress(int(i / max(1, len(rows)) * 100))

        status_text.text("Completed.")
        progress_bar.progress(100)

        # Prepare results dataframe and actions
        results_df = pd.DataFrame(results)
        st.subheader("Run Results")
        st.dataframe(results_df)

        # Summary counts
        counts = results_df["Status"].value_counts().to_dict()
        summary_lines = [f"{k}: {v}" for k, v in counts.items()]
        st.write("**Summary:** " + " | ".join(summary_lines))

        # Download log
        csv = results_df.to_csv(index=False).encode("utf-8")
        st.download_button("📥 Download Send Log (CSV)", csv, "mail_merge_log.csv", "text/csv")

        # Optional: Save credentials note
        st.info("Your OAuth credentials remain stored in this session only (st.session_state). For persistent deployments, use a secure server-side secrets storage and rotate refresh tokens periodically.")

else:
    st.info("Upload a recipient file to begin (CSV or Excel).")
