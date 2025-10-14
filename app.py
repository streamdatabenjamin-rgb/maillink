import streamlit as st
import pandas as pd
import time
import datetime
import base64
import random
import re
import json
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Mail Merge Pro", layout="wide")
st.title("📧 Gmail Mail Merge System")

# ========================================
# 🔒 True UI Lock (Blocks All Clicks, Scrolls, and Drags)
# ========================================
lock_style = """
<style>
#ui_lock_overlay {
    position: fixed;
    top: 0; left: 0;
    width: 100vw; height: 100vh;
    background: rgba(200, 200, 200, 0.5);
    z-index: 99999;
    display: flex;
    justify-content: center;
    align-items: center;
    font-size: 1.8rem;
    color: #000;
    font-weight: bold;
    user-select: none;
}
body.locked {
    overflow: hidden !important;
}
</style>
"""

lock_script = """
<script>
function lockUI() {
    if (!document.getElementById('ui_lock_overlay')) {
        const overlay = document.createElement('div');
        overlay.id = 'ui_lock_overlay';
        overlay.innerText = '🔒 Processing... Please wait and do not click or drag.';
        document.body.appendChild(overlay);
        document.body.classList.add('locked');
    }
}
function unlockUI() {
    const overlay = document.getElementById('ui_lock_overlay');
    if (overlay) {
        overlay.remove();
        document.body.classList.remove('locked');
    }
}
</script>
"""
st.markdown(lock_style + lock_script, unsafe_allow_html=True)

def lock_ui():
    st.session_state["ui_locked"] = True
    st.markdown("<script>lockUI();</script>", unsafe_allow_html=True)

def unlock_ui():
    st.session_state["ui_locked"] = False
    st.markdown("<script>unlockUI();</script>", unsafe_allow_html=True)

if st.session_state.get("ui_locked", False):
    lock_ui()

# ========================================
# File Upload and Data Handling
# ========================================
uploaded_file = st.file_uploader("📂 Upload your CSV file", type=["csv"])

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)

    # ========================================
    # 🧹 Manual Delete Setup (unsubscribe removal)
    # ========================================
    st.markdown("### 🧹 Review & Clean Your Recipient List")
    st.info("Tick 'Remove' next to any unsubscribed contact to exclude them before sending.")

    if "Remove" not in df.columns:
        df.insert(0, "Remove", False)

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        key="recipient_editor_inline"
    )

    # Filter out removed rows
    cleaned_df = edited_df[edited_df["Remove"] != True].copy()
    removed_count = len(edited_df) - len(cleaned_df)

    if removed_count > 0:
        st.warning(f"🗑️ {removed_count} recipient(s) marked for removal and excluded.")
    else:
        st.success("✅ All recipients are included.")

    df = cleaned_df

    # ========================================
    # Compose Mail Section (unchanged)
    # ========================================
    st.markdown("### ✉️ Compose Your Email")
    subject = st.text_input("Subject")
    body = st.text_area("Body", height=200)

    mode = st.radio("Select Mode", ["New Email", "Follow-up (Reply)", "Save as Draft"])

    # Delay Section
    delay = st.number_input("Delay between emails (in seconds)", min_value=30, max_value=600, value=30, step=5)
    st.caption("⏳ Upload maximum of 70–80 rows for smooth run and to protect your Gmail quota.")

    # ETA Calculation
    if len(df) > 0:
        total_seconds = len(df) * delay
        finish_time = datetime.datetime.now() + datetime.timedelta(seconds=total_seconds)
        finish_str = finish_time.strftime("%I:%M %p")
        st.info(f"Estimated completion time for {len(df)} emails: **{finish_str}**")

    # ========================================
    # Preview before sending
    # ========================================
    st.markdown("### 👀 Final Preview Before Sending")
    st.dataframe(df, use_container_width=True)

    # ========================================
    # Send or Draft Emails
    # ========================================
    send_clicked = st.button("🚀 Send Emails / Save Drafts")

    if send_clicked:
        try:
            lock_ui()
            st.info("📤 Sending in progress... please wait and do not click anything.")
            time.sleep(1)

            for index, row in df.iterrows():
                # Simulate email sending (replace with Gmail API logic)
                time.sleep(delay)

            st.success("✅ Process completed successfully.")
            unlock_ui()

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            unlock_ui()

else:
    st.warning("📄 Please upload a CSV file to begin.")
