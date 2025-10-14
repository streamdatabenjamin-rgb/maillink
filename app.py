import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime

st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool")

# OAuth + Gmail setup (reuse your existing logic)
# ------------------------------------------------
# [KEEP all your Gmail API setup, flow, convert_bold(), etc. UNCHANGED here]

# After you finish loading creds, service, and dataframe creation,
# continue below ↓

st.header("📤 Upload Recipient List")
uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith("csv") else pd.read_excel(uploaded_file)
    st.dataframe(df.head())
    st.info("Edit/remove unwanted rows below:")
    df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    st.header("✍️ Compose Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area(
        "Body",
        """Dear {Name},

Welcome to our **Mail Merge App** demo.

Thanks,  
**Your Company**""",
        height=250,
    )

    st.header("🏷️ Options")
    label_name = st.text_input("Label", "Mail Merge Sent")
    delay = st.slider("Delay between mails (seconds)", 20, 75, 20)
    send_mode = st.radio("Mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

    # Store parameters and switch to progress page
    if st.button("🚀 Send Emails / Save Drafts"):
        st.session_state["mailmerge_params"] = {
            "subject_template": subject_template,
            "body_template": body_template,
            "label_name": label_name,
            "delay": delay,
            "send_mode": send_mode,
            "timestamp": datetime.now().strftime("%Y%m%d_%H%M%S"),
        }

        tmp_path = os.path.join("/tmp", f"active_mailmerge_{st.session_state['mailmerge_params']['timestamp']}.csv")
        df.to_csv(tmp_path, index=False)
        st.session_state["active_csv"] = tmp_path

        st.switch_page("pages/progress.py")
