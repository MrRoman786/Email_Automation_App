# main.py
import streamlit as st
import pandas as pd
import smtplib
import re
import email.mime.text
import email.mime.multipart
import io
import time
from typing import Dict, List, Tuple, Optional

def validate_email(email: str) -> bool:
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def extract_placeholders(template: str) -> List[str]:
    placeholders = re.findall(r'\{\{(\w+)\}\}', template)
    return list(set(placeholders))

def replace_placeholders(template: str, data: Dict[str, str]) -> str:
    result = template
    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"
        result = result.replace(placeholder, str(value))
    return result

def validate_csv_columns(df: pd.DataFrame) -> Tuple[bool, str]:
    required_columns = ['name', 'email']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        return False, f"Missing required columns: {', '.join(missing_columns)}"

    for col in required_columns:
        if df[col].isnull().any() or (df[col] == '').any():
            return False, f"Column '{col}' contains empty values"

    invalid_emails = []
    for idx, email in enumerate(df['email']):
        if not validate_email(email):
            invalid_emails.append(f"Row {idx + 1}: {email}")

    if invalid_emails:
        return False, f"Invalid email addresses found:\n" + "\n".join(invalid_emails[:5])

    return True, "CSV validation successful"

def send_email(smtp_server, sender_email: str, recipient_email: str, subject: str, body: str, is_html: bool = False) -> Tuple[bool, str]:
    try:
        msg = email.mime.multipart.MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        if is_html:
            msg.attach(email.mime.text.MIMEText(body, 'html'))
        else:
            msg.attach(email.mime.text.MIMEText(body, 'plain'))

        smtp_server.send_message(msg)
        return True, "Email sent successfully"
    except Exception as e:
        return False, f"Failed to send email: {str(e)}"

def main():
    st.set_page_config(page_title="Email Automation Tool", page_icon="📧", layout="wide")
    st.title("📧 Email Automation Tool")
    st.markdown("Upload your template and CSV file to send personalized emails via Gmail SMTP")

    if 'email_results' not in st.session_state:
        st.session_state.email_results = []
    if 'preview_data' not in st.session_state:
        st.session_state.preview_data = None

    with st.sidebar:
        st.header("Gmail Configuration")
        sender_email = st.text_input("Gmail Address", placeholder="your.email@gmail.com")
        sender_password = st.text_input("App Password", type="password", help="Use Gmail App Password")

        if st.button("Test Connection"):
            if sender_email and sender_password:
                if not validate_email(sender_email):
                    st.error("Invalid email address format")
                else:
                    try:
                        with st.spinner("Testing connection..."):
                            server = smtplib.SMTP('smtp.gmail.com', 587)
                            server.starttls()
                            server.login(sender_email, sender_password)
                            server.quit()
                        st.success("✅ Connection successful!")
                    except Exception as e:
                        st.error(f"❌ Connection failed: {str(e)}")
            else:
                st.warning("Please enter both email and password")

    col1, col2 = st.columns(2)

    with col1:
        st.header("📄 Upload Template")
        template_file = st.file_uploader("Choose template file", type=['txt', 'html'])
        if template_file:
            try:
                template_content = template_file.read().decode('utf-8')
                st.success(f"✅ Template loaded: {template_file.name}")
                with st.expander("Preview Template"):
                    st.code(template_content, language='html' if template_file.name.endswith('.html') else 'text')
                placeholders = extract_placeholders(template_content)
                if placeholders:
                    st.info(f"Placeholders found: {', '.join(placeholders)}")
                else:
                    st.warning("No placeholders found in template")
            except Exception as e:
                st.error(f"Error reading template file: {str(e)}")
                template_content = None
                placeholders = []
        else:
            template_content = None
            placeholders = []

    with col2:
        st.header("📊 Upload CSV Data")
        csv_file = st.file_uploader("Choose CSV file", type=['csv'])
        if csv_file:
            try:
                df = pd.read_csv(csv_file)
                st.success(f"✅ CSV loaded: {csv_file.name} ({len(df)} rows)")
                is_valid, validation_message = validate_csv_columns(df)
                if is_valid:
                    st.success(validation_message)
                    with st.expander("Preview Data"):
                        st.dataframe(df.head())
                    st.info(f"Total recipients: {len(df)}")
                else:
                    st.error(validation_message)
                    df = None
            except Exception as e:
                st.error(f"Error reading CSV file: {str(e)}")
                df = None
        else:
            df = None

    st.header("✉️ Email Configuration")
    email_subject = st.text_input("Email Subject", placeholder="Enter email subject")

    if template_content and df is not None and len(df) > 0:
        st.header("👀 Email Preview")
        if st.button("Generate Preview (First Row)"):
            try:
                first_row = df.iloc[0].to_dict()
                preview_content = replace_placeholders(template_content, first_row)
                st.session_state.preview_data = {
                    'recipient': first_row,
                    'content': preview_content,
                    'is_html': template_file.name.endswith('.html') if template_file else False
                }
            except Exception as e:
                st.error(f"Error generating preview: {str(e)}")

        if st.session_state.preview_data:
            st.subheader(f"Preview for: {st.session_state.preview_data['recipient']['name']} ({st.session_state.preview_data['recipient']['email']})")
            if st.session_state.preview_data['is_html']:
                st.components.v1.html(st.session_state.preview_data['content'], height=400, scrolling=True)
            else:
                st.text_area("Email Content", st.session_state.preview_data['content'], height=300, disabled=True)

    if template_content and df is not None and sender_email and sender_password and email_subject:
        st.header("🚀 Send Emails")
        if st.button("Send All Emails", type="primary"):
            st.session_state.email_results = []
            if not validate_email(sender_email):
                st.error("Invalid sender email address")
                return
            try:
                with st.spinner("Connecting to Gmail SMTP..."):
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(sender_email, sender_password)

                progress_bar = st.progress(0)
                status_text = st.empty()
                is_html = template_file.name.endswith('.html') if template_file else False

                for idx, row in df.iterrows():
                    try:
                        progress = (idx + 1) / len(df)
                        progress_bar.progress(progress)
                        status_text.text(f"Sending email {idx + 1} of {len(df)} to {row['email']}")
                        personalized_content = replace_placeholders(template_content, row.to_dict())
                        success, message = send_email(server, sender_email, row['email'], email_subject, personalized_content, is_html)
                        st.session_state.email_results.append({
                            'recipient': row['email'],
                            'name': row['name'],
                            'success': success,
                            'message': message
                        })
                        time.sleep(0.5)
                    except Exception as e:
                        st.session_state.email_results.append({
                            'recipient': row['email'],
                            'name': row['name'],
                            'success': False,
                            'message': f"Error: {str(e)}"
                        })

                server.quit()
                status_text.text("✅ Email sending completed!")
                progress_bar.progress(1.0)
            except Exception as e:
                st.error(f"SMTP connection error: {str(e)}")

    if st.session_state.email_results:
        st.header("📊 Sending Results")
        successful = [r for r in st.session_state.email_results if r['success']]
        failed = [r for r in st.session_state.email_results if not r['success']]

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Sent", len(st.session_state.email_results))
        with col2:
            st.metric("Successful", len(successful))
        with col3:
            st.metric("Failed", len(failed))

        tab1, tab2 = st.tabs(["✅ Successful", "❌ Failed"])
        with tab1:
            if successful:
                success_df = pd.DataFrame(successful)
                st.dataframe(success_df[['name', 'recipient', 'message']])
            else:
                st.info("No successful emails yet")
        with tab2:
            if failed:
                failed_df = pd.DataFrame(failed)
                st.dataframe(failed_df[['name', 'recipient', 'message']])
                if st.button("Retry Failed Emails"):
                    st.info("Retry functionality would be implemented here")
            else:
                st.success("No failed emails!")

        if st.button("Clear Results"):
            st.session_state.email_results = []
            st.session_state.preview_data = None
            st.rerun()

if __name__ == "__main__":
    main()
