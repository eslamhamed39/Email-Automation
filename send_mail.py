
import streamlit as st
import pandas as pd
import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from streamlit_quill import st_quill
import time

# --- Streamlit UI ---
st.set_page_config(page_title="Email Automation System", layout="wide")
st.title("📧 Geomakani Email Automation")

# --- Sidebar for credentials ---
st.sidebar.header("📌 Email Server Configuration")
smtp_server = "smtp.hostinger.com"
smtp_port = 465
imap_server = "imap.hostinger.com"
imap_port = 993
sender_email = st.sidebar.text_input("Sender Email")
password = st.sidebar.text_input("Password", type="password")
name = st.sidebar.text_input("Name")
title = st.sidebar.text_input("Title")
phone = st.sidebar.text_input("Phone")
email = st.sidebar.text_input("Email")

# --- Main Interface ---
st.header("🚀 Email Content and List Upload")
email_subject = st.text_input("Email Subject", "Potential Collaboration Opportunity with Geomakani")
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])

st.subheader("📝 Compose Email (Rich Text)")
email_body_html = st_quill(html=True, key="editor")

# --- Send Email Logic ---
if st.button("Send Emails"):
    if uploaded_file is not None and sender_email and password and email_body_html:
        data = pd.read_excel(uploaded_file)

        # SMTP Connection
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(sender_email, password)

        success_count, fail_count = 0, 0
        progress_bar = st.progress(0)
        status_text = st.empty()

        for index, row in data.iterrows():
            company_name = row['Company_Name_English']
            recipient_email = row['Email']

            # Create Email
            msg = MIMEMultipart('alternative')
            msg['Subject'] = email_subject
            msg['From'] = sender_email
            msg['To'] = recipient_email
            # msg['Bcc'] = sender_email

            # Personalize the message
            signature_html = '''<br>
            <table style="width:100%; max-width:600px; font-family:Arial, sans-serif;">
                <tr>
                    <td style="vertical-align: center;">
                        <img src="https://i.ibb.co/Hp7L6yg3/image001.png" alt="Geomakani Logo" width="152">
                    </td>
                    <td style="vertical-align: top;">
                        <p style="margin: 0; color: #e67300; font-weight: bold; font-size:20px;">{name}</p>
                        <p style="margin: 0; color: #e67300;font-size:17px;">{title}</p>
                        <br>
                        <p style="margin: 0;"><strong>Mobile:</strong>{phone}</p>
                        <p style="margin: 0;"><strong>Email:</strong> <a href="mailto:{email}">{email}</a></p>
                        <p style="margin: 0;"><strong>Address:</strong> 39A El Kinda Compound, El Kattamya,<br>Cairo, Egypt</p>
                        <a href="https://www.facebook.com/geomakani/?locale=EN"><img src="https://i.ibb.co/KxQVCzQD/image002.png" alt="Facebook" style="margin-right: 8px;"></a>
                        <a href="https://www.linkedin.com/company/geomakani-official-page/posts/?feedView=all"><img src="https://i.ibb.co/0jFWT0PX/image003.png" alt="LinkedIn"></a>
                    </td>
                </tr>
            </table><br>
            <img src="https://i.ibb.co/3yKwkmqm/image004.jpg" width="600" alt="Navigate the Future"><br><br>
            <div style="font-size: 12px; color: gray; max-width: 600px;">
                The content of this email is confidential and intended for the recipient specified in message only. It is strictly forbidden to share any part of this message with any third party, without a written consent of the sender. If you received this message by mistake, please reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future. Please do not print this email unless it is necessary. Every unprinted email helps the environment.
            </div>'''.format(name=name, title=title, phone=phone, email=email)
            personalized_html = email_body_html.replace("{company_name}", company_name) + signature_html
            msg.attach(MIMEText(personalized_html, 'html'))
            try:
                server.sendmail(sender_email, [recipient_email, sender_email], msg.as_string())
                success_count += 1
                status_text.success(f"✅ Email sent to {company_name}")
            except Exception as e:
                fail_count += 1
                status_text.error(f"❌ Failed to send to {recipient_email}: {e}")
                continue
            # Save to Sent Folder
            try:
                imap = imaplib.IMAP4_SSL(imap_server, imap_port)
                imap.login(sender_email, password)
                imap.select('"INBOX.Sent"')
                imap.append('"INBOX.Sent"', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                imap.logout()
            except Exception as e:
                st.warning(f"⚠️ Failed to save email to Sent folder for {recipient_email}: {e}")

            progress_bar.progress((index + 1) / len(data))
            time.sleep(2)

        server.quit()
        st.success(f"🎉 Emails sent: {success_count}, Failures: {fail_count}")
    else:
        st.error("⚠️ Please complete all fields and compose the email body.")


