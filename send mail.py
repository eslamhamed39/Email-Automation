




# import smtplib
# import imaplib
# import email
# import pandas as pd
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# import time

# # --- Email and Server Settings ---
# smtp_server = 'smtp.hostinger.com'
# smtp_port = 465
# imap_server = 'imap.hostinger.com'
# imap_port = 993
# sender_email = 'eslam.hamed@geomakani.com'
# password = 'H..e..240898'

# # --- Read the Excel file ---
# data = pd.read_excel('./Final_Company_Emails (Copy).xlsx')

# # --- SMTP Server Connection ---
# server = smtplib.SMTP_SSL(smtp_server, smtp_port)
# server.login(sender_email, password)

# # --- Loop through each contact ---
# for index, row in data.iterrows():
#     company_name = row['Company_Name_English']
#     recipient_email = row['Email']

#     # Create the Email
#     msg = MIMEMultipart('alternative')
#     msg['Subject'] = "Potential Collaboration Opportunity with Geomak based on my knowledge cutoff in April 2023ani"
#     msg['From'] = sender_email
#     msg['To'] = recipient_email
#     msg['Bcc'] = sender_email 

#     # Email body (HTML with personalization)
#     html = f"""\
#     <html>
#         <body>
            # <p>Dear {company_name},<br><br>
            # Greetings,<br><br>
            # My name is Fatma, Senior Business Developer and representative of the Geospatial Research Center "Geomakani." It is a pleasure to reach out to you to express our sincere interest in establishing a fruitful partnership with your esteemed company. We believe that through collaboration, we can achieve our shared goals and contribute to future successes.<br><br>
            # Geomakani is a leading company specializing in providing innovative technological solutions tailored to the needs of both public and private sectors. With extensive experience and a proven track record, we have consistently delivered effective solutions that enhanced efficiency, increased productivity, and strengthened our clients’ competitiveness.<br><br>
            # Our services span a wide range of fields, including:<br>
            # <ul>
            #     <li>Development of web and artificial intelligence applications</li>
            #     <li>Digital mapping and land surveying</li>
            #     <li>Integrated solutions in public safety and smart city development</li>
            #     <li>Remote sensing applications and geographic information systems (GIS)</li>
            #     <li>Digital elevation models, land management systems, and encroachment detection systems (horizontal and vertical)</li>
            # </ul>
            # Additionally, we possess significant expertise in establishing and managing emergency operations centers, fleet management systems and satellite image processing. We also offer specialized training programs and organize geospatial events and conferences.<br><br>
            # We would be delighted to discuss potential areas of collaboration with you at your earliest convenience. We are prepared to provide a detailed presentation tailored to your specific needs.<br><br>
            # Thank you for your time and consideration. We look forward to the opportunity to work together and wish you continued success and prosperity.<br><br>
            # Yours sincerely,
            # <br>
            # <table style="width:100%; max-width:600px; font-family:Arial, sans-serif;">
            #     <tr>
            #         <td style="vertical-align: center;">
            #             <img src="https://i.ibb.co/Hp7L6yg3/image001.png" alt="Geomakani Logo" width="152">
            #         </td>
            #         <td style="vertical-align: top;">
            #             <p style="margin: 0; color: #e67300; font-weight: bold;">Fatma M. Ibrahim</p>
            #             <p style="margin: 0; color: #e67300;">Senior Business Developer</p>
            #             <br>
            #             <p style="margin: 0;"><strong>Mobile:</strong> 002 01206251425</p>
            #             <p style="margin: 0;"><strong>Email:</strong> <a href="mailto:Fatma.mohamed@geomakani.com">Fatma.mohamed@geomakani.com</a></p>
            #             <p style="margin: 0;"><strong>Address:</strong> 39A El Kinda Compound, El Kattamya,<br>Cairo, Egypt</p>
            #             <a href="https://www.facebook.com/geomakani/?locale=EN"><img src="https://i.ibb.co/KxQVCzQD/image002.png" alt="Facebook" style="margin-right: 8px;"></a>
            #             <a href="https://www.linkedin.com/company/geomakani-official-page/posts/?feedView=all"><img src="https://i.ibb.co/0jFWT0PX/image003.png" alt="LinkedIn"></a>
            #         </td>
            #     </tr>
            # </table>
            # <br>
            # <img src="https://i.ibb.co/3yKwkmqm/image004.jpg" width="600" alt="Navigate the Future"><br><br>
            # <div style="font-size: 12px; color: gray; max-width: 600px;">
            #     The content of this email is confidential and intended for the recipient specified in message only. It is strictly forbidden to share any part of this message with any third party, without a written consent of the sender. If you received this message by mistake, please reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future. Please do not print this email unless it is necessary. Every unprinted email helps the environment.
            # </div>
#         </body>
#     </html>
#     """
#     msg.attach(MIMEText(html, 'html'))

#     # Send Email
#     try:
#         server.sendmail(sender_email, [recipient_email, sender_email], msg.as_string())
#         print(f"Email sent to {company_name} ({recipient_email})")
#     except Exception as e:
#         print(f"Failed to send email to {recipient_email}: {str(e)}")
#         continue

#     # Save to Sent folder using IMAP
#     try:
#         imap = imaplib.IMAP4_SSL(imap_server, imap_port)
#         imap.login(sender_email, password)
#         imap.select('"INBOX.Sent"')  # استخدام المجلد الصحيح
#         imap.append('"INBOX.Sent"', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
#         print(f"Email saved to Sent folder for {recipient_email}")
#         imap.logout()
#     except Exception as e:
#         print(f"Failed to save email to Sent folder for {recipient_email}: {str(e)}")

#     time.sleep(3)  # انتظر 3 ثواني لتجنب الحظر

# # Close SMTP connection
# server.quit()
# print("✅ All emails sent and saved to Sent folder successfully!")





# Install streamlit first by running: pip install streamlit
# import streamlit as st
# import pandas as pd
# import smtplib
# import imaplib
# import email
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# import time

# # --- Streamlit UI ---
# st.set_page_config(page_title="Email Automation System", layout="wide")
# st.title("📧 Geomakani Email Automation")

# # --- Sidebar for credentials ---
# st.sidebar.header("📌 Email Server Configuration")
# smtp_server = st.sidebar.text_input("SMTP Server", "smtp.hostinger.com")
# smtp_port = st.sidebar.number_input("SMTP Port", value=465)
# imap_server = st.sidebar.text_input("IMAP Server", "imap.hostinger.com")
# imap_port = st.sidebar.number_input("IMAP Port", value=993)
# sender_email = st.sidebar.text_input("Sender Email")
# password = st.sidebar.text_input("Password", type="password")

# # --- Main Interface ---
# st.header("🚀 Email Content and List Upload")
# email_subject = st.text_input("Email Subject", "Potential Collaboration Opportunity with Geomakani")
# uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])

# email_body = st.text_area("Email Body (HTML)", height=300)

# if st.button("Send Emails"):
#     if uploaded_file is not None and sender_email and password and email_body:
#         data = pd.read_excel(uploaded_file)

#         # SMTP Connection
#         server = smtplib.SMTP_SSL(smtp_server, smtp_port)
#         server.login(sender_email, password)

#         success_count, fail_count = 0, 0
#         progress_bar = st.progress(0)
#         status_text = st.empty()

#         for index, row in data.iterrows():
#             company_name = row['Company_Name_English']
#             recipient_email = row['Email']

#             # Create Email
#             msg = MIMEMultipart('alternative')
#             msg['Subject'] = email_subject
#             msg['From'] = sender_email
#             msg['To'] = recipient_email
#             msg['Bcc'] = sender_email

#             # Personalized Email body
#             personalized_html = email_body.replace("{company_name}", company_name)
#             msg.attach(MIMEText(personalized_html, 'html'))

#             # Sending Email
#             try:
#                 server.sendmail(sender_email, [recipient_email, sender_email], msg.as_string())
#                 success_count += 1
#                 status_text.success(f"✅ Email sent to {company_name}")
#             except Exception as e:
#                 fail_count += 1
#                 status_text.error(f"❌ Failed to send to {recipient_email}: {e}")
#                 continue

#             # IMAP Saving
#             try:
#                 imap = imaplib.IMAP4_SSL(imap_server, imap_port)
#                 imap.login(sender_email, password)
#                 imap.select('"INBOX.Sent"')
#                 imap.append('"INBOX.Sent"', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
#                 imap.logout()
#             except Exception as e:
#                 st.warning(f"⚠️ Failed to save email to Sent folder for {recipient_email}: {e}")

#             progress_bar.progress((index + 1) / len(data))
#             time.sleep(2)

#         server.quit()
#         st.success(f"🎉 Emails sent: {success_count}, Failures: {fail_count}")
#     else:
#         st.error("⚠️ Please fill all required fields and upload an Excel file.")



# Install streamlit first by running: pip install streamlit
# import streamlit as st
# import pandas as pd
# import smtplib
# import imaplib
# import email
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# import markdown
# import time


# # --- Streamlit UI ---
# st.set_page_config(page_title="Email Automation System", layout="wide")
# st.title("📧 Geomakani Email Automation")

# # --- Sidebar for credentials ---
# st.sidebar.header("📌 Email Server Configuration")
# smtp_server = st.sidebar.text_input("SMTP Server", "smtp.hostinger.com")
# smtp_port = st.sidebar.number_input("SMTP Port", value=465)
# imap_server = st.sidebar.text_input("IMAP Server", "imap.hostinger.com")
# imap_port = st.sidebar.number_input("IMAP Port", value=993)
# sender_email = st.sidebar.text_input("Sender Email")
# password = st.sidebar.text_input("Password", type="password")

# # --- Main Interface ---
# st.header("🚀 Email Content and List Upload")
# email_subject = st.text_input("Email Subject", "Potential Collaboration Opportunity with Geomakani")
# uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])

# email_body_markdown = st.text_area("Email Body (Markdown)", height=300)

# # Convert Markdown to HTML
# email_body = markdown.markdown(email_body_markdown)

# if st.button("Send Emails"):
#     if uploaded_file is not None and sender_email and password and email_body:
#         data = pd.read_excel(uploaded_file)

#         # SMTP Connection
#         server = smtplib.SMTP_SSL(smtp_server, smtp_port)
#         server.login(sender_email, password)

#         success_count, fail_count = 0, 0
#         progress_bar = st.progress(0)
#         status_text = st.empty()

#         for index, row in data.iterrows():
#             company_name = row['Company_Name_English']
#             recipient_email = row['Email']

#             # Create Email
#             msg = MIMEMultipart('alternative')
#             msg['Subject'] = email_subject
#             msg['From'] = sender_email
#             msg['To'] = recipient_email
#             msg['Bcc'] = sender_email

#             # Personalized Email body
#             personalized_html = email_body.replace("{company_name}", company_name)
#             msg.attach(MIMEText(personalized_html, 'html'))

#             # Sending Email
#             try:
#                 server.sendmail(sender_email, [recipient_email, sender_email], msg.as_string())
#                 success_count += 1
#                 status_text.success(f"✅ Email sent to {company_name}")
#             except Exception as e:
#                 fail_count += 1
#                 status_text.error(f"❌ Failed to send to {recipient_email}: {e}")
#                 continue

#             # IMAP Saving
#             try:
#                 imap = imaplib.IMAP4_SSL(imap_server, imap_port)
#                 imap.login(sender_email, password)
#                 imap.select('"INBOX.Sent"')
#                 imap.append('"INBOX.Sent"', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
#                 imap.logout()
#             except Exception as e:
#                 st.warning(f"⚠️ Failed to save email to Sent folder for {recipient_email}: {e}")

#             progress_bar.progress((index + 1) / len(data))
#             time.sleep(2)

#         server.quit()
#         st.success(f"🎉 Emails sent: {success_count}, Failures: {fail_count}")
#     else:
#         st.error("⚠️ Please fill all required fields and upload an Excel file.")



# Install with: pip install streamlit streamlit-quill pandas openpyxl
# Install with: pip install streamlit streamlit-quill pandas openpyxl
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
                        <p style="margin: 0; color: #e67300; font-weight: bold;">{name}</p>
                        <p style="margin: 0; color: #e67300;">{title}</p>
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


