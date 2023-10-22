from dotenv import load_dotenv
import os
from email.message import EmailMessage
import ssl
import smtplib

load_dotenv()

email_sender = os.getenv('EMAIL_SENDER')
email_password = os.getenv('EMAIL_APP_PASSWORD_FOR_PYTHON')
email_recipient = "kopimenan@gmail.com"

email_subject = "test email subject"
email_body = "test email body"

em = EmailMessage()
em['Subject'] = email_subject
em['From'] = email_sender
em['To'] = email_recipient
em.set_content(email_body)

ssl_context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl_context) as server:
    server.login(email_sender, email_password)
    server.send_message(em)

