import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_email_report(subject, body, to_email, attachment_path, logger):
    sender_email = "yourname@gmail.com"  # Your Gmail
    app_password = "xxxx yyyy zzzz aaaa"  # The 16-character code you just copied


