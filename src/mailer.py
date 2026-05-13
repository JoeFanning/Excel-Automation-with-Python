import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def send_email_report(subject, body, to_email, attachments, logger):
    """
    Sends an email with multiple file attachments.
    :param attachments: List of file paths or a single file path string.
    """
    sender_email = "joespirial@hotmail.com"  # Replace with your Gmail
    app_password = "RYXJK-YH4YJ-R55LD-B9P9X-XT7GA"  # Replace with your 16-character App Password

    # Create the root message container
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Ensure attachments is a list (handles single string or list input)
    if isinstance(attachments, str):
        attachments = [attachments]

    # Loop through and attach each file
    for file_path in attachments:
        if not os.path.isfile(file_path):
            logger.warning(f"Attachment not found, skipping: {file_path}")
            continue

        try:
            with open(file_path, "rb") as f:
                # Use application/octet-stream for generic file types
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())

            # Encode in base64 to ensure proper transmission
            encoders.encode_base64(part)

            # Extract only the filename from the full path
            filename = os.path.basename(file_path)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")

            msg.attach(part)
            logger.info(f"Successfully attached: {filename}")
        except Exception as e:
            logger.error(f"Error attaching {file_path}: {e}")
'''
    # Send the email using Gmail's SMTP server
    try:
        # Use Port 587 with STARTTLS
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.send_message(msg)

        logger.info(f"✅ Email sent successfully to {to_email}")
    except Exception as e:
        logger.error(f"❌ Failed to send email: {e}")
'''


