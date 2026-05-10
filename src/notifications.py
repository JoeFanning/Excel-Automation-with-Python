import smtplib
from email.message import EmailMessage

def send_summary_email(subject, body, to_email, sender_email, sender_password):
    """Sends a plain text summary email to a client/manager."""
    msg = EmailMessage()
    msg.set_content(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = to_email

    try:
        # For Gmail/Outlook standard SMTP
        with smtplib.SMTP_SSL('://gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Email failed: {e}")
        return False
