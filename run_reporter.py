import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def main():
    # Pulls your secure Brevo credentials out of the GitHub cloud vault
    sender_email = os.environ.get("BREVO_SENDER_EMAIL")
    brevo_smtp_key = os.environ.get("BREVO_SMTP_KEY")
    
    # Matches the exact file your desktop app uploaded
    file_name = "sales_analysis_report.xlsx" 
    target_recipient = "joespirial@hotmail.com"

    if not os.path.exists(file_name):
        print(f"Error: Target spreadsheet {file_name} was not found in the repository root.")
        return

    # 1. Build the email headers
    msg = MIMEMultipart()
    msg['From'] = f"Automated Reporter <{sender_email}>"
    msg['To'] = target_recipient
    msg['Subject'] = "Weekly Sales Dashboard Performance Report"

    # 2. Design the message body content
    body_content = (
        "Hello Team,\n\n"
        "The weekly sales data automation pipeline has successfully completed data execution.\n\n"
        "Please find your processed performance metrics and compiled data dashboard attached below.\n\n"
        "Best Regards,\n"
        "Automated Reporting Engine"
    )
    msg.attach(MIMEText(body_content, 'plain'))

    # 3. Attach the Excel spreadsheet binary
    with open(file_name, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={file_name}")
    msg.attach(part)

        # 4. Route securely through Brevo SMTP relay using implicit SSL
    try:
        print("Connecting to Brevo Relay Server over Secure Port 465...")
        
        # CHANGED: Using SMTP_SSL and port 465 to bypass firewall drops
        with smtplib.SMTP_SSL("smtp-relay.brevo.com", 465) as server:
            server.ehlo()
            server.login(sender_email, brevo_smtp_key)
            server.send_message(msg)
            
        print("Success! Email sent smoothly via GitHub Cloud.")
    except Exception as e:
        print(f"Delivery Failure Error: {e}")


if __name__ == "__main__":
    main()

