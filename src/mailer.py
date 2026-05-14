import os
import base64
import requests


def send_email_report(subject, body, to_email, attachments, logger):
    """
    Guaranteed safe desktop mailer.
    Uses Resend API to handle processing and attachment delivery.
    """
    # 1. Fetch your Resend API key safely from local system memory
    resend_key = os.environ.get("RESEND_API_KEY")

    if not resend_key:
        logger.error("Delivery Failure: RESEND_API_KEY environment variable is not defined.")
        logger.warning("Please add RESEND_API_KEY to your local PyCharm environment to test.")
        return

    # 2. Extract and compile attachments into Resend's expected base64 API layout
    resend_attachments = []
    if isinstance(attachments, str):
        attachments = [attachments]

    for file_path in attachments:
        if not os.path.isfile(file_path):
            logger.warning(f"File missing: {file_path}")
            continue
        try:
            with open(file_path, "rb") as f:
                encoded_content = base64.b64encode(f.read()).decode("utf-8")

            resend_attachments.append({
                "filename": os.path.basename(file_path),
                "content": encoded_content
            })
            logger.info(f"Attached for API transmission: {os.path.basename(file_path)}")
        except Exception as e:
            logger.error(f"Attachment conversion error: {e}")

    # 3. Construct the clean Resend API JSON Request Payload
    url = "resend.com"
    headers = {
        "Authorization": f"Bearer {resend_key}",
        "Content-Type": "application/json"
    }

    email_payload = {
        "from": "Automation Engine <onboarding@resend.dev>",
        "to": [to_email],
        "subject": subject,
        "text": body,
        "attachments": resend_attachments
    }

    # 4. Transmit data packet over clean HTTP API gateway
    try:
        logger.info("Connecting to Resend Cloud Engine API...")
        response = requests.post(url, json=email_payload, headers=headers, timeout=20)
        response.raise_for_status()
        logger.info(f"Report safely accepted by Resend and dispatched to {to_email}")
    except Exception as e:
        logger.error(f"Resend Delivery Exception: {e}")
