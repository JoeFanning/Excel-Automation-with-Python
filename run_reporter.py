import os
import base64
import requests

def get_azure_token(tenant_id, client_id, client_secret):
    """Exchanges Azure App credentials for an OAuth2 Access Token."""
    url = f"microsoftonline.com{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "microsoft.com"
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

def main():
    # 1. Fetch Azure cloud variables
    tenant_id = os.environ.get("da0617dd-b536-4432-b4bf-390fb40861c8")
    client_id = os.environ.get("3c804a46-8f84-4122-8e63-4804333c568a")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    sender_email = os.environ.get("joespirial@hotmail.com")
    
    file_name = "sales_analysis_report.xlsx"
    target_recipient = "joespirial@hotmail.com"

    if not all([tenant_id, client_id, client_secret, sender_email]):
        print("Error: Missing one or more required AZURE environment variables.")
        return

    if not os.path.exists(file_name):
        print(f"Error: Target spreadsheet {file_name} was not found.")
        return

    # 2. Read and encode the Excel file to Base64
    with open(file_name, "rb") as f:
        file_content = f.read()
        encoded_file = base64.b64encode(file_content).decode("utf-8")

    # 3. Construct Microsoft Graph payload
    email_payload = {
        "message": {
            "subject": "Weekly Sales Dashboard Performance Report",
            "body": {
                "contentType": "Text",
                "content": (
                    "Hello Team,\n\n"
                    "The weekly sales data automation pipeline has successfully completed data execution.\n\n"
                    "Please find your processed performance metrics and compiled data dashboard attached below.\n\n"
                    "Best Regards,\n"
                    "Automated Reporting Engine"
                )
            },
            "toRecipients": [
                {"emailAddress": {"address": target_recipient}}
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": file_name,
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": encoded_file
                }
            ]
        }
    }

    # 4. Authenticate and transmit via API
    try:
        print("Authenticating with Microsoft Identity Platform...")
        token = get_azure_token(tenant_id, client_id, client_secret)
        
        print("Transmitting email via Microsoft Graph API...")
        send_url = f"microsoft.com{sender_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.post(send_url, json=email_payload, headers=headers)
        response.raise_for_status()
        
        print("Success! Email sent smoothly via Microsoft Azure.")
    except Exception as e:
        print(f"Delivery Failure Error: {e}")

if __name__ == "__main__":
    main()

