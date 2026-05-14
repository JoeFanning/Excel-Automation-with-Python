import os
import base64
import sys
import requests
import logging
from pathlib import Path
from src.io_manager import setup_logging, merge_excel_files, save_to_excel, save_analysis_to_excel
from src.clean_sort_data import clean_data, sort_data
from src.calculations import perform_calculations
from src.visuals import create_sales_dashboard

def get_azure_token(tenant_id, client_id, client_secret):
    """Exchanges Azure App credentials for an OAuth2 Access Token via the official endpoint."""
    url = f"microsoftonline.com{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "microsoft.com"
    }
    response = requests.post(url, data=data, timeout=15)
    response.raise_for_status()
    return response.json()["access_token"]

def main():
    # 1. Initialize Headless Cloud Logging
    logger = setup_logging()
    logger.info("Starting Headless Cloud Data Pipeline...")

    try:
        # 2. Fetch Cloud Environment Variables
        tenant_id = os.environ.get("AZURE_TENANT_ID")
        client_id = os.environ.get("AZURE_CLIENT_ID")
        client_secret = os.environ.get("AZURE_CLIENT_SECRET")
        sender_email = os.environ.get("AZURE_SENDER_EMAIL")

        if not all([tenant_id, client_id, client_secret, sender_email]):
            raise ValueError("Missing one or more required AZURE environment variables.")

        # 3. Scan the repository workspace root for any raw data Excel files
        input_folder = Path(".")
        excel_files = list(input_folder.glob("*.xlsx")) + list(input_folder.glob("*.xls"))
        
        # Exclude output files from being merged recursively
        files_to_process = [str(f) for f in excel_files if "merged_sales" not in f.name and "sales_analysis_report" not in f.name]

        if not files_to_process:
            logger.warning("No input Excel files found to process in the repository root.")
            return

        logger.info(f"Found {len(files_to_process)} spreadsheet file(s) to process in cloud.")

        # 4. Run Core Data Processing Pipeline Steps
        merged_df = merge_excel_files(files_to_process, "merged_sales.xlsx", logger)
        merged_df = clean_data(merged_df, logger)
        merged_df = sort_data(merged_df, logger)

        # 5. Run Calculations and Build Charts
        metrics = perform_calculations(merged_df, logger)
        
        # Ensure output folder directory branch structure exists
        Path("output").mkdir(exist_ok=True)
        chart_file = create_sales_dashboard(merged_df, "output", logger)

        # 6. Save Local Cloud Output Files
        save_to_excel(merged_df, "output/merged_sales.xlsx", logger)

        analysis_file = "output/sales_analysis_report.xlsx"
        if "results_dict" in metrics:
            save_analysis_to_excel(metrics["results_dict"], analysis_file, logger)

        # 7. Read and Encode Spreadsheet to Base64 for API transmission
        if not os.path.exists(analysis_file):
            raise FileNotFoundError(f"Target attachment file {analysis_file} not found.")

        with open(analysis_file, "rb") as f:
            encoded_excel = base64.b64encode(f.read()).decode("utf-8")

        # 8. Construct Microsoft Graph JSON Payload Architecture
        email_payload = {
            "message": {
                "subject": "Weekly Sales Dashboard Cloud Automation Performance Report",
                "body": {
                    "contentType": "Text",
                    "content": (
                        f"Hello Team,\n\n"
                        f"The weekly cloud sales data automation pipeline has successfully finished execution.\n\n"
                        f"--- Key Cloud Metrics ---\n"
                        f"Total Revenue: €{metrics.get('revenue', 0):,.2f}\n"
                        f"Top Manager: {metrics.get('top_manager', 'N/A')}\n\n"
                        f"Please find your processed performance metrics spreadsheet attached below."
                    )
                },
                "toRecipients": [
                    {"emailAddress": {"address": "joespirial@hotmail.com"}}
                ],
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": "sales_analysis_report.xlsx",
                        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "contentBytes": encoded_excel
                    }
                ]
            }
        }

        # 9. Fetch OAuth Token & Deliver via Graph Endpoint
        logger.info("Requesting cloud access key from Microsoft Identity...")
        token = get_azure_token(tenant_id, client_id, client_secret)

        logger.info("Transmitting compiled payload via Microsoft Graph Gateway...")
        send_url = f"microsoft.com{sender_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        response = requests.post(send_url, json=email_payload, headers=headers, timeout=30)
        response.raise_for_status()

        logger.info("Cloud execution completed flawlessly! Automated report dispatched.")

    except Exception as pipeline_err:
        logger.error(f"Cloud Execution Failure: {pipeline_err}")
        sys.exit(1)

if __name__ == "__main__":
    main()
