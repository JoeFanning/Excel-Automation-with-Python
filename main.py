import os
import base64
import logging
import requests
from src.gui import ExcelAutomationGUI
from src.io_manager import setup_logging, merge_excel_files, save_to_excel, save_analysis_to_excel
from src.clean_sort_data import clean_data, sort_data
from src.calculations import perform_calculations
from src.visuals import create_sales_dashboard

# We initialize this safely inside main() to prevent WinError 32 file handler locks
logger = None


def get_azure_token(tenant_id, client_id, client_secret):
    """Exchanges Azure App credentials for an OAuth2 Access Token via the official endpoint."""
    # url = f"https://microsoftonline.com/{tenant_id}/oauth2/v2.0/token" #This sent an email
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "microsoft.com"
    }

    response = requests.post(url, data=data, timeout=15)
    response.raise_for_status()
    return response.json()["access_token"]


def run_sales_pipeline(files):
    """
    Processes the data pipeline sequence locally, and then transmits
    the output files over the cloud utilizing Microsoft Graph API.
    """
    pipeline_logger = logging.getLogger("ExcelAutomation")

    try:
        # 1. Fetch Cloud Environment Variables
        tenant_id = os.environ.get("AZURE_TENANT_ID")
        client_id = os.environ.get("AZURE_CLIENT_ID")
        client_secret = os.environ.get("AZURE_CLIENT_SECRET")
        sender_email = os.environ.get("AZURE_SENDER_EMAIL")

        if not all([tenant_id, client_id, client_secret, sender_email]):
            raise ValueError("Missing one or more required AZURE environment variables.")

        pipeline_logger.info(f"Pipeline triggered via desktop app for {len(files)} file(s).")

        # 2. Local File Merge Routine
        merged_df = merge_excel_files(files, "merged_sales.xlsx", pipeline_logger)

        # 3. Process data in merged file
        merged_df = clean_data(merged_df, pipeline_logger)
        merged_df = sort_data(merged_df, pipeline_logger)

        # 4. Run Analysis & Generate Metrics
        metrics = perform_calculations(merged_df, pipeline_logger)

        # 5. Draw Visual Dashboard Canvas
        chart_file = create_sales_dashboard(merged_df, "output", pipeline_logger)

        # 6. Save Local Excel Datasets
        save_to_excel(merged_df, "output/merged_sales.xlsx", pipeline_logger)

        analysis_file = "output/sales_analysis_report.xlsx"
        if "results_dict" in metrics:
            save_analysis_to_excel(metrics["results_dict"], analysis_file, pipeline_logger)

        # 7. Read and Encode Spreadsheet to Base64 for API transmission
        if not os.path.exists(analysis_file):
            raise FileNotFoundError(f"Target attachment {analysis_file} not found.")

        with open(analysis_file, "rb") as f:
            encoded_excel = base64.b64encode(f.read()).decode("utf-8")

        # 8. Construct Microsoft Graph JSON Payload Architecture
        email_payload = {
            "message": {
                "subject": "Weekly Sales Dashboard Performance Report",
                "body": {
                    "contentType": "Text",
                    "content": (
                        f"Hello Team,\n\n"
                        f"The weekly sales automation has completed successfully.\n\n"
                        f"--- Key Metrics ---\n"
                        f"Total Revenue: €{metrics.get('revenue', 0):,.2f}\n"
                        f"Top Manager: {metrics.get('top_manager', 'N/A')}\n\n"
                        f"Please find the processed performance metrics spreadsheet attached below."
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

        # 9. Authenticate and Send Mail via the verified Graph Route URL
        pipeline_logger.info("Authenticating with Microsoft Identity Platform...")
        token = get_azure_token(tenant_id, client_id, client_secret)

        pipeline_logger.info("Transmitting email via Microsoft Graph API gateway...")
        send_url = f"microsoft.com{sender_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        response = requests.post(send_url, json=email_payload, headers=headers, timeout=30)
        response.raise_for_status()

        pipeline_logger.info("Pipeline completed successfully without errors. Email sent via Azure.")
        return True, "Weekly automation completed successfully!\nReports saved and summary email sent."

    except Exception as e:
        pipeline_logger.error(f"Critical Pipeline Error: {e}")
        return False, f"Critical Pipeline Error occurred:\n\n{e}"


def main():
    global logger
    # Initialize logger exactly once to avoid multiple streams conflicting on Windows
    logger = setup_logging()

    # Fire up UI framework loop and hand it our processing engine pipeline link
    app = ExcelAutomationGUI(pipeline_callback=run_sales_pipeline)
    app.run()


if __name__ == "__main__":
    main()
