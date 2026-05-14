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

def main():
    logger = setup_logging()
    logger.info("Initializing Secure Resend API Engine Sequence...")

    try:
        resend_key = os.environ.get("RESEND_API_KEY")
        target_recipient = "joespirial@hotmail.com"

        if not resend_key:
            raise ValueError("CRITICAL FAILURE: The RESEND_API_KEY environment variable is empty.")

        # 1. Gather files from input directory
        input_folder = Path("input")
        excel_files = list(input_folder.glob("*.xlsx")) + list(input_folder.glob("*.xls"))
        files_to_process = [str(f) for f in excel_files if "merged_sales" not in f.name and "sales_analysis_report" not in f.name]

        if not files_to_process:
            logger.warning("No input Excel files detected inside the 'input/' folder. Stopping pipeline.")
            return

        # 2. Run your verified analytics logic blocks
        merged_df = merge_excel_files(files_to_process, "merged_sales.xlsx", logger)
        merged_df = clean_data(merged_df, logger)
        merged_df = sort_data(merged_df, logger)
        metrics = perform_calculations(merged_df, logger)
        
        Path("output").mkdir(exist_ok=True)
        chart_file = create_sales_dashboard(merged_df, "output", logger)

        # 3. Build and save the target attachment file sheets
        save_to_excel(merged_df, "output/merged_sales.xlsx", logger)
        analysis_file = "output/sales_analysis_report.xlsx"
        if "results_dict" in metrics:
            save_analysis_to_excel(metrics["results_dict"], analysis_file, logger)

        # 4. Binary Base64 String Compilation mapping
        with open(analysis_file, "rb") as f:
            encoded_content = base64.b64encode(f.read()).decode("utf-8")

        # 5. Core Resend Network URL API Endpoints Layout
        url = "https://api.resend.com/emails""
        headers = {
            "Authorization": f"Bearer {resend_key}",
            "Content-Type": "application/json"
        }
        
        email_payload = {
            "from": "Automation Engine <onboarding@resend.dev>",
            "to": [target_recipient],
            "subject": "Weekly Sales Dashboard Cloud Performance Report",
            "text": (
                f"Hello Team,\n\n"
                f"The weekly cloud sales data automation pipeline has successfully completed data execution.\n\n"
                f"Total Revenue: €{metrics.get('revenue', 0):,.2f}\n"
                f"Top Manager: {metrics.get('top_manager', 'N/A')}\n\n"
                f"Please find your analytical summary workbook attached below."
            ),
            "attachments": [
                {
                    "filename": "sales_analysis_report.xlsx",
                    "content": encoded_content
                }
            ]
        }

        logger.info("Transmitting encrypted email packet payload via Resend Cloud Engine...")
        response = requests.post(url, json=email_payload, headers=headers, timeout=20)
        response.raise_for_status()
        
        logger.info("Pipeline execution complete! Automated Resend dashboard dispatched successfully.")

    except Exception as pipeline_err:
        logger.error(f"Cloud Execution Failure Blocker: {pipeline_err}")
        sys.exit(1)

if __name__ == "__main__":
    main()
