import logging
from src.gui import ExcelAutomationGUI  
from src.io_manager import setup_logging, merge_excel_files, save_to_excel, save_analysis_to_excel
from src.clean_sort_data import clean_data, sort_data
from src.calculations import perform_calculations
from src.visuals import create_sales_dashboard
from src.mailer import send_email_report

# FIX 1: Remove global logger initialization to prevent duplicate file streams.
# We will initialize this safely in main() and fetch it globally via getLogger.
logger = None 

def run_sales_pipeline(files):
    """
    This function processes the entire data pipeline workflow sequence
    and returns (True, success_message) or (False, error_message).
    """
    # FIX 2: Safely hook back into the single initialized logging stream
    pipeline_logger = logging.getLogger("ExcelAutomation")
    
    try:
        pipeline_logger.info(f"Pipeline triggered via desktop app for {len(files)} file(s).")

        # 1. Merge Excel files
        merged_df = merge_excel_files(files, "merged_sales.xlsx", pipeline_logger)

        # 2. Process data in merged file
        merged_df = clean_data(merged_df, pipeline_logger)
        merged_df = sort_data(merged_df, pipeline_logger)

        # 3. Run Analysis & Charts
        metrics = perform_calculations(merged_df, pipeline_logger)

        # 4. Draw visuals and capture the file path
        chart_file = create_sales_dashboard(merged_df, "output", pipeline_logger)

        # 5. Final Save to Excel
        save_to_excel(merged_df, "output/merged_sales.xlsx", pipeline_logger)

        analysis_file = "output/sales_analysis_report.xlsx"
        if "results_dict" in metrics:
            save_analysis_to_excel(metrics["results_dict"], analysis_file, pipeline_logger)

        # 6. Prepare Email Summary
        summary_text = (
            f"Hello,\n\n"
            f"The weekly sales automation has completed successfully.\n\n"
            f"--- Key Metrics ---\n"
            f"Total Revenue: €{metrics['revenue']:,.2f}\n"
            f"Top Manager: {metrics['top_manager']}\n\n"
            f"Please find the visual dashboard and detailed report attached."
        )

        # 7. Send ONE Final Email with BOTH attachments
        attachments = [chart_file, analysis_file]

        send_email_report(
            "Weekly Sales Dashboard Report",
            summary_text,
            "client@example.com",
            attachments,
            pipeline_logger
        )

        pipeline_logger.info("Pipeline completed successfully without errors.")
        return True, "Weekly automation completed successfully!\nReports saved and summary email sent."

    except Exception as e:
        pipeline_logger.error(f"Critical Pipeline Error: {e}")
        return False, f"Critical Pipeline Error occurred:\n\n{e}"


def main():
    global logger
    # FIX 3: Initialize the logger EXACTLY once right as the application starts
    logger = setup_logging()
    
    # Start up the permanent UI and hand it our pipeline function engine link
    app = ExcelAutomationGUI(pipeline_callback=run_sales_pipeline)
    app.run()


if __name__ == "__main__":
    main()

