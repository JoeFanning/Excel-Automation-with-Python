import logging
from src.gui import ExcelAutomationGUI  # Updated import path
from src.io_manager import setup_logging, merge_excel_files, save_to_excel, save_analysis_to_excel
from src.clean_sort_data import clean_data, sort_data
from src.calculations import perform_calculations
from src.visuals import create_sales_dashboard

# Global logger initialization
logger = setup_logging()


def run_sales_pipeline(files):
    """
    This function processes the entire data pipeline workflow sequence
    and returns (True, success_message) or (False, error_message).
    """
    try:
        logger.info(f"Pipeline triggered via desktop app for {len(files)} file(s).")

        # 1. Merge Excel files
        merged_df = merge_excel_files(files, "merged_sales.xlsx", logger)

        # 2. Process data in merged file
        merged_df = clean_data(merged_df, logger)
        merged_df = sort_data(merged_df, logger)

        # 3. Run Analysis & Charts
        metrics = perform_calculations(merged_df, logger)

        # 4. Draw visuals and capture the file path
        chart_file = create_sales_dashboard(merged_df, "output", logger)

        # 5. Final Save to Excel
        save_to_excel(merged_df, "output/merged_sales.xlsx", logger)

        analysis_file = "output/sales_analysis_report.xlsx"
        if "results_dict" in metrics:
            save_analysis_to_excel(metrics["results_dict"], analysis_file, logger)

        # ----------------------------------------------------------------------
        # LOCAL EMAIL STEPS REMOVED — HAND CONTROL OVER TO GITHUB CLOUD UPLOAD
        # ----------------------------------------------------------------------
        logger.info("Local calculation engine complete. Returning control to GUI for cloud deployment.")
        return True, "Weekly automation calculations completed successfully!"

    except Exception as e:
        logger.error(f"Critical Pipeline Error: {e}")
        return False, f"Critical Pipeline Error occurred:\n\n{e}"


def main():
    # Start up the permanent UI and hand it our pipeline function engine link
    app = ExcelAutomationGUI(pipeline_callback=run_sales_pipeline)
    app.run()


if __name__ == "__main__":
    main()
