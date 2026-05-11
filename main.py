from src.io_manager import setup_logging, merge_excel_files, save_to_excel, save_analysis_to_excel
from src.clean_sort_data import clean_data, sort_data
from src.calculations import perform_calculations
from src.visuals import create_sales_dashboard
from src.mailer import send_email_report

def main():
    logger = setup_logging()
    try:
        # 1. Process Data
        merged_df = merge_excel_files("input", "merged_sales.xlsx", logger)
        merged_df = clean_data(merged_df, logger)
        merged_df = sort_data(merged_df, logger)

        # 2. Run Analysis & Charts
        metrics = perform_calculations(merged_df, logger)

        # Draw visuals and capture the file path (Only call this ONCE)
        chart_file = create_sales_dashboard(merged_df, "output", logger)

        # 3. Final Save to Excel
        save_to_excel(merged_df, "output/merged_sales.xlsx", logger)

        analysis_file = "output/sales_analysis_report.xlsx"
        if "results_dict" in metrics:
            save_analysis_to_excel(metrics["results_dict"], analysis_file, logger)

        # 4. Prepare Email Summary
        summary_text = (
            f"Hello,\n\n"
            f"The weekly sales automation has completed successfully.\n\n"
            f"--- Key Metrics ---\n"
            f"Total Revenue: €{metrics['revenue']:,.2f}\n"
            f"Top Manager: {metrics['top_manager']}\n\n"
            f"Please find the visual dashboard and detailed report attached."
        )

        # 5. Send ONE Final Email with BOTH attachments
        attachments = [chart_file, analysis_file]

        send_email_report(
            "Weekly Sales Dashboard Report",
            summary_text,
            "client@example.com",
            attachments,  # Pass the list here
            logger
        )

    except Exception as e:
        logger.error(f"Critical Error: {e}")

if __name__ == "__main__":
    main()
