from src.io_manager import setup_logging, merge_excel_files, save_to_excel
from src.clean_sort_data import clean_data, sort_data
from src.calculations import perform_calculations
from src.visuals import create_sales_dashboard, create_sales_dashboard  # Combined imports
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

        # This saves the individual chart
        create_sales_dashboard(merged_df, "output", logger)

        # This saves the dashboard and stores the path for the email
        chart_file = create_sales_dashboard(merged_df, "output", logger)

        # 3. Final Save to Excel
        save_to_excel(merged_df, "output/merged_sales.xlsx", logger)

        # 4. Prepare Email Summary (Must be done BEFORE sending)
        summary_text = (
            f"Hello,\n\n"
            f"The weekly sales automation has completed successfully.\n\n"
            f"--- Key Metrics ---\n"
            f"Total Revenue: €{metrics['revenue']:,.2f}\n"
            f"Top Manager: {metrics['top_manager']}\n\n"
            f"Please find the visual dashboard attached."
        )

        # 5. Send ONE Final Email with the Dashboard
        send_email_report(
            "Weekly Sales Dashboard Report",
            summary_text,
            "client@example.com",  # Change to your email for testing!
            chart_file,
            logger
        )

    except Exception as e:
        logger.error(f"Critical Error: {e}")


if __name__ == "__main__":
    main()
