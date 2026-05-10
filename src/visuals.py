import matplotlib.pyplot as plt
import os


def create_sales_dashboard(df, output_folder, logger):
    """
    Generates a multi-chart dashboard of sales data.
    """
    try:
        logger.info("Generating multi-chart dashboard...")

        # Create a figure with 2 subplots (1 row, 2 columns)
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))

        # --- Chart 1: Sales by Payment Type (Pie Chart) ---
        payment_data = df.groupby("PaymentType")["TotalPrice"].sum()
        ax1.pie(payment_data, labels=payment_data.index, autopct='%1.1f%%',
                startangle=140, colors=['#ff9999', '#66b3ff', '#99ff99', '#ffcc99'])
        ax1.set_title('Revenue Distribution by Payment Type', fontsize=14)

        # --- Chart 2: Top 5 Products (Horizontal Bar Chart) ---
        top_products = df.groupby("Product")["TotalPrice"].sum().nlargest(5).sort_values()
        top_products.plot(kind='barh', color='teal', ax=ax2)
        ax2.set_title('Top 5 Best Selling Products', fontsize=14)
        ax2.set_xlabel('Revenue (€)')

        plt.tight_layout()

        # Save dashboard
        dashboard_path = os.path.join(output_folder, "sales_dashboard.png")
        plt.savefig(dashboard_path, dpi=300)
        plt.close()

        logger.info(f"Dashboard saved to {dashboard_path}")
        return dashboard_path

    except Exception as e:
        logger.error(f"Failed to create dashboard: {e}")
        return None
