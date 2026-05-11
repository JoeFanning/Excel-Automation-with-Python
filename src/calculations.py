import pandas as pd

def perform_calculations(df, logger):
    logger.info("Performing calculations...")

    # --- 1. Core KPIs ---
    total_revenue = df["TotalPrice"].sum()
    total_units = df["Quantity"].sum()
    num_transactions = len(df)
    average_order_value = total_revenue / num_transactions if num_transactions > 0 else 0

    # --- 2. Grouped Analysis ---
    sales_by_location = df.groupby("Location").agg({
        "TotalPrice": "sum",
        "Quantity": "sum"
    }).round(2).rename(columns={"TotalPrice": "Revenue", "Quantity": "Units_Sold"})

    top_products = df.groupby("Product")["TotalPrice"].sum().nlargest(5)
    bottom_products = df.groupby("Product")["TotalPrice"].sum().nsmallest(3)
    sales_by_payment = df.groupby("PaymentType")["TotalPrice"].sum().round(2).sort_values(ascending=False)
    sales_by_timeofday = df.groupby("TimeOfDay")["TotalPrice"].sum().round(2).sort_values(ascending=False)

    manager_sales = df.groupby("StoreManager")["TotalPrice"].sum()
    top_manager = manager_sales.idxmax()
    top_manager_revenue = manager_sales.max()

    # --- 3. Prepare the Excel Dictionary ---
    analysis_results = {
        "Executive Summary": pd.DataFrame({
            "Metric": ["Total Revenue", "Total Units", "Transactions", "Avg Order Value"],
            "Value": [total_revenue, total_units, num_transactions, average_order_value]
        }),
        "Sales by Location": sales_by_location,
        "Top Products": top_products.reset_index(),
        "Bottom Products": bottom_products.reset_index(),  # Add this!
        "Time of Day": sales_by_timeofday.reset_index(),  # Add this!
        "Payment Types": sales_by_payment.reset_index()
    }

    # --- 4. Logging Summary (Terminal Output) ---
    logger.info(f"Total Revenue           : €{total_revenue:,.2f}")
    logger.info(f"Total Units Sold        : {total_units:,.0f}")
    logger.info(f"Total Transactions      : {num_transactions:,}")
    logger.info(f"Average Order Value     : €{average_order_value:.2f}\n")

    # Log the Series/Grouped data
    logger.info(f"Top 5 Best Selling Products :\n{top_products.map('€{:,.2f}'.format)}\n")
    logger.info(f"Bottom 3 Selling Products   :\n{bottom_products.map('€{:,.2f}'.format)}\n")

    # Log the DataFrame
    logger.info(
        f"Sales by Location :\n{sales_by_location.assign(Revenue=sales_by_location['Revenue'].map('€{:,.2f}'.format))}\n")

    # Log the remaining breakdowns
    logger.info(f"Sales by Time of Day    :\n{sales_by_timeofday.apply(lambda x: f'€{x:,.2f}')}\n")
    logger.info(f"Sales by Payment Type   :\n{sales_by_payment.apply(lambda x: f'€{x:,.2f}')}\n")
    logger.info(f"Top Store Manager       : {top_manager} (€{top_manager_revenue:,.2f})\n")

    # --- 5. Final Return (One return to rule them all) ---
    return {
        "revenue": total_revenue,
        "top_manager": top_manager,
        "results_dict": analysis_results
    }
