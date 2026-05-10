import pandas as pd

def perform_calculations(df, logger):
    """
    Takes the cleaned dataframe and logger, performs sales analysis,
    and logs the results with Euro formatting.
    """
    logger.info("Performing calculations...")

    # --- Core KPIs ---
    total_revenue = df["TotalPrice"].sum()
    total_units = df["Quantity"].sum()
    num_transactions = len(df)
    average_order_value = total_revenue / num_transactions if num_transactions > 0 else 0

    # --- Sales by Location ---
    sales_by_location = df.groupby("Location").agg({
        "TotalPrice": "sum",
        "Quantity": "sum"
    }).round(2).rename(columns={"TotalPrice": "Revenue", "Quantity": "Units_Sold"})

    # --- Product Performance ---
    top_products = df.groupby("Product")["TotalPrice"].sum().nlargest(5)
    bottom_products = df.groupby("Product")["TotalPrice"].sum().nsmallest(3)

    # --- Grouped Analysis ---
    sales_by_payment = df.groupby("PaymentType")["TotalPrice"].sum().round(2).sort_values(ascending=False)
    sales_by_timeofday = df.groupby("TimeOfDay")["TotalPrice"].sum().round(2).sort_values(ascending=False)

    # --- Top Performing Manager ---
    manager_sales = df.groupby("StoreManager")["TotalPrice"].sum()
    top_manager = manager_sales.idxmax()
    top_manager_revenue = manager_sales.max()

    # ========== LOGGING SUMMARY ==========
    logger.info(f"Total Revenue           : €{total_revenue:,.2f}")
    logger.info(f"Total Units Sold        : {total_units:,.0f}")
    logger.info(f"Total Transactions      : {num_transactions:,}")
    logger.info(f"Average Order Value     : €{average_order_value:.2f}\n")

    # Log Series with .map
    logger.info(f"Top 5 Best Selling Products :\n{top_products.map('€{:,.2f}'.format)}\n")
    logger.info(f"Bottom 3 Selling Products   :\n{bottom_products.map('€{:,.2f}'.format)}\n")

    # Log DataFrame with .assign
    logger.info(f"Sales by Location :\n{sales_by_location.assign(Revenue=sales_by_location['Revenue'].map('€{:,.2f}'.format))}\n")

    # Log remaining Series with .apply
    logger.info(f"Sales by Time of Day    :\n{sales_by_timeofday.apply(lambda x: f'€{x:,.2f}')}\n")
    logger.info(f"Sales by Payment Type   :\n{sales_by_payment.apply(lambda x: f'€{x:,.2f}')}\n")
    logger.info(f"Top Store Manager       : {top_manager} (€{top_manager_revenue:,.2f})\n")

    # Return the metrics in case you want to use them for a final dashboard/email
    return {
        "revenue": total_revenue,
        "top_manager": top_manager,
        "sales_by_location": sales_by_location
    }
