# 📊 Sales Analysis & Excel Automation System

A professional Python-based solution designed to automate the merging, analysis, and reporting of weekly sales data. This system transforms raw data into a multi-tab executive report and a high-impact visual dashboard.

## 🚀 Key Features for Clients


| Feature | Description |
| :--- | :--- |
| **Data Consolidation** | Automatically merges multiple weekly Excel workbooks into a single master file. |
| **Multi-Tab Reporting** | Generates a professional `.xlsx` file with dedicated sheets for each KPI. |
| **Sales by Payment** | Breakdown of revenue across **Cash, Card, Online, and Gift Cards**. |
| **Visual Insights** | Automatically generates a `.png` dashboard for quick performance reviews. |
| **Direct Delivery** | Built-in email module to send reports directly to stakeholders. |

---

## 📈 Multi-Tab Report Structure
The generated Excel report (`sales_analysis_report.xlsx`) includes the following dedicated tabs:

1. **Executive Summary** - Core KPIs: Revenue, Units Sold, and Average Order Value.
2. **Sales by Location** - Regional performance across all store branches.
3. **Product Performance** - Analysis of Top 5 Best Sellers and Bottom 3 Performers.
4. **Payment Type Breakdown** - Revenue split by different payment methods.
5. **Time of Day** - Identifying peak sales hours (Morning, Afternoon, Evening).

---

## ⚡ One-Click Execution
Designed for non-technical users. No need to open a terminal—simply use the launcher for your system:

*   **Windows:** Double-click `run_automation.bat`
*   **macOS/Linux:** Double-click `run_automation.command`

---

## 🛠 Project Architecture

*   `main.py`: The central coordinator for the entire automation pipeline.
*   `src/io_manager.py`: Manages file merging and complex multi-sheet Excel writing.
*   `src/calculations.py`: The engine performing sales math and trend analysis.
*   `src/visuals.py`: Generates the graphical charts and dashboards.
*   `src/mailer.py`: Handles secure SMTP email delivery with attachments.

---

## 💻 How to Use

1. **Input:** Drop your weekly Excel files into the `/input` folder.
2. **Process:** Run the automation using the `.bat` / `.command` files or `python main.py`.
3. **Output:** Your finalized report and dashboard will be waiting in the `/output` folder.

---
