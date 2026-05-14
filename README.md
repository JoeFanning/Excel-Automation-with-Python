# 📊 Python Desktop Application for Excel Sales Pipeline Automation

A professional, cross-platform **Desktop GUI Application** designed to automate the merging, analysis, cleaning, and reporting of weekly sales data. This system transforms raw spreadsheets into a multi-tab executive workbook and high-impact visual dashboards. 

The core data engineering, consolidation, and chart generation happen **100% locally and offline**, ensuring sensitive financial information never leaves the local machine. An active internet connection is utilized exclusively during the final stage to securely transmit reports via email.

---

## 🚀 Key Features for Clients


| Feature | Description |
| :--- | :--- |
| **User-Friendly Interface** | Clean desktop GUI window requiring zero technical background, command line, or terminal knowledge. |
| **Secure Offline Core** | Data processing runs completely locally on the hard drive, minimizing corporate data leak risks. |
| **Data Consolidation** | Automatically handles merging, cleaning, and structural sorting of multiple weekly workbooks. |
| **Multi-Tab Reporting** | Compiles a professional `.xlsx` master file with dedicated analytical tabs for core KPIs. |
| **Visual Performance Dashboards** | Renders high-impact graph visual charts (`.png`) for quick, boardroom-ready reviews. |
| **Automated Delivery** | Integrated mailing module to automatically send generated files to dynamic recipient addresses. |

---

## 📈 Multi-Tab Report Structure
The generated Excel report (`sales_analysis_report.xlsx`) automatically populates the following dedicated sheets:

*   **Executive Summary**: High-level corporate KPIs tracking Total Revenue, Units Sold, and Average Order Value (AOV).
*   **Sales by Location**: Regional store performance breakdowns used to identify geographic trends.
*   **Product Performance**: Deep-dive analytics tracking the Top 5 Best Sellers and Bottom 3 Performers.
*   **Payment Type Breakdown**: Revenue splits across customer transaction methods (Cash, Card, Online, Gift Cards).
*   **Time of Day Analysis**: Identifies peak operational transaction hours (Morning, Afternoon, Evening) to optimize staffing.

---

## 💻 Seamless Desktop User Experience

The application wraps a professional graphical user interface over a high-performance analytics backend. It is designed for straightforward operation:

1.  **Launch**: Double-click the desktop application icon to initialize the interface launcher window.
2.  **Configure**: Type the target recipient's email address directly into the secure interface text entry field.
3.  **Upload & Execute**: Click **"Select Files & Launch Pipeline"**. The native OS file explorer will pop up, allowing you to select your raw weekly spreadsheets.
4.  **Processing & Transmission**: The interface blocks double-clicks to prevent duplicate inputs, cleanses data offline, and packages files. 
    *   *⚠️ **Network Note**: The computer must be **online (connected to the internet)** during this step so the mailing module can connect to the SMTP servers and deliver the files.*
5.  **Complete**: A native desktop notification pop-up screen alerts the user of a successful pipeline run.

---

## 🛠️ Application Architecture

The software uses a clear **Front-End / Back-End Split** architecture, driving data processing through a smooth UI layout layer:

### 🖥️ Front-End (The Graphical Shell)
*   `src/gui.py`: Manages screen coordinates, input field states, geometry placement, and OS file explorer window handles using Tkinter.

### ⚙️ Back-End (The Engine Under the Hood)
*   `main.py`: The central operational driver coordinating application lifecycle events and backend logic routing.
*   `src/io_manager.py`: Handles high-speed file ingestion and complex multi-sheet Excel writing configurations via Pandas.
*   `src/clean_sort_data.py`: Handles data validation, normalizes text schemas, and drops malformed or partial data rows.
*   `src/calculations.py`: The data engine performing algorithm arithmetic, metric tracking, and sequence analysis.
*   `src/visuals.py`: Converts raw calculation matrices into polished, customer-facing graphics.
*   `src/mailer.py`: Handles low-level network commands to attach generated analytics files and mail them to stakeholders.

---

## 📦 Cross-Platform Cloud Compilation Matrix

This project leverages an automated **GitHub Actions CI/CD Pipeline** to compile native, standalone execution packages for all major operating systems. Every time code is pushed, isolated cloud runners compile separate deployment binaries:

*   **Windows App Bundle**: Compiles down to an independent **`.exe` file** (`dist/main.exe`).
*   **macOS App Bundle**: Compiles down to an independent Apple **`.app` package** (`dist/main.app`).
*   **Linux App Bundle**: Compiles down to a native Linux binary execute block (`dist/main`).

*Clients do not need Python, Pandas, or any dependencies installed on their computers to run the final application bundles.*

---

## 📂 Local Workspace Organization
To run the automation locally or deploy the standalone app package, ensure the application binary sits directly adjacent to your active input and output workspaces or folders:


## Project Architecture Blueprint
Here is how the automation network is working:

                  │   ExcelAutomation Repository  │
                  └───────────────┬───────────────┘
                                  │
         ┌────────────────────────┴────────────────────────┐
         ▼                                                 ▼
 🖥️ LOCAL DESKTOP APP                              ☁️ HEADLESS CLOUD PIPELINE
   - File: main.py                                   - File: run_reporter.py
   - Uses PyCharm's system environment               - Triggered via GitHub Actions on push
   - Launches a real Tkinter GUI frame               - Runs on a remote Linux server container
   - Compiles local data sheets                      - Merges all monthly sales sheets
   - Sends emails securely via Resend                - Emails the €2,207,643.55 report
```
