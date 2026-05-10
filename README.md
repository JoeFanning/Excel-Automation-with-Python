# Excel Automation with Python

This repository provides a streamlined way to automate Microsoft Excel tasks using Python. It is built for efficiency, using modern standards to manage dependencies and providing simple, one-click shortcuts for users on any operating system.

## 🚀 Key Features

*   **One-Click Workflows:** Execute complex Python logic via simple batch or command scripts.
*   **Modern Management:** Uses `pyproject.toml` for seamless installation and dependency handling.
*   **Data Power:** Leverages **pandas** for heavy data crunching and **openpyxl** for fine-tuned Excel formatting.
*   **Cross-Platform:** Native execution support for Windows and macOS.

## 🖱️ One-Click Execution

No need to open a code editor or use the terminal manually. Just use the launcher for your system:

### **Windows (`.bat`)**
- **To Run:** Double-click `run_automation.bat`.
- **Benefit:** It triggers the Python environment and keeps the window open so you can see the "Success" message or troubleshoot any data errors.

### **macOS (`.command`)**
- **To Run:** Double-click `run_automation.command`.
- **Note:** On first use, you may need to grant permission by right-clicking the file and selecting "Open," or by running `chmod +x run_automation.command` in your terminal to make it executable.

## 🛠️ Installation & Setup

1.  **Clone the Repo:**
    ```bash
    git clone https://github.com/JoeFanning/Excel-Automation-with-Python.git)
    cd Excel-Automation-with-Python
    ```

2.  **Install Dependencies:**
    This project is configured for a standard installation. Simply run:
    ```bash
    pip install .
    ```
    This will automatically read the `pyproject.toml` configuration and install all necessary tools like `pandas` and `openpyxl`.

## 💻 Usage

1. Place your source Excel files in the designated input folder.
2. Run your OS-specific launcher (the `.bat` or `.command` file).
3. Find your processed results in the output folder.

## 📄 License
This project is open-source and free to use for personal or commercial automation.

