# Warehouse Safety Checklist Application

## Overview

This application allows Warehouse Safety Champions and Managers to easily complete the standardized safety checklist, record near-miss incidents, and track compliance data for their specific warehouse.

The goal is to streamline the reporting process and ensure consistent data collection across all locations.

## Using the Application (`.exe` File)

1.  **No Installation Needed:** Simply double-click the `WarehouseSafetyTool.exe` file provided to you. The application should open directly.
2.  **Fill Metadata:** Complete all the fields in the "Report Information" section at the top (Warehouse Name, Location, Your Name, Role, etc.). This is crucial for tracking.
3.  **Complete Checklist:** Go through the "Checklist Items" tab and answer each question (Yes, No, N/A, or text input). Use the scrollbar on the right if needed.
4.  **Record Near Miss (If Applicable):**
    *   Go to the "Near Miss Report" tab.
    *   Fill in the details (Date, Location, Description, Action, Prevention).
    *   Click "Add Link..." to paste URLs (e.g., from Google Drive, OneDrive) for any supporting photos or documents related *specifically* to the near miss. Ensure these links are shared correctly so the administrator can view them.
5.  **Add General Links (Optional):**
    *   Go to the "General Links" tab.
    *   Click "Add Link..." to paste URLs for any general supporting documents or evidence related to the main checklist items (e.g., a link to the current Fire NOC document). Ensure link permissions are correct.
6.  **Record Action Points:** Use the "Action Points" tab to note any follow-up actions required or further recommendations.
7.  **Save Progress (Optional):** If you need to stop and resume later, go to `File -> Save Project As...` to save your work as a `.json` file on your computer. You can reopen it later using `File -> Open Project...`.
8.  **Export Report (CRITICAL STEP):**
    *   Once the checklist is complete for the reporting period (e.g., end of the month/week), go to `File -> Export Report As`.
    *   Choose either `Excel (.xlsx)` or `PDF (.pdf)`. PDF is often preferred for final reports.
    *   Save the exported report file to your computer.
9.  **Submit Report and Links (CRITICAL STEP):**
    *   You **MUST** send the exported report file (the `.xlsx` or `.pdf` you just saved) to the central administrator/project lead (e.g., via email, shared drive upload, as instructed).
    *   **VERY IMPORTANT:** Double-check that all the links you pasted into the application (for Near Misses or General Links) have the correct **sharing permissions** set (e.g., "Anyone with the link can view") so the administrator can actually open and see the evidence files. The application only includes the *link* in the report, not the file itself.

## Data Consolidation

Please note: This application generates *individual* reports for your warehouse. The central administrator is responsible for consolidating reports from all locations. Your timely submission of the standardized report is essential for this process.

---

## For Developers / Rebuilding the EXE (Optional)

This section is for those with the source code who need to set up the development environment or rebuild the executable.

### Setup Instructions

1.  **Prerequisites:** Python 3.7+ installed, `pip`.
2.  **Clone/Download Source:** Get the source code folder containing `main.py`, `requirements.txt`, etc.
3.  **Create Virtual Environment:**
    *   Open a terminal in the project folder.
    *   Run: `python -m venv venv`
4.  **Activate Virtual Environment:**
    *   Windows CMD: `venv\Scripts\activate`
    *   Windows PowerShell: `.\venv\Scripts\Activate.ps1`
    *   macOS/Linux: `source venv/bin/activate`
5.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    *(Installs `customtkinter`, `openpyxl`, `reportlab`)*

### How to Run (Development)

1.  Activate the virtual environment.
2.  Run: `python main.py`

### How to Generate Windows Executable (.exe)

1.  Activate the virtual environment.
2.  Install PyInstaller: `pip install pyinstaller`
3.  Run PyInstaller from the project directory:
    ```bash
    # Standard command
    pyinstaller --name WarehouseSafetyTool --windowed --onefile main.py

    # Optional: Add an icon (place .ico file in project root)
    # pyinstaller --name WarehouseSafetyTool --windowed --onefile --icon=app_icon.ico main.py
    ```
4.  The `.exe` will be in the `dist` folder.