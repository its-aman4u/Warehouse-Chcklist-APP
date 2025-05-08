# Warehouse Safety Checklist Application (v5 - Branded UI)

## Overview

This application provides a **modern, branded, user-friendly interface** for warehouse personnel to complete a standardized safety checklist, record metadata, report near-miss incidents, and attach **links** to evidence files stored in cloud drives. It streamlines data collection across multiple sites for central administration.

**Workflow:** (Remains the same as V4)
1. Open `.exe`.
2. Fill Metadata & Checklist.
3. Record Near Miss + Add Evidence Links.
4. Add General Evidence Links.
5. Save Progress (Optional `.json`).
6. Export Report (Excel/PDF with hyperlinks).
7. **Manual Submission & Link Permissions:** User sends exported report AND ensures all links are accessible to the admin.

## Features

*   **Branded User Interface:** Built with `CustomTkinter`, incorporating specified brand colors for a professional look. Forced Light Mode for consistency.
*   **Metadata Capture:** Dedicated section for essential report information.
*   **Structured Checklist:** Standardized checklist items. Improved alignment and visual separators.
*   **Structured Near Miss Reporting:** Dedicated tab with specific fields and link attachments.
*   **General Evidence Links:** Separate tab for attaching general evidence links. Improved list alignment.
*   **Link Attachments:** URL inputs for cloud-stored evidence.
*   **Robust Save/Load:** Save/load progress using `.json` files.
*   **Hyperlinked Exports:** Generate comprehensive reports in Excel (`.xlsx`) or PDF (`.pdf`) with clickable hyperlinks (requires `openpyxl`, `reportlab`).
*   **Standalone Executable:** Instructions included to package using PyInstaller.

## Setup Instructions (For Developers/Packaging)

1.  **Prerequisites:** Python 3.7+, `pip`.
2.  **Save Files:** Save `main.py`, `requirements.txt`, `README.md` (this file) into a project directory.
3.  **Install Dependencies:**
    ```bash
    cd path/to/your/project_directory
    pip install -r requirements.txt
    ```
    *(Installs `customtkinter`, `openpyxl`, `reportlab`)*

## How to Run (Development/Testing)

1.  Navigate to the project directory in your terminal.
2.  Run: `python main.py`

## How to Generate Windows Executable (.exe) (For Distribution)

1.  **Install PyInstaller:** `pip install pyinstaller`
2.  **Navigate to Project Directory.**
3.  **Run PyInstaller:**
    ```bash
    # Usually sufficient
    pyinstaller --name WarehouseSafetyReportTool --windowed --onefile main.py

    # Optional: Add an icon
    # pyinstaller --name WarehouseSafetyReportTool --windowed --onefile --icon=app_icon.ico main.py
    ```
    *(Ensure CustomTkinter assets are bundled; usually automatic now)*
4.  **Find Executable:** In the `dist` folder.
5.  **Distribute:** Share the `.exe`. **Instruct users about ensuring link permissions and manual submission.**

## Data Consolidation Note

This application generates standardized individual reports per submission. **It does not automatically consolidate data from multiple warehouses.** The administrator receiving the exported reports needs to perform consolidation using external tools (Excel, scripts, etc.).

## Screenshots

*(Placeholder: Add screenshots showing the new branded UI)*