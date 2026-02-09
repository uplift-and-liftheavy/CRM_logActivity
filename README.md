# 20250519_logActivity_rev2.py Usage

## Quick Start (Read This First)
1. **Install Python 3.11+** from [python.org/downloads](https://www.python.org/downloads/).
   - During install, check **“Add python.exe to PATH”** if prompted.
2. **Double‑click** `run_20250519_logActivity_rev2_self_healing.bat` in this folder.
   - It **creates a virtual environment** and **installs required packages** automatically the first time.
3. Follow the on‑screen prompts in the terminal window that opens.

## Overview
This script automates logging CRM activities using Selenium and an Excel workbook
(`Log Activity.xlsx`) that lives next to the script. It reads the
`Activity Report` sheet, opens CRM records, and logs activity details.

## Prerequisites
- **Windows** (script uses `ctypes.windll` to prevent sleep).
- **Python 3.11+** installed. Download from: [python.org/downloads](https://www.python.org/downloads/).
  - Tip: check **“Add python.exe to PATH”** during setup.
- **Google Chrome** installed.

## Required Files
Ensure these files are in the same folder:
- `20250519_logActivity_rev2.py`
- `Log Activity.xlsx`

The workbook must be a valid `.xlsx` file. The script validates the file before
loading it.

## Excel Sheet Requirements
The script reads the `Activity Report` sheet and expects the following header
names (case/spacing-insensitive):
- `Account ID (Not Required)`
- `Contact Name (Not Required)`
- `Contact ID (Required)`
- `Description (Required)`
- `Activity Date (Required)`
- `Notes (Data required in this field)`

It will add a `Logged Successfully` column if it is missing.

## Easiest Way to Run (Recommended)
Use the **self‑healing batch file**. It will:
- create a Python virtual environment in this folder,
- install missing dependencies automatically,
- then run the script.

### Step‑by‑Step (Beginner Friendly)
1. Open the folder that contains the files.
2. Find `run_20250519_logActivity_rev2_self_healing.bat`.
3. **Double‑click** it.
4. A black terminal window appears. Wait while it installs and runs.
5. When finished, press any key to close the window.

If you see an error that Python is missing, install it from
[python.org/downloads](https://www.python.org/downloads/) and run the BAT file again.

## Running the Script Manually (Advanced)
From the script directory:
```bash
python 20250519_logActivity_rev2.py
```

## Notes
- The script launches a Chrome browser window and drives the UI with Selenium.
- Keep the browser in the foreground during execution for best reliability.
- If a run fails, review `error.log` for details.

## Other Files in This Folder
- `20250519_logActivity_rev2 - backup.py`: Backup copy of the script for reference or rollback.
- `Log Activity - Copy.xlsx`: Optional duplicate workbook; not used by the script unless renamed to `Log Activity.xlsx`.
- `Old/`: Archive of older scripts/files.
- `Selenium_Excel_logActivity_2024.yaml`: Environment/dependency reference used to document required packages.
- `run_20250519_logActivity_rev2_self_healing.bat`: Windows batch file that installs dependencies and runs the script.
- `run_log_activity.bat`: Windows batch file to launch older log activity scripts.
- `LICENSE`: Repository license.
- `.gitattributes`: Git attributes configuration for the repository.
- `error.log`: Log file written by the script when exceptions occur.
