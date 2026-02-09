# \# 20250519\_logActivity\_rev2.py Usage

# 

# \## Overview

# This script automates logging CRM activities using Selenium and an Excel workbook

# (`Log Activity.xlsx`) that lives next to the script. It reads the

# `Activity Report` sheet, opens CRM records, and logs activity details.

# 

# \## Prerequisites

# \- \*\*Windows\*\* (script uses `ctypes.windll` to prevent sleep).

# \- \*\*Python 3.9+\*\* recommended.

# \- \*\*Google Chrome\*\* installed.

# \- Python dependencies:

# &nbsp; - `selenium`

# &nbsp; - `openpyxl`

# 

# You can install dependencies with:

# ```bash

# pip install selenium openpyxl

# ```

# 

# \## Required Files

# Ensure these files are in the same folder:

# \- `20250519\_logActivity\_rev2.py`

# \- `Log Activity.xlsx`

# 

# The workbook must be a valid `.xlsx` file. The script validates the file before

# loading it.

# 

# \## Excel Sheet Requirements

# The script reads the `Activity Report` sheet and expects the following header

# names (case/spacing-insensitive):

# \- `Account ID (Not Required)`

# \- `Contact Name (Not Required)`

# \- `Contact ID (Required)`

# \- `Description (Required)`

# \- `Activity Date (Required)`

# \- `Notes (Data required in this field)`

# 

# It will add a `Logged Successfully` column if it is missing.

# 

# \## Running the Script

# From the script directory:

# ```bash

# python 20250519\_logActivity\_rev2.py

# ```

# 

# \## Notes

# \- The script launches a Chrome browser window and drives the UI with Selenium.

# \- Keep the browser in the foreground during execution for best reliability.

# \- If a run fails, review `error.log` for details.

# 

# \## Other Files in This Folder

# \- `20250519\_logActivity\_rev2 - backup.py`: Backup copy of the script for reference or rollback.

# \- `Log Activity - Copy.xlsx`: Optional duplicate workbook; not used by the script unless renamed to `Log Activity.xlsx`.

# \- `Old/`: Archive of older scripts/files.

# \- `Selenium\_Excel\_logActivity\_2024.yaml`: Environment/dependency reference used to document required packages.

# \- `run\_20250519\_logActivity\_rev2\_self\_healing.bat`: Windows batch file to launch the script with any self-healing steps configured.

# \- `run\_log\_activity.bat`: Windows batch file to launch older log activity scripts.

# \- `LICENSE`: Repository license.

# \- `.gitattributes`: Git attributes configuration for the repository.

# \- `error.log`: Log file written by the script when exceptions occur.

# 

