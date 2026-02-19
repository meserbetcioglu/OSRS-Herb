OSRS Herblore Calculator - Setup
=================================

This folder is portable when zipped, but Python dependencies must be installed
on the target machine. Follow these steps after unzipping:

1) Install Python 3 (if not already installed)
   - Download from https://www.python.org/downloads/

2) Run the setup script
   - Double-click setup_venv.bat
   - This creates a .venv folder and installs dependencies

3) Open the Excel file
   - Open Herbology.xlsm
   - Enable macros when prompted

4) Use the "Update Prices" button
   - It will run update_prices.py using the local .venv

Notes:
- If macros are blocked: Excel -> File -> Options -> Trust Center -> Macro Settings
- The Prices sheet updates even while Excel is open (xlwings)
- Config sheet requires:
  - Discord handle in B4
  - E-mail in B5

Troubleshooting:
- If the button does nothing, confirm:
  - setup_venv.bat ran successfully
  - Python and .venv exist in the same folder
  - update_prices.py is in the same folder as the workbook
