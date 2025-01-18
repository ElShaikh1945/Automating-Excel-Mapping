# Automating-Excel-Mapping
```markdown
# Final Task - Expiration Date Script

This Python script is designed to process Excel files by copying and mapping data from specific sheets and columns. The script also includes an expiration date check to ensure it is used within a valid timeframe.

## Prerequisites

Before running the script, ensure you have the following installed:

- Python 3.x
- Required Python packages:
  - `openpyxl`
  - `pandas`
  - `tkinter`

You can install the required packages using pip:

```bash
pip install openpyxl pandas
```

## Usage

1. **Run the Script**: Execute the script using Python:

   ```bash
   python Final_Task_Expiration_Date.py
   ```

2. **Select Input File**: A file dialog will appear, prompting you to select an Excel file (`*.xlsx` or `*.xls`). Choose the file you want to process.

3. **Processing**: The script will:
   - Load the selected Excel file.
   - Map data from `Sheet1` and `Sheet2` to `Sheet3` based on predefined column mappings.
   - Perform additional data processing, such as calculating dates and mapping service codes.

4. **Save Output File**: After processing, a save file dialog will appear. Choose a location and filename to save the updated Excel file.

5. **Completion**: Once the file is saved, a success message will be displayed, and the script will exit.

## Expiration Date

The script includes an expiration date check. If the current date is beyond the expiration date (`2024-08-05`), the script will display an error message and terminate. If this happens, please contact `Muhammad.AlShaikh@Outlook.com` for further assistance.

## Column Mappings

The script uses the following column mappings to copy data from `Sheet1` and `Sheet2` to `Sheet3`:

- **Employee Code**: Copied from `Sheet1` column 2 to `Sheet3` column 4.
- **National ID**: Copied from `Sheet1` column 2 to `Sheet3` column 5.
- **DOB**: Calculated based on a formula and written to `Sheet3` column 7.
- **DOJ**: Copied from `Sheet1` column 4 to `Sheet3` column 13.
- **Effective From Date**: Copied from `Sheet1` column 3 to `Sheet3` column 10.
- **Service Description**: Copied from `Sheet1` column 7 to `Sheet3` column 21.
- **Employee Type**: Set to "Contract" in `Sheet3` column 6.
- **Supplier Code**: Set to "1600009494" in `Sheet3` column 18.
- **Supplier Name**: Set to "Ergo" in `Sheet3` column 19.
- **Market Unit**: Set to "EGYPT Foods" in `Sheet3` column 51.
- **Function**: Copied from `Sheet1` column 9 to `Sheet3` column 52.
- **Sub Function**: Copied from `Sheet1` column 10 to `Sheet3` column 53.
- **Business Category**: Copied from `Sheet1` column 11 to `Sheet3` column 56.
- **Product Line**: Copied from `Sheet1` column 12 to `Sheet3` column 57.
- **Employee Arabic Name**: Copied from `Sheet1` column 1 to `Sheet3` column 59.

Additionally, the script maps department and work location data from `Sheet1` to `Sheet3`.

## Service Code Mapping

The script also maps service codes and descriptions from `Sheet2` to `Sheet3` based on matching values in specific columns.

## Troubleshooting

- **Error Loading Workbook**: If the script fails to load the workbook, ensure the file is not corrupted and is in the correct format (`*.xlsx` or `*.xls`).
- **Expiration Error**: If the script has expired, contact `Muhammad.AlShaikh@Outlook.com` for assistance.

## License

This script is provided as-is, without any warranties. Use it at your own risk.

---

For any questions or issues, please contact `Muhammad.AlShaikh@Outlook.com`.
```

This `README.md` file provides a comprehensive guide on how to use the script, including prerequisites, usage instructions, and troubleshooting tips. You can customize it further based on your specific needs.
