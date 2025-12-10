# Item Bills CLI

This project orchestrates item bill synchronisation between an Excel workbook and QuickBooks Desktop. Students implement the Excel reader, QBXML gateway, and comparison logic.

JSON reports contain keys: status, generated_at, added_itembills, conflicts, same_itembills, error. A success report lists each added bill and conflict; a failure report sets status to "error" and populates the error string.

## Installation

Install dependencies using Poetry:
```bash
poetry install
```

## Usage

Command-line usage:
```bash
poetry run python -m quickbook_vendor_item_bills --workbook company_data.xlsx [--output report.json]
```

If you omit `--output`, the report defaults to `item_bills_report.json` in the current directory.

## Building as Executable

To build the project as a standalone `.exe`:

1. Install dependencies (including PyInstaller):
   ```bash
   poetry install
   ```

2. Build the executable:
   ```bash
   poetry run pyinstaller --onefile --name quickbook_vendor_item_bills --hidden-import win32timezone --hidden-import win32com.client build_exe.py
   ```

3. The executable will be created in the `dist` folder.

The `--hidden-import` flags ensure PyInstaller includes the Windows COM dependencies needed for QuickBooks integration.

### Running the Executable (Windows)

From Command Prompt or PowerShell:

```cmd
cd dist
quickbook_vendor_item_bills.exe --workbook C:\path\to\company_data.xlsx --output C:\path\to\item_bills_report.json
```

If you omit `--output`, the report defaults to `item_bills_report.json` in the current directory. You can also invoke it without `cd` by using the absolute path, e.g.:

```cmd
C:\Users\ChristianD\Projects\QB_Connector_Item_Bill_Python_Fall_2025\dist\quickbook_vendor_item_bills.exe --workbook C:\path\to\company_data.xlsx
```

## Notes

- Ensure QuickBooks Desktop is open with the target company file when running the CLI (reads current session) or during batch add operations.
- Excel reader expects the item bills worksheet in `company_data.xlsx`; adjust sheet/column mapping in `excel_reader.py` if your format differs.
