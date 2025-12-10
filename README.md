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

### Sample Output JSON file

```
{
  "status": "success",
  "generated_at": "2025-11-24T20:34:02.458327+00:00",
  "added_itembills": [
    {
      "record_id": "44444-11111",
      "supplier_name": "C",
      "invoice_number": "33333",
      "invoice_date": "2023-12-08"
    }
  ],
  "conflicts": [
    {
      "record_id": "44733-",
      "reason": "data_mismatch",
      "excel_supplier_name": "B",
      "qb_supplier_name": "B",
      "excel_invoice_number": "60956",
      "qb_invoice_number": "609",
      "excel_invoice_date": "2023-11-09",
      "qb_invoice_date": "2023-11-09"
    },
    {
      "record_id": "987",
      "reason": "missing_in_excel",
      "excel_supplier_name": null,
      "qb_supplier_name": "Piston Experts",
      "excel_invoice_number": null,
      "qb_invoice_number": "123",
      "excel_invoice_date": null,
      "qb_invoice_date": "2025-09-17"
    }
  ],
  "same_itembills": 7,
  "error": null
}

```


## Notes

- Ensure QuickBooks Desktop is open with the target company file when running the CLI (reads current session) or during batch add operations.
- Excel reader expects the item bills worksheet in `company_data.xlsx`; adjust sheet/column mapping in `excel_reader.py` if your format differs.
