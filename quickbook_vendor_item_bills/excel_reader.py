"""Excel reader for ItemBill data in company_data.xlsx.

This reader expects the worksheet to contain the following columns (exact
column headings, case-insensitive when trimmed):

- "Supplier Name"
- "Invoice Date"
- "Invoice Num"

It reads the first worksheet in the workbook, converts Excel date values to
ISO-8601 strings for `invoice_date`, and preserves numeric `invoice_number`
values when they appear numeric. Rows missing supplier or invoice number are
skipped.
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional
from datetime import date, datetime

from openpyxl import load_workbook

from .models import ItemBill


def _normalise(header: Optional[str]) -> str:
    return str(header).strip().lower() if header is not None else ""


def extract_item_bills(workbook_path: Path) -> List[ItemBill]:
    """Read `workbook_path` and return a list of ItemBill objects.

    Raises FileNotFoundError when the workbook doesn't exist. Rows missing
    required fields are skipped.
    """

    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    workbook = load_workbook(filename=workbook_path, read_only=True, data_only=True)
    try:
        # Require the specific sheet used by the input file
        sheet_name = "account debit vendor"
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' not found in workbook")
        sheet = workbook[sheet_name]

        rows = sheet.iter_rows(values_only=True)
        headers_row = next(rows, None)
        if headers_row is None:
            workbook.close()
            return []

        headers = [_normalise(h) for h in headers_row]
        header_index = {h: idx for idx, h in enumerate(headers)}

        supplier_key = _normalise("Supplier Name")
        date_key = _normalise("Invoice Date")
        invoice_key = _normalise("Invoice Num")

        supplier_idx = header_index.get(supplier_key)
        date_idx = header_index.get(date_key)
        invoice_idx = header_index.get(invoice_key)

        bills: List[ItemBill] = []
        for row in rows:
            # supplier
            supplier_val = None
            if supplier_idx is not None and supplier_idx < len(row):
                supplier_val = row[supplier_idx]
            if supplier_val is None:
                continue
            supplier_name = str(supplier_val).strip()
            if not supplier_name:
                continue

            # invoice date
            invoice_date_val = None
            if date_idx is not None and date_idx < len(row):
                invoice_date_val = row[date_idx]

            if isinstance(invoice_date_val, (date, datetime)):
                invoice_date = invoice_date_val.isoformat()
            elif invoice_date_val is None:
                invoice_date = ""
            else:
                invoice_date = str(invoice_date_val).strip()

            # invoice number
            invoice_number_val = None
            if invoice_idx is not None and invoice_idx < len(row):
                invoice_number_val = row[invoice_idx]
            if invoice_number_val is None:
                # skip if invoice number missing
                continue

            try:
                if (
                    isinstance(invoice_number_val, float)
                    and invoice_number_val.is_integer()
                ):
                    invoice_number = int(invoice_number_val)
                else:
                    invoice_number = invoice_number_val
            except Exception:
                invoice_number = int(str(invoice_number_val).strip())

            bills.append(
                ItemBill(
                    supplier_name=supplier_name,
                    invoice_date=invoice_date,
                    invoice_number=invoice_number,
                    source="excel",
                )
            )

    finally:
        workbook.close()

    return bills


__all__ = ["extract_item_bills"]
