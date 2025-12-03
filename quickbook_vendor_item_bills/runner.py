"""High-level orchestration for the item bills CLI and summary writer."""

from __future__ import annotations

from pathlib import Path
from typing import Dict
import argparse

from . import comparer, excel_reader, qb_gateway
from .models import ItemBill, Conflict
from .reporting import iso_timestamp, write_report

DEFAULT_REPORT_NAME = "item_bills_report.json"


def _iso(val):
    """Return ISO string for date-like values; pass through strings; else None.

    Supports both datetime/date objects and pre-existing strings from tests.
    """
    if val is None:
        return None
    if hasattr(val, "isoformat"):
        try:
            return val.isoformat()
        except Exception:
            pass
    # If already a string (e.g., test fixtures), keep as-is
    return val


def _bill_to_dict(bill: ItemBill) -> Dict[str, object]:
    return {
        "supplier_name": bill.supplier_name,
        "invoice_date": _iso(bill.invoice_date),
        "invoice_number": bill.invoice_number,
        "source": bill.source,
    }


def _conflict_to_dict(conflict: Conflict) -> Dict[str, object]:
    """Serialize a Conflict in the detailed item-bills shape."""
    return {
        "record_id": conflict.id,
        "reason": conflict.reason,
        "excel_supplier_name": conflict.excel_supplier_name,
        "qb_supplier_name": conflict.qb_supplier_name,
        "excel_invoice_number": conflict.excel_invoice_number,
        "qb_invoice_number": conflict.qb_invoice_number,
        "excel_invoice_date": _iso(conflict.excel_invoice_date),
        "qb_invoice_date": _iso(conflict.qb_invoice_date),
    }


def _missing_in_excel_conflict(bill: ItemBill) -> Dict[str, object]:
    return {
        "record_id": bill.id or bill.invoice_number,
        "reason": "missing_in_excel",
        "excel_supplier_name": None,
        "qb_supplier_name": bill.supplier_name,
        "excel_invoice_number": None,
        "qb_invoice_number": bill.invoice_number,
        "excel_invoice_date": None,
        "qb_invoice_date": _iso(bill.invoice_date),
    }


def _missing_in_quickbooks_conflict(bill: ItemBill) -> Dict[str, object]:
    return {
        "record_id": bill.id or bill.invoice_number,
        "reason": "missing_in_quickbooks",
        "excel_supplier_name": bill.supplier_name,
        "qb_supplier_name": None,
        "excel_invoice_number": bill.invoice_number,
        "qb_invoice_number": None,
        "excel_invoice_date": _iso(bill.invoice_date),
        "qb_invoice_date": None,
    }


def run_item_bills(
    company_file_path: str,
    workbook_path: str,
    *,
    output_path: str | None = None,
) -> Path:
    """Contract entry point for synchronising item bills.

    Args:
        company_file_path: Path to the QuickBooks company file. Use an empty
            string to reuse the currently open company file.
        workbook_path: Path to the Excel workbook containing the
            item bills worksheet.
        output_path: Optional JSON output path. Defaults to
            item_bills_report.json in the current working directory.

    Returns:
        Path to the generated JSON report.
    """

    report_path = Path(output_path) if output_path else Path(DEFAULT_REPORT_NAME)
    report_payload: Dict[str, object] = {
        "status": "success",
        "generated_at": iso_timestamp(),
        "added_itembills": [],
        "conflicts": [],
        "same_itembills": 0,
        "error": None,
    }

    try:
        excel_bills = excel_reader.extract_item_bills(Path(workbook_path))
        qb_bills = qb_gateway.fetch_item_bills(company_file_path)
        comparison = comparer.compare_item_bills(excel_bills, qb_bills)

        # Build conflicts list in required shape
        conflicts: list[Dict[str, object]] = []
        conflicts.extend(
            _conflict_to_dict(conflict) for conflict in comparison.conflicts
        )
        conflicts.extend(
            _missing_in_excel_conflict(bill) for bill in comparison.qb_only
        )
        conflicts.extend(
            _missing_in_quickbooks_conflict(bill) for bill in comparison.excel_only
        )

        report_payload["conflicts"] = conflicts

        # Compute same_itembills: matched records without mismatches
        matched_pairs = len(excel_bills) - len(comparison.excel_only)
        same_itembills = matched_pairs - len(comparison.conflicts)
        report_payload["same_itembills"] = max(0, same_itembills)

        # Populate added_itembills by creating bills in QB (off by default)
        added_bills = qb_gateway.add_item_bills_batch(
            company_file_path, comparison.excel_only
        )
        report_payload["added_itembills"] = [
            {
                "record_id": b.id or b.invoice_number,
                "supplier_name": b.supplier_name,
                "invoice_number": b.invoice_number,
                "invoice_date": _iso(b.invoice_date),
            }
            for b in added_bills
        ]

    except Exception as exc:  # pragma: no cover - behaviour verified via tests
        report_payload["status"] = "error"
        report_payload["error"] = str(exc)

    write_report(report_payload, report_path)
    return report_path


__all__ = ["run_item_bills", "DEFAULT_REPORT_NAME"]

if __name__ == "__main__":  # pragma: no cover - manual execution helper
    parser = argparse.ArgumentParser(
        description="Synchronise item bills and write a JSON report."
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Path to Excel workbook containing the account debit vendor worksheet",
    )
    parser.add_argument("--output", help="Optional JSON output path")
    args = parser.parse_args()

    path = run_item_bills("", args.workbook, output_path=args.output)
    print(f"Report written to {path}")
