"""High-level orchestration for the payment term CLI."""

from __future__ import annotations

from pathlib import Path
from typing import Dict

from . import comparer, excel_reader, qb_gateway
from .models import ItemBill
from .reporting import iso_timestamp, write_report

DEFAULT_REPORT_NAME = "item_bills_report.json"


def _bill_to_dict(bill: ItemBill) -> Dict[str, object]:
    return {
        "supplier_name": bill.supplier_name,
        "invoice_date": bill.invoice_date,
        "invoice_number": bill.invoice_number,
        "source": bill.source,
    }


def _conflict_to_dict(conflict) -> Dict[str, object]:
    # Support both the new Conflict dataclass and older shapes by falling
    # back to multiple possible attribute names.
    invoice_number = getattr(conflict, "invoice_number", None)
    if invoice_number is None:
        invoice_number = getattr(conflict, "record_id", None)

    excel_supplier = getattr(conflict, "excel_supplier", None)
    if excel_supplier is None:
        excel_supplier = getattr(conflict, "excel_name", None)

    qb_supplier = getattr(conflict, "qb_supplier", None)
    if qb_supplier is None:
        qb_supplier = getattr(conflict, "qb_name", None)

    excel_date = getattr(conflict, "excel_date", None)
    if excel_date is None:
        excel_date = getattr(conflict, "excel_date", None)

    qb_date = getattr(conflict, "qb_date", None)
    if qb_date is None:
        qb_date = getattr(conflict, "qb_date", None)

    return {
        "invoice_number": invoice_number,
        "excel_supplier": excel_supplier,
        "qb_supplier": qb_supplier,
        "excel_date": excel_date,
        "qb_date": qb_date,
        "reason": getattr(conflict, "reason", None),
    }


def _missing_in_excel_conflict(bill: ItemBill) -> Dict[str, object]:
    return {
        "invoice_number": bill.invoice_number,
        "excel_supplier": None,
        "qb_supplier": bill.supplier_name,
        "excel_date": None,
        "qb_date": bill.invoice_date,
        "reason": "missing_in_excel",
    }


def _missing_in_quickbooks_conflict(bill: ItemBill) -> Dict[str, object]:
    return {
        "invoice_number": bill.invoice_number,
        "excel_supplier": bill.supplier_name,
        "qb_supplier": None,
        "excel_date": bill.invoice_date,
        "qb_date": None,
        "reason": "missing_in_quickbooks",
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
        "added_bills": [],
        "conflicts": [],
        "error": None,
    }

    try:
        excel_bills = excel_reader.extract_item_bills(Path(workbook_path))
        qb_bills = qb_gateway.fetch_item_bills(company_file_path)
        comparison = comparer.compare_item_bills(excel_bills, qb_bills)

        # added_bills = qb_gateway.add_item_bills_batch(
        #     company_file_path, comparison.excel_only
        # )

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

        # report_payload["added_bills"] = [_bill_to_dict(bill) for bill in added_bills]
        report_payload["conflicts"] = conflicts

    except Exception as exc:  # pragma: no cover - behaviour verified via tests
        report_payload["status"] = "error"
        report_payload["error"] = str(exc)

    write_report(report_payload, report_path)
    return report_path


__all__ = ["run_item_bills", "DEFAULT_REPORT_NAME"]
