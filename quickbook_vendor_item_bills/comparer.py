"""Comparison helpers for item bills."""

from __future__ import annotations

from typing import Iterable

from .models import ComparisonReport, Conflict, ItemBill
from .excel_reader import extract_item_bills

import argparse
from pathlib import Path


def compare_item_bills(
    excel_bills: Iterable[ItemBill],
    qb_bills: Iterable[ItemBill],
) -> ComparisonReport:
    """Compare Excel and QuickBooks Item bills and identify discrepancies."""

    # Build maps keyed by preferred id; fallback to invoice_number
    def _key(b: ItemBill) -> str:
        return str(b.id) if b.id else b.invoice_number

    excel_map = {_key(bill): bill for bill in excel_bills}
    qb_map = {_key(bill): bill for bill in qb_bills}

    excel_only: list[ItemBill] = []
    qb_only: list[ItemBill] = []
    conflicts: list[Conflict] = []

    # Bills present in Excel but not QuickBooks
    for k, e_bill in excel_map.items():
        if k not in qb_map:
            excel_only.append(e_bill)
        else:
            qb_bill = qb_map[k]

            # Determine if any mismatch exists between paired bills
            mismatch = False
            if (e_bill.supplier_name or "") != (qb_bill.supplier_name or ""):
                mismatch = True
            if e_bill.invoice_date != qb_bill.invoice_date:
                mismatch = True
            if (e_bill.invoice_number or "") != (qb_bill.invoice_number or ""):
                mismatch = True

            def _parts_map(b: ItemBill) -> dict[str, str]:
                pm: dict[str, str] = {}
                for p in b.parts or []:
                    name = (p.name or "").strip().lower()
                    qty = (p.quantity or "").strip()
                    if name:
                        pm[name] = qty
                return pm

            if _parts_map(e_bill) != _parts_map(qb_bill):
                mismatch = True

            if mismatch:
                conflicts.append(
                    Conflict(
                        id=str(e_bill.id or qb_bill.id or k),
                        excel_supplier_name=e_bill.supplier_name,
                        qb_supplier_name=qb_bill.supplier_name,
                        excel_invoice_number=e_bill.invoice_number,
                        qb_invoice_number=qb_bill.invoice_number,
                        excel_invoice_date=e_bill.invoice_date,
                        qb_invoice_date=qb_bill.invoice_date,
                        reason="data_mismatch",
                    )
                )

    # Bills present in QuickBooks but not Excel
    for k, bill in qb_map.items():
        if k not in excel_map:
            qb_only.append(bill)

    return ComparisonReport(excel_only=excel_only, qb_only=qb_only, conflicts=conflicts)


__all__ = ["compare_item_bills"]

if __name__ == "__main__":  # pragma: no cover - manual execution helper
    parser = argparse.ArgumentParser(
        description="Compare Excel and QuickBooks item bills and optionally add Excel-only bills to QuickBooks."
    )
    parser.add_argument(
        "workbook",
        help="Path to Excel workbook containing the account debit vendor worksheet",
    )
    parser.add_argument(
        "--output",
        help="Optional path to write a JSON report (if provided, a basic summary report is generated).",
    )
    args = parser.parse_args()
    workbook_arg = args.workbook
    output_path = args.output

    # Lazy import to avoid circular dependency with qb_gateway
    try:
        from .qb_gateway import read_item_bills, add_item_bills_batch
    except Exception as e:
        print(f"Error importing QuickBooks gateway: {e}")
        raise SystemExit(1)

    # Read from QuickBooks (currently open company file)
    try:
        qb_bills = read_item_bills()
    except Exception as e:
        print(f"Error reading from QuickBooks: {e}")
        raise SystemExit(1)

    print("-- QuickBooks Item Bills --")
    for b in qb_bills:
        print(str(b))

    # Read from Excel
    try:
        excel_bills = extract_item_bills(Path(workbook_arg))
    except Exception as e:
        print(f"Error reading Excel workbook '{workbook_arg}': {e}")
        raise SystemExit(1)

    print("\n-- Excel Item Bills --")
    for b in excel_bills:
        print(str(b))

    # Compare
    comparison = compare_item_bills(excel_bills, qb_bills)

    def _d(v):
        try:
            return v.isoformat()
        except Exception:
            return v

    print("\n-- Comparison Summary --")
    print(f"Excel-only: {len(comparison.excel_only)}")
    for bill in comparison.excel_only:
        print(f"  [EXCEL-ONLY] {bill}")

    print(f"QuickBooks-only: {len(comparison.qb_only)}")
    for bill in comparison.qb_only:
        print(f"  [QB-ONLY] {bill}")

    print(f"Conflicts: {len(comparison.conflicts)}")
    for c in comparison.conflicts:
        inv = c.excel_invoice_number or c.qb_invoice_number
        print(f"  [CONFLICT] {c}")

    """Runner for item bills synchronization."""

    print("\n-- Adding Excel-only bills to QuickBooks --")
    newly_added_bills = add_item_bills_batch(None, comparison.excel_only)
    for added_bill in newly_added_bills:
        print(f"  Added bill: {added_bill}")

    if output_path:
        from datetime import datetime, timezone
        import json

        def _iso(d):
            if d is None:
                return None
            try:
                return d.isoformat()
            except Exception:
                return str(d)

        # Build conflict entries (mismatch + missing cases)
        conflict_entries = []
        for c in comparison.conflicts:
            conflict_entries.append(
                {
                    "record_id": c.id,
                    "reason": c.reason,
                    "excel_supplier_name": c.excel_supplier_name,
                    "qb_supplier_name": c.qb_supplier_name,
                    "excel_invoice_number": c.excel_invoice_number,
                    "qb_invoice_number": c.qb_invoice_number,
                    "excel_invoice_date": _iso(c.excel_invoice_date),
                    "qb_invoice_date": _iso(c.qb_invoice_date),
                }
            )
        for bill in comparison.qb_only:
            conflict_entries.append(
                {
                    "record_id": bill.id or bill.invoice_number,
                    "reason": "missing_in_excel",
                    "excel_supplier_name": None,
                    "qb_supplier_name": bill.supplier_name,
                    "excel_invoice_number": None,
                    "qb_invoice_number": bill.invoice_number,
                    "excel_invoice_date": None,
                    "qb_invoice_date": _iso(bill.invoice_date),
                }
            )
        for bill in comparison.excel_only:
            conflict_entries.append(
                {
                    "record_id": bill.id or bill.invoice_number,
                    "reason": "missing_in_quickbooks",
                    "excel_supplier_name": bill.supplier_name,
                    "qb_supplier_name": None,
                    "excel_invoice_number": bill.invoice_number,
                    "qb_invoice_number": None,
                    "excel_invoice_date": _iso(bill.invoice_date),
                    "qb_invoice_date": None,
                }
            )

        # same_itembills: matched pairs without data_mismatch
        matched_pairs = len(excel_bills) - len(comparison.excel_only)
        same_itembills = matched_pairs - len(comparison.conflicts)
        if same_itembills < 0:
            same_itembills = 0

        added_items_payload = [
            {
                "record_id": b.id or b.invoice_number,
                "supplier_name": b.supplier_name,
                "invoice_number": b.invoice_number,
                "invoice_date": _iso(b.invoice_date),
            }
            for b in newly_added_bills
        ]

        summary = {
            "status": "success",
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "added_itembills": added_items_payload,
            "conflicts": conflict_entries,
            "same_itembills": same_itembills,
            "error": None,
        }
        try:
            with Path(output_path).open("w", encoding="utf-8") as fh:
                json.dump(summary, fh, indent=2)
            print(f"Summary report written to {output_path}")
        except Exception as e:
            print(f"Failed to write summary report to '{output_path}': {e}")
