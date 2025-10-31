"""Comparison helpers for item bills."""

from __future__ import annotations

from typing import Iterable

from .models import ComparisonReport, Conflict, ItemBill


def compare_item_bills(
    excel_bills: Iterable[ItemBill],
    qb_bills: Iterable[ItemBill],
) -> ComparisonReport:
    """Compare Excel and QuickBooks Item bills and identify discrepancies."""
    # Build maps keyed by invoice_number for quick lookup
    excel_map = {bill.id: bill for bill in excel_bills}
    qb_map = {bill.id: bill for bill in qb_bills}

    print("Debug: Comparing item bills")

    for inv, bill in excel_map.items():
        print(f"{inv}: {bill}")

    print("Debug: Comparing item bills")

    for inv, bill in qb_map.items():
        print(f"{inv}: {bill}")

    excel_only: list[ItemBill] = []
    qb_only: list[ItemBill] = []
    conflicts: list[Conflict] = []

    # Bills present in Excel but not QuickBooks
    for inv, bill in excel_map.items():
        if inv not in qb_map:
            excel_only.append(bill)
        else:
            qb_bill = qb_map[inv]
            # supplier mismatch
            if (bill.supplier_name or "") != (qb_bill.supplier_name or ""):
                conflicts.append(
                    Conflict(
                        id=str(bill.id or qb_bill.id or inv),
                        excel_supplier_name=bill.supplier_name,
                        qb_supplier_name=qb_bill.supplier_name,
                        excel_invoice_number=bill.invoice_number,
                        qb_invoice_number=qb_bill.invoice_number,
                        excel_invoice_date=bill.invoice_date,
                        qb_invoice_date=qb_bill.invoice_date,
                        reason="supplier_name_mismatch",
                    )
                )
            # invoice date mismatch (date objects or None)
            if bill.invoice_date != qb_bill.invoice_date:
                conflicts.append(
                    Conflict(
                        id=str(bill.id or qb_bill.id or inv),
                        excel_supplier_name=bill.supplier_name,
                        qb_supplier_name=qb_bill.supplier_name,
                        excel_invoice_number=bill.invoice_number,
                        qb_invoice_number=qb_bill.invoice_number,
                        excel_invoice_date=bill.invoice_date,
                        qb_invoice_date=qb_bill.invoice_date,
                        reason="invoice_date_mismatch",
                    )
                )

    # Bills present in QuickBooks but not Excel
    for inv, bill in qb_map.items():
        if inv not in excel_map:
            qb_only.append(bill)

    return ComparisonReport(excel_only=excel_only, qb_only=qb_only, conflicts=conflicts)


__all__ = ["compare_item_bills"]
