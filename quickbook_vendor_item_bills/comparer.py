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
    excel_map = {bill.invoice_number: bill for bill in excel_bills}
    qb_map = {bill.invoice_number: bill for bill in qb_bills}

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
                        invoice_number=inv,
                        excel_supplier=bill.supplier_name,
                        qb_supplier=qb_bill.supplier_name,
                        excel_date=bill.invoice_date,
                        qb_date=qb_bill.invoice_date,
                        reason="supplier_mismatch",
                    )
                )
            # date mismatch (allow empty strings)
            if (bill.invoice_date or "") != (qb_bill.invoice_date or ""):
                conflicts.append(
                    Conflict(
                        invoice_number=inv,
                        excel_supplier=bill.supplier_name,
                        qb_supplier=qb_bill.supplier_name,
                        excel_date=bill.invoice_date,
                        qb_date=qb_bill.invoice_date,
                        reason="date_mismatch",
                    )
                )

    # Bills present in QuickBooks but not Excel
    for inv, bill in qb_map.items():
        if inv not in excel_map:
            qb_only.append(bill)

    return ComparisonReport(excel_only=excel_only, qb_only=qb_only, conflicts=conflicts)


__all__ = ["compare_item_bills"]
