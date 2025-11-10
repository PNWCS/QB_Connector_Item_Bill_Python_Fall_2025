"""Comparison helpers for item bills."""

from __future__ import annotations

from typing import Iterable

from .models import ComparisonReport, Conflict, ConflictReason, ItemBill


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

            reasons: list[ConflictReason] = []
            # supplier name mismatch
            if (e_bill.supplier_name or "") != (qb_bill.supplier_name or ""):
                reasons.append("supplier_name_mismatch")
            # invoice date mismatch (date objects or None)
            if e_bill.invoice_date != qb_bill.invoice_date:
                reasons.append("invoice_date_mismatch")
            # invoice number mismatch
            if (e_bill.invoice_number or "") != (qb_bill.invoice_number or ""):
                reasons.append("invoice_number_mismatch")

            # parts mismatch (compare as name->quantity dicts, case-insensitive names)
            def _parts_map(b: ItemBill) -> dict[str, str]:
                pm: dict[str, str] = {}
                for p in b.parts or []:
                    name = (p.name or "").strip().lower()
                    qty = (p.quantity or "").strip()
                    if name:
                        pm[name] = qty
                return pm

            if _parts_map(e_bill) != _parts_map(qb_bill):
                reasons.append("part_mismatch")

            if reasons:
                conflicts.append(
                    Conflict(
                        id=str(e_bill.id or qb_bill.id or k),
                        excel_supplier_name=e_bill.supplier_name,
                        qb_supplier_name=qb_bill.supplier_name,
                        excel_invoice_number=e_bill.invoice_number,
                        qb_invoice_number=qb_bill.invoice_number,
                        excel_invoice_date=e_bill.invoice_date,
                        qb_invoice_date=qb_bill.invoice_date,
                        reason=reasons,
                    )
                )

    # Bills present in QuickBooks but not Excel
    for k, bill in qb_map.items():
        if k not in excel_map:
            qb_only.append(bill)

    return ComparisonReport(excel_only=excel_only, qb_only=qb_only, conflicts=conflicts)


__all__ = ["compare_item_bills"]
