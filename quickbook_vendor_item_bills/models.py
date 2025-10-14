"""Domain models for payment term synchronisation."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["name_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class ItemBill:
    """Represents a item bill synchronised between Excel and QuickBooks."""

    supplier_name: str
    invoice_date: str  # ISO 8601 date string
    # part: str # From Bills sheet
    # quantity: str # From Bills sheet
    invoice_number: int
    source: SourceLiteral


# @dataclass(slots=True)
# class Conflict:
#     """Describes a discrepancy between Excel and QuickBooks payment terms."""

#     record_id: str
#     excel_name: str | None
#     qb_name: str | None
#     reason: ConflictReason


# @dataclass(slots=True)
# class ComparisonReport:
#     """Groups comparison outcomes for later processing."""

#     excel_only: list[ItemBill] = field(default_factory=list)
#     qb_only: list[ItemBill] = field(default_factory=list)
#     conflicts: list[Conflict] = field(default_factory=list)


__all__ = [
    "ItemBill",
    # "Conflict",
    # "ComparisonReport",
    "ConflictReason",
    "SourceLiteral",
]
