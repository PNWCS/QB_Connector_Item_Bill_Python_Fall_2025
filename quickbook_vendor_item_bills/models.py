"""Domain models for item bill comparison and synchronisation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, List


SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal[
    "supplier_mismatch",
    "date_mismatch",
    "missing_in_excel",
    "missing_in_quickbooks",
]


@dataclass(slots=True)
class Part:
    """Represents a part of an item bill."""

    name: str
    quantity: str


@dataclass(slots=True)
class ItemBill:
    """Represents an item bill synchronised between Excel and QuickBooks."""

    supplier_name: str
    invoice_date: str
    parts: list[Part]
    invoice_number: str | int
    source: SourceLiteral


@dataclass(slots=True)
class Conflict:
    """Describes a discrepancy between Excel and QuickBooks item bills."""

    invoice_number: str | int
    excel_supplier: str | None
    qb_supplier: str | None
    excel_date: str | None
    qb_date: str | None
    reason: ConflictReason


@dataclass(slots=True)
class ComparisonReport:
    """Groups comparison outcomes for later processing."""

    excel_only: List[ItemBill] = field(default_factory=list)
    qb_only: List[ItemBill] = field(default_factory=list)
    conflicts: List[Conflict] = field(default_factory=list)


__all__ = [
    "ItemBill",
    "Conflict",
    "ComparisonReport",
    "ConflictReason",
    "SourceLiteral",
]
