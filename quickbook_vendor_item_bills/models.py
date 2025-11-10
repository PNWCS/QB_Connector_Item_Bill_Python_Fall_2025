"""Domain models for item bill comparison and synchronisation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, List
from datetime import date


SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal[
    "supplier_name_mismatch",
    "invoice_date_mismatch",
    "invoice_number_mismatch",
    "part_mismatch",
    "missing_in_excel",
    "missing_in_quickbooks",
]


@dataclass(slots=True)
class Part:
    """Represents a part of an item bill."""

    name: str
    quantity: str

    def __str__(self) -> str:
        return f"Part(name='{self.name}', quantity='{self.quantity}')"


@dataclass(slots=True)
class ItemBill:
    """Represents an item bill synchronised between Excel and QuickBooks."""

    supplier_name: str
    invoice_date: date | None
    invoice_number: str
    source: SourceLiteral
    parts: list[Part] = field(default_factory=list)
    id: str | None = None

    def __str__(self) -> str:
        parts_str = ", ".join(str(p) for p in self.parts) if self.parts else ""
        return (
            "ItemBill("
            f"supplier_name='{self.supplier_name}', "
            f"invoice_date='{self.invoice_date.isoformat() if self.invoice_date else ''}', "
            f"invoice_number='{self.invoice_number}', "
            f"parts=[{parts_str}], "
            f"source='{self.source}', "
            f"id='{self.id}'"
            ")"
        )


@dataclass(slots=True)
class Conflict:
    """Describes a discrepancy between Excel and QuickBooks item bills."""

    id: str
    excel_supplier_name: str | None
    qb_supplier_name: str | None
    excel_invoice_number: str | None
    qb_invoice_number: str | None
    excel_invoice_date: date | None
    qb_invoice_date: date | None
    reason: List[ConflictReason]

    def __str__(self) -> str:
        return (
            "Conflict("
            f"id='{self.id}', "
            f"excel_supplier_name='{self.excel_supplier_name}', "
            f"qb_supplier_name='{self.qb_supplier_name}', "
            f"excel_invoice_number='{self.excel_invoice_number}', "
            f"qb_invoice_number='{self.qb_invoice_number}', "
            f"excel_invoice_date='{self.excel_invoice_date.isoformat() if self.excel_invoice_date else ''}', "
            f"qb_invoice_date='{self.qb_invoice_date.isoformat() if self.qb_invoice_date else ''}', "
            f"reason={self.reason}"
            ")"
        )


@dataclass(slots=True)
class ComparisonReport:
    """Groups comparison outcomes for later processing."""

    excel_only: List[ItemBill] = field(default_factory=list)
    qb_only: List[ItemBill] = field(default_factory=list)
    conflicts: List[Conflict] = field(default_factory=list)


__all__ = [
    "Part",
    "ItemBill",
    "Conflict",
    "ComparisonReport",
    "ConflictReason",
    "SourceLiteral",
]
