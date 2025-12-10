"""QuickBooks COM gateway helpers for item bills."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

from .models import ItemBill, Part
from .excel_reader import extract_item_bills
from .comparer import compare_item_bills
from pathlib import Path
import sys
from datetime import datetime, date


APP_NAME = "Quickbooks Connector"  # do not change this


def _require_win32com() -> None:
    if win32com is None:  # pragma: no cover - exercised via tests
        raise RuntimeError("pywin32 is required to communicate with QuickBooks")


@contextmanager
def _qb_session() -> Iterator[tuple[object, object]]:
    _require_win32com()
    session = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    session.OpenConnection2("", APP_NAME, 1)
    ticket = session.BeginSession("", 0)
    try:
        yield session, ticket
    finally:
        try:
            session.EndSession(ticket)
        finally:
            session.CloseConnection()


def _send_qbxml(qbxml: str) -> ET.Element:
    with _qb_session() as (session, ticket):
        print(f"Sending QBXML:\n{qbxml}")  # Debug output
        raw_response = session.ProcessRequest(ticket, qbxml)  # type: ignore[attr-defined]
        print(f"Received response:\n{raw_response}")  # Debug output
    return _parse_response(raw_response)


def _parse_response(raw_xml: str) -> ET.Element:
    print(f"Parsing QuickBooks response XML: {raw_xml}")  # Debug output
    root = ET.fromstring(raw_xml)
    response = root.find(".//*[@statusCode]")
    if response is None:
        raise RuntimeError("QuickBooks response missing status information")

    status_code = int(response.get("statusCode", "0"))
    status_message = response.get("statusMessage", "")
    # Status code 1 means "no matching objects found" - this is OK for queries
    if status_code != 0 and status_code != 1:
        print(f"QuickBooks error ({status_code}): {status_message}")
        raise RuntimeError(status_message)
    return root


def fetch_item_bills(company_file: str | None = None) -> List[ItemBill]:
    """Return item bills currently stored in QuickBooks."""

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillQueryRq>\n"
        "      <IncludeLineItems>true</IncludeLineItems>\n"
        "    </BillQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )
    root = _send_qbxml(qbxml)
    bills: List[ItemBill] = []
    for bill_ret in root.findall(".//BillRet"):
        memo = (bill_ret.findtext("Memo") or "").strip()
        _vendor_ref = bill_ret.find("VendorRef")
        supplier_name = ""
        if _vendor_ref is not None:
            supplier_name = (_vendor_ref.findtext("FullName") or "").strip()
        time_created = bill_ret.findtext("TxnDate") or ""
        try:
            invoice_date: date | None = (
                datetime.fromisoformat(time_created).date() if time_created else None
            )
        except Exception:
            invoice_date = None
        invoice_number = bill_ret.findtext("RefNumber") or ""

        if not invoice_number:
            continue

        parts: List[Part] = []
        for part in bill_ret.findall(".//ItemLineRet"):
            item_ref = part.find("ItemRef")
            part_name = ""
            if item_ref is not None:
                part_name = (item_ref.findtext("FullName") or "").strip()
            quantity = part.findtext("Quantity") or ""
            if part_name:
                parts.append(Part(name=part_name, quantity=quantity))

        bills.append(
            ItemBill(
                supplier_name=supplier_name,
                invoice_date=invoice_date,
                invoice_number=invoice_number,
                parts=parts,
                source="quickbooks",
                id=(memo or None),
            )
        )

    return bills


def read_item_bills() -> List[ItemBill]:
    """Read bills from the currently open QuickBooks company file.

    This function takes no arguments, per assignment requirements. It uses an
    empty company file path to instruct QuickBooks to use the currently open
    company file/session.
    """

    return fetch_item_bills("")


def add_item_bill(company_file: str | None, bill: ItemBill) -> ItemBill:
    """Create an Item Bill in QuickBooks and return the created record.

    This uses BillAddRq with the following fields from `bill`:
    - VendorRef/FullName: bill.supplier_name
    - TxnDate: bill.invoice_date
    - RefNumber: bill.invoice_number
    - ItemLineAdd entries for each part with ItemRef/FullName and Quantity
    """

    # Validate required fields
    if not bill.supplier_name or not str(bill.supplier_name).strip():
        raise ValueError("supplier_name (VendorRef FullName) is required to add a bill")
    if not bill.invoice_number or not str(bill.invoice_number).strip():
        raise ValueError("invoice_number (RefNumber) is required to add a bill")

    # Build ItemLineAdd entries for parts
    item_lines = []
    for p in bill.parts:
        name = _escape_xml(p.name)
        qty = _escape_xml(str(p.quantity))
        item_lines.append(
            "        <ItemLineAdd>\n"
            "          <ItemRef>\n"
            f"            <FullName>{name}</FullName>\n"
            "          </ItemRef>\n"
            f"          <Quantity>{qty}</Quantity>\n"
            "        </ItemLineAdd>\n"
        )
    item_lines_xml = "".join(item_lines)

    vendor_name = _escape_xml(bill.supplier_name)
    # Use ISO date (YYYY-MM-DD) for TxnDate when available
    date_only = bill.invoice_date.isoformat() if bill.invoice_date else ""
    txn_date_line = (
        f"        <TxnDate>{_escape_xml(date_only)}</TxnDate>\n" if date_only else ""
    )
    ref_number = _escape_xml(str(bill.invoice_number))

    memo_line = f"        <Memo>{_escape_xml(str(bill.id))}</Memo>\n" if bill.id else ""
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillAddRq>\n"
        "      <BillAdd>\n"
        "        <VendorRef>\n"
        f"          <FullName>{vendor_name}</FullName>\n"
        "        </VendorRef>\n"
        f"{txn_date_line}"
        f"        <RefNumber>{ref_number}</RefNumber>\n"
        f"{memo_line}"
        f"{item_lines_xml}"
        "      </BillAdd>\n"
        "    </BillAddRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    root = _send_qbxml(qbxml)

    # Parse BillRet from response
    bill_ret = root.find(".//BillRet")
    if bill_ret is None:
        # If no detailed return, fall back to the input
        return ItemBill(
            supplier_name=bill.supplier_name,
            invoice_date=bill.invoice_date,
            invoice_number=bill.invoice_number,
            parts=bill.parts,
            source="quickbooks",
        )

    memo = (bill_ret.findtext("Memo") or "").strip() or None
    _vendor_ref = bill_ret.find("VendorRef")
    out_supplier = ""
    if _vendor_ref is not None:
        out_supplier = (_vendor_ref.findtext("FullName") or "").strip()
    # Parse invoice date back to date object if possible
    out_invoice_date_str = bill_ret.findtext("TxnDate")
    out_invoice_date: date | None
    if out_invoice_date_str:
        try:
            out_invoice_date = date.fromisoformat(out_invoice_date_str)
        except Exception:
            out_invoice_date = None
    else:
        time_created_str = bill_ret.findtext("TimeCreated") or ""
        if time_created_str:
            try:
                out_invoice_date = datetime.fromisoformat(time_created_str).date()
            except Exception:
                out_invoice_date = None
        else:
            out_invoice_date = None
    out_invoice_number = bill_ret.findtext("RefNumber") or bill.invoice_number

    out_parts: List[Part] = []
    for line_ret in bill_ret.findall(".//ItemLineRet"):
        item_ref = line_ret.find("ItemRef")
        part_name = ""
        if item_ref is not None:
            part_name = (item_ref.findtext("FullName") or "").strip()
        quantity = line_ret.findtext("Quantity") or ""
        if part_name:
            out_parts.append(Part(name=part_name, quantity=quantity))

    return ItemBill(
        supplier_name=out_supplier,
        invoice_date=out_invoice_date,
        invoice_number=out_invoice_number,
        parts=out_parts,
        source="quickbooks",
        id=memo,
    )


def _escape_xml(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def add_item_bills_batch(
    company_file: str | None, bills: List[ItemBill]
) -> List[ItemBill]:
    """Create multiple Item Bills in a single QuickBooks request.

    Returns the list of created ItemBill objects as parsed from the response.
    If the request fails entirely, a RuntimeError will be raised by _send_qbxml.
    """

    if not bills:
        return []

    def _bill_add_xml(bill: ItemBill) -> str:
        if not bill.supplier_name or not str(bill.supplier_name).strip():
            raise ValueError("supplier_name is required to add a bill")
        if not bill.invoice_number or not str(bill.invoice_number).strip():
            raise ValueError("invoice_number is required to add a bill")

        vendor_name = _escape_xml(bill.supplier_name)
        date_only = bill.invoice_date.isoformat() if bill.invoice_date else ""
        txn_date_line = (
            f"        <TxnDate>{_escape_xml(date_only)}</TxnDate>\n"
            if date_only
            else ""
        )
        ref_number = _escape_xml(str(bill.invoice_number))
        memo_line = (
            f"        <Memo>{_escape_xml(str(bill.id))}</Memo>\n" if bill.id else ""
        )

        item_lines = []
        for p in bill.parts:
            name = _escape_xml(p.name)
            qty = _escape_xml(str(p.quantity))
            item_lines.append(
                "        <ItemLineAdd>\n"
                "          <ItemRef>\n"
                f"            <FullName>{name}</FullName>\n"
                "          </ItemRef>\n"
                f"          <Quantity>{qty}</Quantity>\n"
                "        </ItemLineAdd>\n"
            )
        item_lines_xml = "".join(item_lines)

        return (
            "    <BillAddRq>\n"
            "      <BillAdd>\n"
            "        <VendorRef>\n"
            f"          <FullName>{vendor_name}</FullName>\n"
            "        </VendorRef>\n"
            f"{txn_date_line}"
            f"        <RefNumber>{ref_number}</RefNumber>\n"
            f"{memo_line}"
            f"{item_lines_xml}"
            "      </BillAdd>\n"
            "    </BillAddRq>"
        )

    requests = [_bill_add_xml(b) for b in bills]
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n' + "\n".join(requests) + "\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    root = _send_qbxml(qbxml)

    created: List[ItemBill] = []
    for bill_ret in root.findall(".//BillRet"):
        memo = (bill_ret.findtext("Memo") or "").strip() or None
        _vendor_ref = bill_ret.find("VendorRef")
        out_supplier = ""
        if _vendor_ref is not None:
            out_supplier = (_vendor_ref.findtext("FullName") or "").strip()
        out_invoice_number = (bill_ret.findtext("RefNumber") or "").strip()

        out_date: date | None = None
        out_date_str = (
            bill_ret.findtext("TxnDate") or bill_ret.findtext("TimeCreated") or ""
        )
        if out_date_str:
            try:
                out_date = date.fromisoformat(out_date_str)
            except Exception:
                out_date = None

        out_parts: List[Part] = []
        for line_ret in bill_ret.findall(".//ItemLineRet"):
            item_ref = line_ret.find("ItemRef")
            part_name = ""
            if item_ref is not None:
                part_name = (item_ref.findtext("FullName") or "").strip()
            quantity = line_ret.findtext("Quantity") or ""
            if part_name:
                out_parts.append(Part(name=part_name, quantity=quantity))

        created.append(
            ItemBill(
                supplier_name=out_supplier,
                invoice_date=out_date,
                invoice_number=out_invoice_number,
                parts=out_parts,
                source="quickbooks",
                id=memo,
            )
        )

    return created


__all__ = [
    "fetch_item_bills",
    "read_item_bills",
    "add_item_bill",
    "add_item_bills_batch",
]


if __name__ == "__main__":  # pragma: no cover - manual execution helper
    # Usage: python -m quickbook_vendor_item_bills.qb_gateway <workbook_path>
    if len(sys.argv) < 2:
        print("Usage: python -m quickbook_vendor_item_bills.qb_gateway <workbook_path>")
        raise SystemExit(2)

    workbook_arg = sys.argv[1]

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
