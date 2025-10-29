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
        txn_id = bill_ret.findtext("TxnID") or ""
        memo = (bill_ret.findtext("Memo") or "").strip()
        _vendor_ref = bill_ret.find("VendorRef")
        supplier_name = ""
        if _vendor_ref is not None:
            supplier_name = (_vendor_ref.findtext("FullName") or "").strip()
        invoice_date = bill_ret.findtext("TimeCreated") or ""
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
                id=(memo or txn_id or None),
            )
        )

    return bills


# def add_payment_terms_batch(
#     company_file: str | None, terms: List[PaymentTerm]
# ) -> List[PaymentTerm]:
#     """Create multiple payment terms in QuickBooks in a single batch request."""

#     if not terms:
#         return []

#     # Build the QBXML with multiple StandardTermsAddRq entries
#     requests = []
#     for term in terms:
#         try:
#             days_value = int(term.record_id)
#         except ValueError as exc:
#             raise ValueError(
#                 f"record_id must be numeric for QuickBooks payment terms: {term.record_id}"
#             ) from exc

#         requests.append(
#             f"    <StandardTermsAddRq>\n"
#             f"      <StandardTermsAdd>\n"
#             f"        <Name>{_escape_xml(term.name)}</Name>\n"
#             f"        <StdDiscountDays>{days_value}</StdDiscountDays>\n"
#             f"        <DiscountPct>0</DiscountPct>\n"
#             f"      </StandardTermsAdd>\n"
#             f"    </StandardTermsAddRq>"
#         )

#     qbxml = (
#         '<?xml version="1.0"?>\n'
#         '<?qbxml version="13.0"?>\n'
#         "<QBXML>\n"
#         '  <QBXMLMsgsRq onError="continueOnError">\n' + "\n".join(requests) + "\n"
#         "  </QBXMLMsgsRq>\n"
#         "</QBXML>"
#     )

#     try:
#         root = _send_qbxml(qbxml)
#     except RuntimeError as exc:
#         # If the entire batch fails, return empty list
#         print(f"Batch add failed: {exc}")
#         return []

#     # Parse all responses
#     added_terms: List[PaymentTerm] = []
#     for term_ret in root.findall(".//StandardTermsRet"):
#         record_id = term_ret.findtext("StdDiscountDays")
#         if not record_id:
#             continue
#         try:
#             record_id = str(int(record_id))
#         except ValueError:
#             record_id = record_id.strip()
#         name = (term_ret.findtext("Name") or "").strip()
#         added_terms.append(
#             PaymentTerm(record_id=record_id, name=name, source="quickbooks")
#         )

#     return added_terms


# def add_item_bill(company_file: str | None, bill: ItemBill) -> ItemBill:
#     """Create an Item Bill in QuickBooks and return the created record.

#     This uses BillAddRq with the following fields from `bill`:
#     - VendorRef/FullName: bill.supplier_name
#     - TxnDate: bill.invoice_date
#     - RefNumber: bill.invoice_number
#     - ItemLineAdd entries for each part with ItemRef/FullName and Quantity
#     """

#     # Validate required fields
#     if not bill.supplier_name or not str(bill.supplier_name).strip():
#         raise ValueError("supplier_name (VendorRef FullName) is required to add a bill")

#     # Build ItemLineAdd entries for parts
#     item_lines = []
#     for p in bill.parts:
#         name = _escape_xml(p.name)
#         qty = _escape_xml(str(p.quantity))
#         item_lines.append(
#             "        <ItemLineAdd>\n"
#             "          <ItemRef>\n"
#             f"            <FullName>{name}</FullName>\n"
#             "          </ItemRef>\n"
#             f"          <Quantity>{qty}</Quantity>\n"
#             "        </ItemLineAdd>\n"
#         )
#     item_lines_xml = "".join(item_lines)

#     vendor_name = _escape_xml(bill.supplier_name)
#     # Normalize date to YYYY-MM-DD if a datetime string is provided
#     date_raw = (bill.invoice_date or "").strip()
#     date_only = date_raw[:10] if date_raw else ""
#     txn_date_line = f"        <TxnDate>{_escape_xml(date_only)}</TxnDate>\n" if date_only else ""
#     ref_number = _escape_xml(str(bill.invoice_number))

#     memo_line = f"        <Memo>{_escape_xml(str(bill.id))}</Memo>\n" if bill.id else ""
#     qbxml = (
#         '<?xml version="1.0"?>\n'
#         '<?qbxml version="16.0"?>\n'
#         "<QBXML>\n"
#         '  <QBXMLMsgsRq onError="stopOnError">\n'
#         "    <BillAddRq>\n"
#         "      <BillAdd>\n"
#         "        <VendorRef>\n"
#         f"          <FullName>{vendor_name}</FullName>\n"
#         "        </VendorRef>\n"
#         f"{txn_date_line}"
#         f"        <RefNumber>{ref_number}</RefNumber>\n"
#         f"{memo_line}"
#         f"{item_lines_xml}"
#         "      </BillAdd>\n"
#         "    </BillAddRq>\n"
#         "  </QBXMLMsgsRq>\n"
#         "</QBXML>"
#     )

#     root = _send_qbxml(qbxml)

#     # Parse BillRet from response
#     bill_ret = root.find(".//BillRet")
#     if bill_ret is None:
#         # If no detailed return, fall back to the input
#         return ItemBill(
#             supplier_name=bill.supplier_name,
#             invoice_date=bill.invoice_date,
#             invoice_number=bill.invoice_number,
#             parts=bill.parts,
#             source="quickbooks",
#         )

#     txn_id = bill_ret.findtext("TxnID") or None
#     memo = (bill_ret.findtext("Memo") or "").strip() or None
#     _vendor_ref = bill_ret.find("VendorRef")
#     out_supplier = ""
#     if _vendor_ref is not None:
#         out_supplier = (_vendor_ref.findtext("FullName") or "").strip()
#     out_invoice_date = bill_ret.findtext("TxnDate") or bill_ret.findtext("TimeCreated") or ""
#     out_invoice_number = bill_ret.findtext("RefNumber") or bill.invoice_number

#     out_parts: List[Part] = []
#     for line_ret in bill_ret.findall(".//ItemLineRet"):
#         item_ref = line_ret.find("ItemRef")
#         part_name = ""
#         if item_ref is not None:
#             part_name = (item_ref.findtext("FullName") or "").strip()
#         quantity = line_ret.findtext("Quantity") or ""
#         if part_name:
#             out_parts.append(Part(name=part_name, quantity=quantity))

#     return ItemBill(
#         supplier_name=out_supplier,
#         invoice_date=out_invoice_date,
#         invoice_number=out_invoice_number,
#         parts=out_parts,
#         source="quickbooks",
#         id=(memo or txn_id),
#     )


def _escape_xml(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


__all__ = ["fetch_item_bills"]
