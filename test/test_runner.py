from __future__ import annotations
from unittest.mock import patch
import quickbook_vendor_item_bills.runner as runner_mod
from quickbook_vendor_item_bills.models import ItemBill

"""Tests for the item bills runner module (coverage only)."""


# 100% coverage: test run_item_bills success and error, and _bill_to_dict


@patch("quickbook_vendor_item_bills.runner.write_report")
@patch("quickbook_vendor_item_bills.runner.excel_reader.extract_item_bills")
def test_run_item_bills_success(mock_extract, mock_write, tmp_path):
    bills = [
        ItemBill(
            supplier_name="A",
            invoice_date="2025-01-01",
            invoice_number=1,
            source="excel",
        )
    ]
    mock_extract.return_value = bills
    out = tmp_path / "r.json"
    result = runner_mod.run_item_bills("", "file.xlsx", output_path=str(out))
    assert result == out
    payload = mock_write.call_args[0][0]
    assert payload["status"] == "success"
    assert payload["added_bills"][0]["supplier_name"] == "A"


@patch("quickbook_vendor_item_bills.runner.write_report")
@patch("quickbook_vendor_item_bills.runner.excel_reader.extract_item_bills")
def test_run_item_bills_error(mock_extract, mock_write, tmp_path):
    mock_extract.side_effect = Exception("fail")
    out = tmp_path / "r.json"
    result = runner_mod.run_item_bills("", "file.xlsx", output_path=str(out))
    assert result == out
    payload = mock_write.call_args[0][0]
    assert payload["status"] == "error"
    assert "fail" in payload["error"]


def test__bill_to_dict_coverage():
    bill = ItemBill(
        supplier_name="B", invoice_date="2025-01-02", invoice_number=2, source="excel"
    )
    d = runner_mod._bill_to_dict(bill)
    assert d["supplier_name"] == "B"
