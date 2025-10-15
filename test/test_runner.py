"""Tests for the item bills runner module."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest

from quickbook_vendor_item_bills.models import ComparisonReport, Conflict, ItemBill
from quickbook_vendor_item_bills.runner import run_item_bills


@pytest.fixture
def mock_excel_bills():
    """Mock Excel item bills."""
    return [
        ItemBill(
            supplier_name="A",
            invoice_date="2025-01-01",
            invoice_number=1,
            parts=[],
            source="excel",
        ),
        ItemBill(
            supplier_name="B",
            invoice_date="2025-01-02",
            invoice_number=2,
            parts=[],
            source="excel",
        ),
        ItemBill(
            supplier_name="C",
            invoice_date="2025-01-03",
            invoice_number=3,
            parts=[],
            source="excel",
        ),
    ]


@pytest.fixture
def mock_qb_bills():
    """Mock QuickBooks item bills."""
    return [
        ItemBill(
            supplier_name="X",
            invoice_date="2025-01-10",
            invoice_number=10,
            parts=[],
            source="quickbooks",
        ),
        ItemBill(
            supplier_name="Y",
            invoice_date="2025-01-11",
            invoice_number=11,
            parts=[],
            source="quickbooks",
        ),
    ]


@pytest.fixture
def mock_comparison_no_conflicts():
    """Mock comparison with no conflicts - only new bills to add."""
    comparison = ComparisonReport()
    comparison.excel_only = [
        ItemBill(
            supplier_name="A",
            invoice_date="2025-01-01",
            invoice_number=1,
            parts=[],
            source="excel",
        ),
        ItemBill(
            supplier_name="B",
            invoice_date="2025-01-02",
            invoice_number=2,
            parts=[],
            source="excel",
        ),
        ItemBill(
            supplier_name="C",
            invoice_date="2025-01-03",
            invoice_number=3,
            parts=[],
            source="excel",
        ),
    ]
    comparison.qb_only = [
        ItemBill(
            supplier_name="X",
            invoice_date="2025-01-10",
            invoice_number=10,
            parts=[],
            source="quickbooks",
        ),
        ItemBill(
            supplier_name="Y",
            invoice_date="2025-01-11",
            invoice_number=11,
            parts=[],
            source="quickbooks",
        ),
    ]
    comparison.conflicts = []
    return comparison


@pytest.fixture
def mock_comparison_with_conflicts():
    """Mock comparison with name mismatch conflicts."""
    comparison = ComparisonReport()
    comparison.excel_only = [
        ItemBill(
            supplier_name="A",
            invoice_date="2025-01-01",
            invoice_number=1,
            parts=[],
            source="excel",
        ),
    ]
    comparison.qb_only = []
    comparison.conflicts = [
        Conflict(
            invoice_number=2,
            excel_supplier="B",
            qb_supplier="B_Changed",
            excel_date="2025-01-02",
            qb_date="2025-01-02",
            reason="supplier_mismatch",
        )
    ]
    return comparison


class TestRunItemBills:
    """Test suite for run_item_bills function."""

    @patch("quickbook_vendor_item_bills.excel_reader.extract_item_bills")
    @patch("quickbook_vendor_item_bills.qb_gateway.fetch_item_bills")
    @patch("quickbook_vendor_item_bills.comparer.compare_item_bills")
    # @patch("quickbook_vendor_item_bills.runner.qb_gateway.add_item_bills_batch")
    @patch("quickbook_vendor_item_bills.runner.write_report")
    def test_successful_sync_no_conflicts(
        self,
        mock_write_report,
        # mock_add_batch,
        mock_compare,
        mock_fetch_qb,
        mock_extract_excel,
        mock_excel_bills,
        mock_qb_bills,
        mock_comparison_no_conflicts,
        tmp_path,
    ):
        """Test successful synchronization with no conflicts."""
        # Arrange
        workbook_path = "test_workbook.xlsx"
        output_path = tmp_path / "report.json"

        mock_extract_excel.return_value = mock_excel_bills
        mock_fetch_qb.return_value = mock_qb_bills
        mock_compare.return_value = mock_comparison_no_conflicts

        # added_bills = [
        #     ItemBill(supplier_name="A", invoice_date="2025-01-01", invoice_number=15, source="quickbooks"),
        #     ItemBill(supplier_name="B", invoice_date="2025-01-02", invoice_number=45, source="quickbooks"),
        # ]
        # mock_add_batch.return_value = added_bills

        # Act
        result_path = run_item_bills("", workbook_path, output_path=str(output_path))

        # Assert
        assert result_path == output_path
        mock_extract_excel.assert_called_once_with(Path(workbook_path))
        mock_fetch_qb.assert_called_once_with("")
        mock_compare.assert_called_once_with(mock_excel_bills, mock_qb_bills)
        # mock_add_batch.assert_called_once_with(
        #     "", mock_comparison_no_conflicts.excel_only
        # )

        # Verify report payload
        report_call = mock_write_report.call_args[0]
        report_payload = report_call[0]

        assert report_payload["status"] == "success"
        # assert len(report_payload["added_terms"]) == 2
        # assert report_payload["added_terms"][0]["record_id"] == "15"
        # assert report_payload["added_terms"][1]["record_id"] == "45"
        # assert len(report_payload["conflicts"]) == 1  # One missing_in_excel conflict
        # assert report_payload["conflicts"][0]["reason"] == "missing_in_excel"
        assert report_payload["error"] is None
