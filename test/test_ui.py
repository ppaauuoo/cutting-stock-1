"""
Unit tests for the CuttingOptimizerUI class in cuttingstock/ui.py.
Focuses on testing the UI's interaction with core logic.
"""

import sys
from unittest.mock import patch

import pytest
import polars as pl
from PyQt5.QtWidgets import QApplication, QMessageBox

from cuttingstock.ui import CuttingOptimizerUI


@pytest.fixture(scope="session")
def qt_app():
    """
    Creates a single QApplication instance for the entire test session.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def ui(qt_app, monkeypatch):
    """
    Fixture to create a CuttingOptimizerUI instance for testing.
    """
    # Prevent QMessageBox from blocking tests
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: None)
    monkeypatch.setattr(QMessageBox, "information", lambda *args: None)
    monkeypatch.setattr(QMessageBox, "critical", lambda *args: None)

    window = CuttingOptimizerUI()
    yield window
    window.close()


@patch('cuttingstock.ui.generate_suggestions')
def test_get_all_suggestions_calls_core_function(mock_generate_suggestions, ui):
    """Test that the UI method calls the core generate_suggestions function."""
    mock_generate_suggestions.return_value = [{'width': '80', 'spec': {}}]
    ui.cleaned_orders_df = pl.DataFrame({"order_number": ["1"]})
    ui.ROLL_SPECS = {"80": {}}
    ui.factory_combo.setCurrentText("รวม")

    suggestions = ui.get_all_suggestions()

    assert suggestions == [{'width': '80', 'spec': {}}]
    mock_generate_suggestions.assert_called_once_with(
        ui.cleaned_orders_df, ui.ROLL_SPECS, "รวม"
    )

@patch('cuttingstock.ui.generate_suggestions')
def test_get_all_suggestions_handles_exception_from_core(mock_generate_suggestions, ui):
    """Test that the UI method catches exceptions from the core function."""
    mock_generate_suggestions.side_effect = Exception("Core Error")
    ui.cleaned_orders_df = pl.DataFrame({"order_number": ["1"]})
    ui.ROLL_SPECS = {"80": {}}
    ui.factory_combo.setCurrentText("รวม")

    suggestions = ui.get_all_suggestions()

    assert suggestions == []
    # In a real app, we'd also check if QMessageBox.critical was called.
    # The monkeypatch for this test just suppresses it.

def test_get_all_suggestions_handles_no_orders_gracefully(ui):
    """Test that no suggestions are generated when there is no order data, without calling core logic."""
    ui.cleaned_orders_df = None
    suggestions = ui.get_all_suggestions()
    assert suggestions == []

    ui.cleaned_orders_df = pl.DataFrame()
    suggestions = ui.get_all_suggestions()
    assert suggestions == []
