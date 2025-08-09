"""
Unit tests for the CuttingOptimizerUI class in cuttingstock/ui.py.
Focuses on testing the non-GUI logic, particularly suggestion generation.
"""

import sys
import pytest
import polars as pl
from PyQt5.QtWidgets import QApplication, QMessageBox

from cuttingstock.ui import CuttingOptimizerUI


@pytest.fixture(scope="session")
def qt_app():
    """
    Creates a single QApplication instance for the entire test session.
    This is necessary for instantiating any Qt-based widgets, including the main window.
    """
    app = QApplication.instance()
    if app is None:
        # Pass an empty list to avoid parsing command-line arguments
        app = QApplication([])
    return app


@pytest.fixture
def ui(qt_app, monkeypatch):
    """
    Fixture to create a CuttingOptimizerUI instance for testing.
    - Mocks QMessageBox to prevent popups from blocking automated test execution.
    - The UI's __init__ starts background threads for file watching. These will
      log errors about missing files but won't affect the test logic since we
      manually provide the necessary dataframes.
    - The window is closed after the test, which triggers the closeEvent handler
      to gracefully stop the background threads.
    """
    # Prevent QMessageBox from blocking tests
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: None)
    monkeypatch.setattr(QMessageBox, "information", lambda *args: None)
    monkeypatch.setattr(QMessageBox, "critical", lambda *args: None)

    window = CuttingOptimizerUI()
    yield window
    window.close()


def test_get_all_suggestions_no_orders(ui):
    """Test that no suggestions are generated when there is no order data."""
    ui.cleaned_orders_df = pl.DataFrame()
    ui.ROLL_SPECS = {
        "80": {"M1": {1: {"id": "R1", "length": 1000}}}
    }
    ui.factory_combo.setCurrentText("รวม")

    suggestions = ui.get_all_suggestions()
    assert suggestions == []

    # Test with None dataframe
    ui.cleaned_orders_df = None
    suggestions = ui.get_all_suggestions()
    assert suggestions == []


def test_get_all_suggestions_no_stock(ui):
    """Test that no suggestions are generated when there is no stock data."""
    ui.cleaned_orders_df = pl.DataFrame({
        "order_number": ["1"], "front": ["M1"], "c": [None],
        "middle": [None], "b": [None], "back": [None],
    })
    ui.ROLL_SPECS = {}
    ui.factory_combo.setCurrentText("รวม")

    suggestions = ui.get_all_suggestions()
    assert suggestions == []


def test_get_all_suggestions_simple_case(ui):
    """Test a basic case with one matching suggestion."""
    ui.cleaned_orders_df = pl.DataFrame({
        "order_number": ["1"],
        "front": ["M1"], "c": ["C1"], "middle": [None], "b": [None], "back": ["M2"],
    })
    ui.ROLL_SPECS = {
        "80": {
            "M1": {1: {"id": "R1", "length": 1000}},
            "C1": {1: {"id": "R2", "length": 1000}},
            "M2": {1: {"id": "R3", "length": 1000}},
        },
        "90": {  # Incomplete stock for this width
            "M1": {1: {"id": "R4", "length": 1000}},
            "C1": {1: {"id": "R5", "length": 1000}},
        }
    }
    ui.factory_combo.setCurrentText("รวม")

    suggestions = ui.get_all_suggestions()

    assert len(suggestions) == 1
    assert suggestions[0]['width'] == "80"
    assert suggestions[0]['spec'] == {'front': 'M1', 'c': 'C1', 'middle': '', 'b': '', 'back': 'M2'}


def test_get_all_suggestions_factory_filter(ui):
    """Test that suggestions are correctly filtered by factory selection."""
    ui.cleaned_orders_df = pl.DataFrame({
        "order_number": ["1218001", "30002"],  # Factories 1&2, 3
        "front": ["M1", "M3"],
        "c": ["C1", "C3"],
        "middle": [None, None], "b": [None, None], "back": [None, None],
    })
    ui.ROLL_SPECS = {
        "80": {"M1": {1: {}}, "C1": {1: {}}},
        "90": {"M3": {1: {}}, "C3": {1: {}}},
    }

    # Test for factory "1&2"
    ui.factory_combo.setCurrentText("1&2")
    suggestions = ui.get_all_suggestions()
    assert len(suggestions) == 1
    assert suggestions[0]['width'] == '80'

    # Test for factory "3"
    ui.factory_combo.setCurrentText("3")
    suggestions = ui.get_all_suggestions()
    assert len(suggestions) == 1
    assert suggestions[0]['width'] == '90'

    # Test for factory "4" (no orders for this factory)
    ui.factory_combo.setCurrentText("4")
    suggestions = ui.get_all_suggestions()
    assert len(suggestions) == 0


def test_get_all_suggestions_sorting_default(ui):
    """Test default sorting of suggestions by width (as integer)."""
    ui.cleaned_orders_df = pl.DataFrame({
        "order_number": ["1"],
        "front": ["M1"], "c": [None], "middle": [None], "b": [None], "back": [None]
    })
    ui.ROLL_SPECS = {
        "100": {"M1": {1: {"id": "R1", "length": 1000}}},
        "80": {"M1": {1: {"id": "R2", "length": 1000}}},
        "90": {"M1": {1: {"id": "R3", "length": 1000}}},
    }
    ui.factory_combo.setCurrentText("รวม")

    suggestions = ui.get_all_suggestions()

    assert len(suggestions) == 3
    assert [s['width'] for s in suggestions] == ['80', '90', '100']


def test_get_all_suggestions_sorting_factory_1_2(ui):
    """Test special sorting logic for factory '1&2'."""
    ui.cleaned_orders_df = pl.DataFrame({
        "order_number": ["1218001"],
        "front": ["M1"], "c": [None], "middle": [None], "b": [None], "back": [None]
    })
    # Filter orders first, so only 1218001 is considered
    ui.factory_combo.setCurrentText("1&2")

    ui.ROLL_SPECS = {
        "75": {"M1": {1: {"id": "R1", "length": 1000}}},  # Group 1
        "85": {"M1": {1: {"id": "R2", "length": 1000}}},  # Group 0
        "95": {"M1": {1: {"id": "R3", "length": 1000}}},  # Group 0
        "100": {"M1": {1: {"id": "R4", "length": 1000}}}, # Group 2
        "78": {"M1": {1: {"id": "R5", "length": 1000}}},  # Group 1
        "82": {"M1": {1: {"id": "R6", "length": 1000}}},  # Group 0
        "79": {"M1": {1: {"id": "R7", "length": 1000}}},  # Group 1
    }

    suggestions = ui.get_all_suggestions()

    # Expected sort order:
    # Group 0 (82-97), sorted by width: 82, 85, 95
    # Group 1 (73-79), sorted by width: 75, 78, 79
    # Group 2 (others), sorted by width: 100
    expected_order = ['82', '85', '95', '75', '78', '79', '100']

    assert [s['width'] for s in suggestions] == expected_order
