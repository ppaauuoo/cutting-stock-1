import os
import sqlite3
import sys
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import polars as pl
import pytest

from cuttingstock.cleaning import clean_data, load_data
from cuttingstock.core import (
    OutOfStockError,
    generate_suggestions,
)
from cuttingstock.mlmodel import _predict_with_xgboost
from cuttingstock.linear import solve_linear_program

def test_sqlite_basic_operations():
    """
    Tests basic SQLite operations: connection, table creation, insertion, and querying.
    """
    # 1. Connect to an in-memory database
    conn = sqlite3.connect(':memory:')
    cursor = conn.cursor()

    # 2. Create a table
    cursor.execute('''
        CREATE TABLE users (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            email TEXT NOT NULL
        )
    ''')

    # 3. Insert some data
    cursor.execute("INSERT INTO users (name, email) VALUES (?, ?)", ('Alice', 'alice@example.com'))
    cursor.execute("INSERT INTO users (name, email) VALUES (?, ?)", ('Bob', 'bob@example.com'))
    conn.commit()

    # 4. Query the data
    cursor.execute("SELECT name, email FROM users WHERE name = ?", ('Alice',))
    result = cursor.fetchone()

    # 5. Assert the result
    assert result is not None
    assert result[0] == 'Alice'
    assert result[1] == 'alice@example.com'

    # 6. Close the connection
    conn.close()

@pytest.mark.asyncio
async def test_solve_linear_program_simple_case():
    """
    Tests the LP solver with a simple, solvable scenario.
    """
    orders_df = pl.DataFrame({
        "width": [10],
        "length": [100],
        "quantity": [1],
        "type": ["A"],
        "component_type": ["compA"],
    })
    roll_width = 55
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)

    assert result['status'] == 'Optimal'
    assert result['variables']['cuts'] == 5
    assert result['variables']['trim'] == 5

@pytest.mark.asyncio
async def test_solve_linear_program_complex_case():
    """
    Tests the LP solver with a complex, solvable scenario.
    """
    orders_df = pl.DataFrame({
        "order_number": ["ORDER-001"],
        "width": [24.0],
        "length": [100.0],
        "quantity": [10],
        "type": ["A"],
        "component_type": ["None"],
        "demand": [1000.0],
        "front": ["FPAPER"],
        "c": ["CPAPER"],
        "middle": [None],
        "b": [None],
        "back": [None],
        "due_date": ["2025-01-01"],
        "die_cut": [None],
    })
    roll_width = 75
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)

    assert result['status'] == 'Optimal'
    assert result['variables']['cuts'] == 3.0
    assert result['variables']['trim'] == 3.0

@pytest.mark.asyncio
async def test_solve_linear_program_infeasible():
    """
    Tests the LP solver with an infeasible scenario (no orders).
    """
    orders_df = pl.DataFrame()
    roll_width = 55
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)

    assert "Infeasible" in result['status']


def test_generate_suggestions_no_orders():
    """Test that no suggestions are generated when there is no order data."""
    orders_df = pl.DataFrame()
    roll_specs = {
        "80": {"M1": {1: {"id": "R1", "length": 1000}}}
    }
    factory = "รวม"

    suggestions = generate_suggestions(orders_df, roll_specs, factory)
    assert suggestions == []

    # Test with None dataframe
    suggestions = generate_suggestions(None, roll_specs, factory)
    assert suggestions == []


def test_generate_suggestions_no_stock():
    """Test that no suggestions are generated when there is no stock data."""
    orders_df = pl.DataFrame({
        "order_number": ["1"], "front": ["M1"], "c": [None],
        "middle": [None], "b": [None], "back": [None],
    })
    roll_specs = {}
    factory = "รวม"

    suggestions = generate_suggestions(orders_df, roll_specs, factory)
    assert suggestions == []


def test_generate_suggestions_simple_case():
    """Test a basic case with one matching suggestion."""
    orders_df = pl.DataFrame({
        "order_number": ["1"],
        "front": ["M1"], "c": ["C1"], "middle": [None], "b": [None], "back": ["M2"],
    })
    roll_specs = {
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
    factory = "รวม"

    suggestions = generate_suggestions(orders_df, roll_specs, factory)

    assert len(suggestions) == 1
    assert suggestions[0]['width'] == "80"
    assert suggestions[0]['spec'] == {'front': 'M1', 'c': 'C1', 'middle': None, 'b': None, 'back': 'M2'}

def test_generate_suggestions_sorting_default():
    """Test default sorting of suggestions by width (as integer)."""
    orders_df = pl.DataFrame({
        "order_number": ["1"],
        "front": ["M1"], "c": [None], "middle": [None], "b": [None], "back": [None]
    })
    roll_specs = {
        "100": {"M1": {1: {"id": "R1", "length": 1000}}},
        "80": {"M1": {1: {"id": "R2", "length": 1000}}},
        "90": {"M1": {1: {"id": "R3", "length": 1000}}},
    }
    factory = "รวม"

    suggestions = generate_suggestions(orders_df, roll_specs, factory)

    assert len(suggestions) == 3
    assert [s['width'] for s in suggestions] == ['80', '90', '100']


def test_generate_suggestions_sorting_factory_1_2():
    """Test special sorting logic for factory '1&2'."""
    orders_df = pl.DataFrame({
        "order_number": ["1218001"],
        "front": ["M1"], "c": [None], "middle": [None], "b": [None], "back": [None]
    })

    roll_specs = {
        "75": {"M1": {1: {"id": "R1", "length": 1000}}},  # Group 1
        "85": {"M1": {1: {"id": "R2", "length": 1000}}},  # Group 0
        "95": {"M1": {1: {"id": "R3", "length": 900}}},  # Group 0
        "100": {"M1": {1: {"id": "R4", "length": 1000}}}, # Group 2
        "78": {"M1": {1: {"id": "R5", "length": 1000}}},  # Group 1
        "82": {"M1": {1: {"id": "R6", "length": 900}}},  # Group 0
        "79": {"M1": {1: {"id": "R7", "length": 1000}}},  # Group 1
    }

    factory = "1"
    suggestions = generate_suggestions(orders_df, roll_specs, factory)
    expected_order = ['75', '78', '79']

    assert [s['width'] for s in suggestions] == expected_order

    factory = "2"
    suggestions = generate_suggestions(orders_df, roll_specs, factory)
    expected_order = ['85', '82', '95']

    assert [s['width'] for s in suggestions] == expected_order


def test_generate_suggestions_sorting_by_spec():
    """Test that suggestions are sorted by the material spec."""
    orders_df = pl.DataFrame({
        "order_number": ["1", "2", "3", "4","5","6" ],
        "front":        ["M1","M2", "M1", "M1","M1", "M1"],
        "c":            [None,"M2", "M3", None,None, "M2"],
        "middle":       [None,None, "M3", None,None, "M2"],
        "b":            ["M2",None, "M2", "M2","M3", "M2"],
        "back":         ["M3","M4", "M2", "M3","M3", "M2"],
    })
    roll_specs = {
        "80": {
            "M1": {1: {"id": "R1", "length": 1000}},
            "M2": {1: {"id": "R2", "length": 1000}},
            "M3": {1: {"id": "R3", "length": 1000}},
            "M4": {1: {"id": "R4", "length": 1000}},
        },
        "82": {
            "M1": {1: {"id": "R1", "length": 1000}},
            "M2": {1: {"id": "R2", "length": 1000}},
            "M3": {1: {"id": "R3", "length": 1000}},
            "M4": {1: {"id": "R4", "length": 1000}},
        }
    }
    factory = "รวม"

    suggestions = generate_suggestions(orders_df, roll_specs, factory)

    assert len(suggestions) == 10

    # Expected order is based on sorted list of material names
    # ['M1']
    # ['M1', 'M2']
    # ['M1', 'M2', 'M3']
    # ['M2', 'M4']
    specs = [s['spec'] for s in suggestions]
    expected_specs = [
        {'front': 'M1', 'b': 'M2', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M1', 'b': 'M2', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M1', 'b': 'M3', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M1', 'b': 'M3', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M2', 'b': None, 'middle': None, 'c': 'M2', 'back': 'M4'},
        {'front': 'M2', 'b': None, 'middle': None, 'c': 'M2', 'back': 'M4'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M2', 'c': 'M2', 'back': 'M2'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M2', 'c': 'M2', 'back': 'M2'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M3', 'c': 'M3', 'back': 'M2'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M3', 'c': 'M3', 'back': 'M2'},
    ]
    assert specs == expected_specs
    widths = [s['width'] for s in suggestions]
    expected_widths = ["80", "82","80", "82","80", "82","80", "82","80", "82",]
    assert widths == expected_widths

def test_generate_suggestions_sorting_by_len():
    """Test that suggestions are sorted by the material spec."""
    orders_df = pl.DataFrame({
        "order_number": ["1", "2", "3", "4","5","6" ],
        "front":        ["M1","M2", "M1", "M1","M1", "M1"],
        "c":            [None,"M2", "M3", None,None, "M2"],
        "middle":       [None,None, "M3", None,None, "M2"],
        "b":            ["M2",None, "M2", "M2","M3", "M2"],
        "back":         ["M3","M4", "M2", "M3","M3", "M2"],
    })
    roll_specs = {
        "80": {
            "M1": {1: {"id": "R1", "length": 1000}},
            "M2": {1: {"id": "R2", "length": 1000}},
            "M3": {1: {"id": "R3", "length": 1000}},
            "M4": {1: {"id": "R4", "length": 1000}},
        },
        "82": {
            "M1": {1: {"id": "R1", "length": 1000}},
            "M2": {1: {"id": "R2", "length": 1000}},
            "M3": {1: {"id": "R3", "length": 4000}}, # test for priority most in specs stock
            "M4": {1: {"id": "R4", "length": 1000}},
        }
    }
    factory = "รวม"

    suggestions = generate_suggestions(orders_df, roll_specs, factory)

    assert len(suggestions) == 10

    # Expected order is based on sorted list of material names
    # ['M1']
    # ['M1', 'M2']
    # ['M1', 'M2', 'M3']
    # ['M2', 'M4']
    specs = [s['spec'] for s in suggestions]
    expected_specs = [
        {'front': 'M1', 'b': 'M2', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M1', 'b': 'M2', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M1', 'b': 'M3', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M1', 'b': 'M3', 'middle': None, 'c': None, 'back': 'M3'},
        {'front': 'M2', 'b': None, 'middle': None, 'c': 'M2', 'back': 'M4'},
        {'front': 'M2', 'b': None, 'middle': None, 'c': 'M2', 'back': 'M4'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M2', 'c': 'M2', 'back': 'M2'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M2', 'c': 'M2', 'back': 'M2'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M3', 'c': 'M3', 'back': 'M2'},
        {'front': 'M1', 'b': 'M2', 'middle': 'M3', 'c': 'M3', 'back': 'M2'},
    ]
    assert specs == expected_specs
    widths = [s['width'] for s in suggestions]
    expected_widths = ["82", "80", "82", "80", "80", "82", "80", "82","82","80", ]
    assert widths == expected_widths
