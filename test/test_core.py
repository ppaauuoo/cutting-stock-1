import os
import sqlite3
import sys
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import polars as pl
import pytest

from cuttingstock.core import (
    _find_and_update_roll,
    solve_linear_program,
)
from cuttingstock.xgboost import predict_with_xgboost


def test_find_and_update_roll_sufficient_single_roll():
    """
    Tests the case where a single new roll from stock is sufficient.
    """
    roll_specs = {
        '100': {
            'KA125': {
                'R1': {'id': 'R1', 'length': 1000},
                'R2': {'id': 'R2', 'length': 500}
            }
        }
    }
    width = '100'
    material = 'KA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids)
    
    assert "-> เปิดม้วนใหม่: R1 (ยาว 1000 ม., เหลือ 600 ม.)" == result
    assert roll_specs['100']['KA125']['R1']['length'] == 600
    assert 'R1' in used_roll_ids
    assert 'R2' not in used_roll_ids

def test_find_and_update_roll_same_order_same_material_multiple_roll():
    """
    Tests the case where a single new roll from stock is sufficient.
    """
    roll_specs = {
        '100': {
            'KA125': {
                'R1': {'id': 'R1', 'length': 1000},
                'R2': {'id': 'R2', 'length': 500}
            },
        }
    }
    width = '100'
    material = 'KA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 1000 ม., เหลือ 600 ม.)" == result


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids)
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
   
    assert "-> เปิดม้วนใหม่: R2 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert roll_specs['100']['KA125']['R2']['length'] == 100
    assert 'R2' in used_roll_ids
    assert 'R1' in used_roll_ids

def test_find_and_update_roll_multiple_order_same_material_same_roll():
    """
    Tests the case where a single new roll from stock is sufficient.
    """
    roll_specs = {
        '100': {
            'KA125': {
                'R1': {'id': 'R1', 'length': 1000},
                'R2': {'id': 'R2', 'length': 500}
            },
        }
    }
    width = '100'
    material = 'KA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    
    order_number1 = '1'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 1000 ม., เหลือ 600 ม.)" == result

    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)

    order_number2 = '2'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
   
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)


    assert "-> ใช้ม้วนต่อเนื่อง: R1 (ยาว 600 ม., เหลือ 200 ม.)" == result
    assert roll_specs['100']['KA125']['R1']['length'] == 200
    assert 'R2' not in used_roll_ids
    assert 'R1' in used_roll_ids

def test_find_and_update_roll_multiple_order_multiple_different_material_multiple_roll():
    """
    Tests the case where a multiple new roll with different material several time.
    """
    roll_specs = {
        '100': {
            'LA125': {
                'L1': {'id': 'L1', 'length': 500},
                'L2': {'id': 'L2', 'length': 500},
                'L3': {'id': 'L3', 'length': 500},
                'L4': {'id': 'L4', 'length': 500},
            },
            'KA125': {
                'R1': {'id': 'R1', 'length': 500},
                'R2': {'id': 'R2', 'length': 500},
                'R3': {'id': 'R3', 'length': 500},
                'R4': {'id': 'R4', 'length': 500},
                'R5': {'id': 'R5', 'length': 500},
            },
        }
    }
    width = '100'
    material = 'KA125'
    sec_material = 'LA125'
    required_length = 400
    sec_required_length = 600
    used_roll_ids = set()
    last_used_roll_ids = {}
    
    order_number1 = '1'
    result = _find_and_update_roll(roll_specs, width, material, sec_required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 500 ม., ใช้หมด) + R2 (ยาว 500 ม., เหลือ 400 ม.)" == result
 
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: L1 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: R3 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)


    order_number2 = '2'
    result = _find_and_update_roll(roll_specs, width, material, sec_required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: R2 (ยาว 400 ม., ใช้หมด) + R4 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))
   
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: L1 (ยาว 100 ม., ใช้หมด) + L2 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: R3 (ยาว 100 ม., ใช้หมด) + R5 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))
   
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)

    assert roll_specs['100']['KA125']['R1']['length'] == 0
    assert roll_specs['100']['KA125']['R2']['length'] == 0
    assert roll_specs['100']['KA125']['R3']['length'] == 0
    assert roll_specs['100']['KA125']['R4']['length'] == 300
    assert roll_specs['100']['KA125']['R5']['length'] == 200
    assert 'R1' in used_roll_ids
    assert 'R2' in used_roll_ids
    assert 'R3' in used_roll_ids
    assert 'R4' in used_roll_ids
    assert 'R5' in used_roll_ids

def test_find_and_update_roll_no_stock():
    """
    Tests the case where there is no stock for the requested material.
    """
    roll_specs = {
        '100': {
            'KA125': {}
        }
    }
    width = '100'
    material = 'KA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids)
    
    assert "-> (ไม่มีข้อมูลสต็อก)" == result

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


@patch('cuttingstock.xgboost.load_models')
def test_predict_with_xgboost(mock_load_models):
    """
    Tests the xgboost prediction function with mocked models.
    """
    # 1. Setup mock models and mappings
    mock_out_model = MagicMock()
    mock_out_model.predict.return_value = [0, 1]  # Mock predictions (indices)

    mock_roll_width_model = MagicMock()
    mock_roll_width_model.predict.return_value = [0, 1] # Mock predictions (indices)

    mock_load_models.return_value = {
        'out_model': mock_out_model,
        'roll_width_model': mock_roll_width_model,
        'reverse_label_mapping_out': {0: 3, 1: 4},  # Mock mapping from index to value
        'reverse_label_mapping_roll_width': {0: 75, 1: 80} # Mock mapping
    }

    # 2. Create sample input DataFrame
    orders_df = pl.DataFrame({
        "width": [24.0, 25.0],
        "length": [100.0, 110.0],
        "quantity": [10, 20],
        "component_type": ["A", "B"],
        "front": ["FP1", "FP2"],
        "c": ["CP1", "CP2"],
        "middle": [None, None],
        "b": [None, None],
        "back": [None, None],
        "type": ["typeA", "typeB"],
    })

    # 3. Call the function to be tested
    out_preds, roll_preds = predict_with_xgboost(orders_df)

    # 4. Assertions
    assert out_preds == [3, 4]
    assert roll_preds == [75, 80]
    
    # Verify that the model's predict method was called once
    mock_out_model.predict.assert_called_once()
    mock_roll_width_model.predict.assert_called_once()

