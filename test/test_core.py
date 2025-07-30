import os
import sys
from unittest.mock import patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import polars as pl
import pytest

from core import _find_and_update_roll, main_algorithm, solve_linear_program


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
    Tests the case where a single new roll from stock is sufficient.
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
            },
        }
    }
    width = '100'
    material = 'KA125'
    sec_material = 'LA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    
    order_number1 = '1'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: L1 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: R2 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)


    order_number2 = '2'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: R1 (ยาว 100 ม., ใช้หมด) + R3 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))
   
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: L1 (ยาว 100 ม., ใช้หมด) + L2 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))

    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: R2 (ยาว 100 ม., ใช้หมด) + R4 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))
   
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)

    assert roll_specs['100']['KA125']['R1']['length'] == 0
    assert roll_specs['100']['KA125']['R2']['length'] == 0
    assert roll_specs['100']['KA125']['R3']['length'] == 200
    assert roll_specs['100']['KA125']['R4']['length'] == 200
    assert 'R1' in used_roll_ids
    assert 'R2' in used_roll_ids
    assert 'R3' in used_roll_ids
    assert 'R4' in used_roll_ids

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
async def test_solve_linear_program_infeasible():
    """
    Tests the LP solver with an infeasible scenario (no orders).
    """
    orders_df = pl.DataFrame()
    roll_width = 55
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)
    
    assert "Infeasible" in result['status']


@pytest.mark.asyncio
async def test_main_algorithm_simple_run():
    """
    Tests main_algorithm with a simple, successful run.
    """
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORD001"], "width": [10], "length": [100], "quantity": [100],
        "type": ["A"], "component_type": ["compA"], "due_date": ["2025-01-01"],
        "front": ["KA125"], "c": [None], "middle": [None], "b": [None], "back": [None], "die_cut": [None],
    }).with_columns([
        pl.col(c).cast(pl.Utf8) for c in ["c", "middle", "b", "back", "die_cut"]
    ])

    # Mock file and cleaning operations to isolate algorithm logic
    with patch("cleaning.load_data", return_value=mock_orders_df), \
         patch("cleaning.clean_data", return_value=mock_orders_df), \
         patch("os.path.exists", return_value=False), \
         patch("os.makedirs"), \
         patch("polars.DataFrame.write_database"):

        roll_specs = {
            '54': {
                'KA125': {
                    'R1': {'id': 'R1', 'length': 1000},
                    'R2': {'id': 'R2', 'length': 500}
                },
            }
        }

        results = await main_algorithm(
            roll_width=54, roll_length=10000, file_path="dummy.csv", roll_specs=roll_specs, front="KA125"
        )

    assert len(results) == 1
    result = results[0]
    assert result["order_number"] == "ORD001"
    assert "-> (ประมวลผลไม่สำเร็จ" in result["front_roll_info"]
    assert result["front_roll_info"] == "dog"
    assert result["cuts"] == 0
    assert result["front"] == "KA125"
    assert int(result["rem_roll_l"]) == 0


@pytest.mark.asyncio
async def test_main_algorithm_insufficient_stock():
    """
    Tests main_algorithm when stock is insufficient for an order.
    """
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORD002"], "width": [10], "length": [100], "quantity": [1],
        "type": ["A"], "component_type": ["compA"], "due_date": ["2025-01-01"],
        "front": ["KA125"], "c": [None], "middle": [None], "b": [None], "back": [None], "die_cut": [None],
    }).with_columns([
        pl.col(c).cast(pl.Utf8) for c in ["c", "middle", "b", "back", "die_cut"]
    ])

    with patch("cleaning.load_data", return_value=mock_orders_df), \
         patch("cleaning.clean_data", return_value=mock_orders_df), \
         patch("os.path.exists", return_value=False), \
         patch("os.makedirs"), \
         patch("polars.DataFrame.write_database"):

        roll_specs = {'54': {'KA125': {'R1': {'id': 'R1', 'length': 1}}}}  # Not enough length

        results = await main_algorithm(
            roll_width=54, roll_length=10000, file_path="dummy.csv", roll_specs=roll_specs, front="KA125"
        )

    assert len(results) == 1
    result = results[0]
    assert result["order_number"] == "ORD002"
    assert "-> (ประมวลผลไม่สำเร็จ" in result["front_roll_info"]


@pytest.mark.asyncio
async def test_main_algorithm_infeasible_order():
    """
    Tests main_algorithm with an order that is infeasible to process.
    """
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORD003"], "width": [60], "length": [100], "quantity": [1],  # width > roll_width
        "type": ["A"], "component_type": ["compA"], "due_date": ["2025-01-01"],
        "front": ["KA125"], "c": [None], "middle": [None], "b": [None], "back": [None], "die_cut": [None],
    }).with_columns([
        pl.col(c).cast(pl.Utf8) for c in ["c", "middle", "b", "back", "die_cut"]
    ])

    with patch("cleaning.load_data", return_value=mock_orders_df), \
         patch("cleaning.clean_data", return_value=mock_orders_df), \
         patch("os.path.exists", return_value=False), \
         patch("os.makedirs"), \
         patch("polars.DataFrame.write_database"):

        roll_specs = {'55': {'KA125': {'R1': {'id': 'R1', 'length': 10000}}}}

        results = await main_algorithm(
            roll_width=55, roll_length=10000, file_path="dummy.csv", roll_specs=roll_specs, front="KA125"
        )

    assert len(results) == 1
    result = results[0]
    assert result["order_number"] == "ORD003"
    assert result["roll_w"] == "Failed/Infeasible"
    assert "-> (ประมวลผลไม่สำเร็จ" in result["front_roll_info"]
