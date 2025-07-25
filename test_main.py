import polars as pl
import pytest

from main import _find_and_update_roll, solve_linear_program


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
