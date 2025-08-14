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
    _find_and_update_roll,
    solve_linear_program,
)
from cuttingstock.mlmodel import predict_with_xgboost


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
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 500 ม., ใช้หมด) + R2 (ยาว 500 ม., เหลือ 400 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: L1 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, sec_material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R3 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    order_number2 = '2'
    result = _find_and_update_roll(roll_specs, width, material, sec_required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R2 (ยาว 400 ม., ใช้หมด) + R4 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: L1 (ยาว 100 ม., ใช้หมด) + L2 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, sec_material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R3 (ยาว 100 ม., ใช้หมด) + R5 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))

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

def test_find_and_update_roll_multiple_order_width_change():
    """
    Tests the case where width change
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
        },
        '200': {
            'LA125': {
                'L1': {'id': 'L1', 'length': 501},
                'L2': {'id': 'L2', 'length': 502},
                'L3': {'id': 'L3', 'length': 503},
                'L4': {'id': 'L4', 'length': 504},
            },
            'KA125': {
                'R1': {'id': 'R1', 'length': 500},
                'R2': {'id': 'R2', 'length': 500},
                'R3': {'id': 'R3', 'length': 500},
                'R4': {'id': 'R4', 'length': 500},
                'R5': {'id': 'R5', 'length': 500},
            },
        },
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
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 500 ม., ใช้หมด) + R2 (ยาว 500 ม., เหลือ 400 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: L1 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, sec_material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R3 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    order_number2 = '2'
    width = '120'
    with pytest.raises(OutOfStockError) as excinfo:
        _find_and_update_roll(roll_specs, width, material, sec_required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "ไม่มีข้อมูลสต็อก" in str(excinfo.value)
    assert excinfo.value.width == width
    assert excinfo.value.material == material

    order_number3 = '3'
    width = '200'
    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, sec_material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: L4 (ยาว 504 ม., เหลือ 104 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, sec_material))

    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, sec_material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: L3 (ยาว 503 ม., เหลือ 103 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, sec_material))

def test_find_and_update_roll_multiple_order_multiple_different_material_five_roll():
    """
    Tests the case where a five new roll with different material several time.
    """
    roll_specs = {
        '100': {
            'LA125': {
                'L1': {'id': 'L1', 'length': 500},
                'L2': {'id': 'L2', 'length': 500},
                'L3': {'id': 'L3', 'length': 500},
                'L4': {'id': 'L4', 'length': 500},
                'L5': {'id': 'L5', 'length': 500},
                'L6': {'id': 'L6', 'length': 500},
                'L7': {'id': 'L7', 'length': 500},
                'L8': {'id': 'L8', 'length': 500},
                'L9': {'id': 'L9', 'length': 500},
            },
             'CA125': {
                'C1': {'id': 'C1', 'length': 500},
                'C2': {'id': 'C2', 'length': 500},
                'C3': {'id': 'C3', 'length': 500},
                'C4': {'id': 'C4', 'length': 500},
                'C5': {'id': 'C5', 'length': 500},
                'C6': {'id': 'C6', 'length': 500},
                'C7': {'id': 'C7', 'length': 500},
                'C8': {'id': 'C8', 'length': 500},
                'C9': {'id': 'C9', 'length': 500},
            },
            'KA125': {
                'R1': {'id': 'R1', 'length': 500},
                'R2': {'id': 'R2', 'length': 500},
                'R3': {'id': 'R3', 'length': 500},
                'R4': {'id': 'R4', 'length': 500},
                'R5': {'id': 'R5', 'length': 500},
                'R6': {'id': 'R6', 'length': 500},
                'R7': {'id': 'R7', 'length': 500},
                'R8': {'id': 'R8', 'length': 500},
                'R9': {'id': 'R9', 'length': 500},
                'R10': {'id': 'R10', 'length': 500},
                'R11': {'id': 'R11', 'length': 500},
                'R12': {'id': 'R12', 'length': 500},
                'R13': {'id': 'R13', 'length': 500},
                'R14': {'id': 'R14', 'length': 500},
                'R15': {'id': 'R15', 'length': 500},
                'R16': {'id': 'R16', 'length': 500},
                'R17': {'id': 'R17', 'length': 500},
                'R18': {'id': 'R18', 'length': 500},
                'R19': {'id': 'R19', 'length': 500},
                'R20': {'id': 'R20', 'length': 500},
            },
        }
    }
    width = '100'
    material = 'KA125'
    sec_material = 'LA125'
    thr_material = 'CA125'
    required_length = 400
    sec_required_length = 600
    used_roll_ids = set()
    last_used_roll_ids = {}

    order_number1 = '1'
    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, sec_material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: L1 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, sec_material))

    result = _find_and_update_roll(roll_specs, width, material, sec_required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 500 ม., ใช้หมด) + R2 (ยาว 500 ม., เหลือ 400 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R3 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 2 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R4 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, material)
    assert 3 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: R5 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, material))


    result = _find_and_update_roll(roll_specs, width, thr_material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    position_key = ('_position', width, thr_material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> เปิดม้วนใหม่: C1 (ยาว 500 ม., เหลือ 100 ม.)" == result
    assert order_number1 == last_used_roll_ids.get(('_last_order', width, thr_material))


    order_number2 = '2'
    required_length = 300
    sec_required_length = 500
    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, sec_material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: L1 (ยาว 100 ม., ใช้หมด) + L2 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, sec_material))

    result = _find_and_update_roll(roll_specs, width, material, sec_required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R2 (ยาว 400 ม., ใช้หมด) + R6 (ยาว 500 ม., เหลือ 400 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R3 (ยาว 100 ม., ใช้หมด) + R7 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))


    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, material)
    assert 2 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R4 (ยาว 100 ม., ใช้หมด) + R8 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, material)
    assert 3 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R5 (ยาว 100 ม., ใช้หมด) + R9 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, thr_material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    position_key = ('_position', width, thr_material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: C1 (ยาว 100 ม., ใช้หมด) + C2 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number2 == last_used_roll_ids.get(('_last_order', width, thr_material))

    order_number3 = '3'
    required_length = 500
    sec_required_length = 700
    result = _find_and_update_roll(roll_specs, width, sec_material, required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, sec_material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: L2 (ยาว 300 ม., ใช้หมด) + L3 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, sec_material))

    result = _find_and_update_roll(roll_specs, width, material, sec_required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R6 (ยาว 400 ม., ใช้หมด) + R10 (ยาว 500 ม., เหลือ 200 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, material)
    assert 1 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R7 (ยาว 300 ม., ใช้หมด) + R11 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, material)
    assert 2 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R8 (ยาว 300 ม., ใช้หมด) + R12 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, material)
    assert 3 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: R9 (ยาว 300 ม., ใช้หมด) + R13 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, material))

    result = _find_and_update_roll(roll_specs, width, thr_material, required_length, used_roll_ids, last_used_roll_ids, order_number3)
    position_key = ('_position', width, thr_material)
    assert 0 == last_used_roll_ids.get(position_key, 0)
    assert "-> ใช้ม้วนต่อเนื่อง: C2 (ยาว 300 ม., ใช้หมด) + C3 (ยาว 500 ม., เหลือ 300 ม.)" == result
    assert order_number3 == last_used_roll_ids.get(('_last_order', width, thr_material))

    assert roll_specs['100']['KA125']['R1']['length'] == 0
    assert roll_specs['100']['KA125']['R2']['length'] == 0
    assert roll_specs['100']['KA125']['R3']['length'] == 0
    assert roll_specs['100']['KA125']['R4']['length'] == 0
    assert roll_specs['100']['KA125']['R5']['length'] == 0
    assert roll_specs['100']['KA125']['R6']['length'] == 0
    assert roll_specs['100']['KA125']['R7']['length'] == 0
    assert roll_specs['100']['KA125']['R8']['length'] == 0
    assert roll_specs['100']['KA125']['R9']['length'] == 0
    assert roll_specs['100']['KA125']['R10']['length'] == 200
    assert roll_specs['100']['KA125']['R11']['length'] == 300
    assert roll_specs['100']['KA125']['R12']['length'] == 300
    assert roll_specs['100']['KA125']['R13']['length'] == 300
    assert roll_specs['100']['LA125']['L1']['length'] == 0
    assert roll_specs['100']['LA125']['L2']['length'] == 0
    assert roll_specs['100']['LA125']['L3']['length'] == 300
    assert roll_specs['100']['CA125']['C1']['length'] == 0
    assert roll_specs['100']['CA125']['C2']['length'] == 0
    assert roll_specs['100']['CA125']['C3']['length'] == 300
    assert 'R1' in used_roll_ids
    assert 'R2' in used_roll_ids
    assert 'R3' in used_roll_ids
    assert 'R4' in used_roll_ids
    assert 'R5' in used_roll_ids
    assert 'R6' in used_roll_ids
    assert 'R7' in used_roll_ids
    assert 'R8' in used_roll_ids
    assert 'R9' in used_roll_ids
    assert 'R10' in used_roll_ids
    assert 'R11' in used_roll_ids
    assert 'R12' in used_roll_ids
    assert 'R13' in used_roll_ids
    assert 'R14' not in used_roll_ids


def test_find_and_update_roll_order_switching_and_positioning():
    """
    Tests the logic of switching between orders and tracking material position within an order.
    - An order using the same material multiple times should use a new roll each time.
    - A new order should continue using a roll from a previous order if possible.
    """
    roll_specs = {
        '100': {
            'MAT_A': {
                'R1': {'id': 'R1', 'length': 1000},
                'R2': {'id': 'R2', 'length': 1000},
                'R3': {'id': 'R3', 'length': 1000},
            }
        }
    }
    width = '100'
    material = 'MAT_A'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}

    # --- Order 1 ---
    # First component of Order 1, Material A
    order_number1 = 'ORDER_1'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: R1 (ยาว 1000 ม., เหลือ 600 ม.)" == result
    assert roll_specs['100']['MAT_A']['R1']['length'] == 600
    position_key = ('_position', width, material)
    assert 0 == last_used_roll_ids.get(position_key, 0) # position starts at 0

    # Second component of Order 1, same Material A -> should use a NEW roll
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number1)
    assert "-> เปิดม้วนใหม่: R2 (ยาว 1000 ม., เหลือ 600 ม.)" == result
    assert roll_specs['100']['MAT_A']['R2']['length'] == 600
    assert 1 == last_used_roll_ids.get(position_key, 0) # position increments

    # --- Order 2 ---
    # First component of Order 2, Material A -> should CONTINUE using last roll (R2)
    order_number2 = 'ORDER_2'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: R1 (ยาว 600 ม., เหลือ 200 ม.)" == result
    assert roll_specs['100']['MAT_A']['R1']['length'] == 200
    assert 0 == last_used_roll_ids.get(position_key, 0) # position resets for new order

    # Second component of Order 2, same Material A -> should use a NEW roll (R3)
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids, order_number2)
    assert "-> ใช้ม้วนต่อเนื่อง: R2 (ยาว 600 ม., เหลือ 200 ม.)" == result
    assert roll_specs['100']['MAT_A']['R2']['length'] == 200
    assert 1 == last_used_roll_ids.get(position_key, 0) # position increments


def test_find_and_update_roll_no_duplicate_roll_id_in_same_order():
    """
    Tests that a single roll ID is not used for different materials within the same order processing cut.
    """
    roll_specs = {
        '100': {
            'MAT_A': {
                'R1_key': {'id': 'ID1', 'length': 500},
                'R2_key': {'id': 'ID2', 'length': 500},
            },
            'MAT_B': {
                'R3_key': {'id': 'ID1', 'length': 500}, # Same ID as R1_key
                'R4_key': {'id': 'ID3', 'length': 500},
            }
        }
    }
    width = '100'
    material_A = 'MAT_A'
    material_B = 'MAT_B'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    order_number = 'ORDER_1'

    # First call for material A. Should use ID1.
    result_A = _find_and_update_roll(roll_specs, width, material_A, required_length, used_roll_ids, last_used_roll_ids, order_number)
    assert "ID1" in result_A
    assert 'ID1' in used_roll_ids
    assert 'ID2' not in used_roll_ids
    assert 'ID3' not in used_roll_ids
    assert roll_specs['100']['MAT_A']['R1_key']['length'] == 100

    # Second call for material B. Should NOT use ID1 again. It should use ID3.
    # Because ID1 is already in used_roll_ids
    result_B = _find_and_update_roll(roll_specs, width, material_B, required_length, used_roll_ids, last_used_roll_ids, order_number)
    assert "ID1" not in result_B
    assert "ID3" in result_B
    assert 'ID1' in used_roll_ids
    assert 'ID3' in used_roll_ids
    assert roll_specs['100']['MAT_B']['R3_key']['length'] == 500 # Unchanged
    assert roll_specs['100']['MAT_B']['R4_key']['length'] == 100


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

    with pytest.raises(OutOfStockError) as excinfo:
        _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids)
    assert "ไม่มีข้อมูลสต็อก" in str(excinfo.value)
    assert excinfo.value.width == width
    assert excinfo.value.material == material

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
