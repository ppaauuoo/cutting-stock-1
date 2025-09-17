import pytest
from cuttingstock.material import _find_and_update_roll, OutOfStockError

def test_find_and_update_roll_sufficient_single_roll():

    """
    Tests the case where a single new roll from stock is sufficient.
    """
    roll_specs = {
        '100': {
            'KA125': {
                'R1': {'id': 'R1', 'length': 1000},
                'R2': {'id': 'R2', 'length': 800}
            },
            'LA125': {
                'L1': {'id': 'L1', 'length': 1000},
                'L2': {'id': 'L2', 'length': 800}
            }
        }
    }
    width = '100'
    material = 'KA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    roll_positions = {}
    positions = {}
    order_number = '123'
    spec_key = 'test'

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=None, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['KA125']['R1']['length'] == 600
    assert 'R1' in used_roll_ids
    assert 'R2' not in used_roll_ids
    position_key = (width, material, 0)
    assert 'R1' == last_used_roll_ids.get(position_key)
    assert 0 == positions[order_number]
    roll_key = (order_number, material, spec_key)
    assert 1 == roll_positions[roll_key]

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=None, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    # When same material_key is encountered again, should return 'กลุ่มเดียวกัน'
    assert result == 'กลุ่มเดียวกัน'
    # Roll specs should remain unchanged since no new roll was processed
    assert roll_specs['100']['KA125']['R2']['length'] == 800
    assert 'R2' not in used_roll_ids

    material = 'LA125'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=None, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['LA125']['L1']['length'] == 600
    assert 'L1' in used_roll_ids
    position_key = (width, material, positions[order_number])
    assert 'L1' == last_used_roll_ids.get(position_key)
    roll_key = (order_number, material, spec_key)
    assert 1 == roll_positions[roll_key]


    order_number = '321'
    material = 'KA125'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=None, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    # This should work since it's a new order_number (different material_key)
    assert roll_specs['100']['KA125']['R1']['length'] == 200  # 1000 - 400 - 400 = 200
    assert 'R1' in used_roll_ids
    position_key = (width, material, 0)
    assert 'R1' == last_used_roll_ids.get(position_key)
    assert 0 == positions[order_number]
    roll_key = (order_number, material, spec_key)
    assert 1 == roll_positions[roll_key]

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=None, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    # When same material_key is encountered again, should return 'กลุ่มเดียวกัน'
    assert result == 'กลุ่มเดียวกัน'
    # Roll specs should remain unchanged since no new roll was processed
    assert roll_specs['100']['KA125']['R2']['length'] == 800
    assert 'R2' not in used_roll_ids


    material = 'LA125'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=None, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    # This should work since it's a different material (different material_key)
    assert roll_specs['100']['LA125']['L1']['length'] == 200  # 1000 - 400 - 400 = 200
    assert 'L1' in used_roll_ids
    position_key = (width, material, positions[order_number])
    assert 'L1' == last_used_roll_ids.get(position_key)
    assert 1 == positions[order_number]
    roll_key = (order_number, material, spec_key)
    assert 1 == roll_positions[roll_key]

    # KA125 KA125 LA125
    # | 0 | 1 | 2 |  | 0 | 1 | 2 |
    # | 1 | 2 | 1 |  | 2 | 2 | 1 |


def test_find_and_update_roll_group_id():

    """
    Tests the case where a single new roll from stock is sufficient.
    """
    roll_specs = {
        '100': {
            'KA125': {
                'R1': {'id': 'R1', 'length': 1000},
                'R2': {'id': 'R2', 'length': 800},
                'R3': {'id': 'R3', 'length': 800},
                'R4': {'id': 'R4', 'length': 800},
                'R5': {'id': 'R5', 'length': 800}
            },
            'LA125': {
                'L1': {'id': 'L1', 'length': 1000},
                'L2': {'id': 'L2', 'length': 800},
                'L3': {'id': 'L3', 'length': 800}
            }
        }
    }
    width = '100'
    material = 'KA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    roll_positions = {}
    positions = {}
    order_number = '123'
    group_id = '123-321'
    spec_key = 'front'

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['KA125']['R1']['length'] == 600
    assert 'R1' in used_roll_ids
    assert 'R2' not in used_roll_ids
    position_key = (width, material, 0)
    assert 'R1' == last_used_roll_ids.get(position_key)
    assert 0 == positions[group_id]
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]

    spec_key = 'middle'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['KA125']['R2']['length'] == 400
    assert 'R2' in used_roll_ids
    position_key = (width, material, 1)
    assert 'R2' == last_used_roll_ids.get(position_key)
    assert 1 == positions[group_id]
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]

    material = 'LA125'
    spec_key = 'back'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['LA125']['L1']['length'] == 600
    assert 'L1' in used_roll_ids
    position_key = (width, material, positions[group_id])
    assert 'L1' == last_used_roll_ids.get(position_key)
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]


    order_number = '321'
    material = 'KA125'
    group_id = '123-321'
    spec_key = 'front'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)
    assert result == 'กลุ่มเดียวกัน'

    spec_key = 'middle'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)


    assert result == 'กลุ่มเดียวกัน'


    material = 'LA125'
    spec_key = 'back'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)
    assert result == 'กลุ่มเดียวกัน'

def test_find_and_update_roll_countinuous_group_id():

    """
    Tests the case where a single new roll from stock is sufficient.
    """
    roll_specs = {
        '100': {
            'KA125': {
                'R1': {'id': 'R1', 'length': 1000},
                'R2': {'id': 'R2', 'length': 800},
                'R3': {'id': 'R3', 'length': 800},
                'R4': {'id': 'R4', 'length': 800},
                'R5': {'id': 'R5', 'length': 800}
            },
            'LA125': {
                'L1': {'id': 'L1', 'length': 1000},
                'L2': {'id': 'L2', 'length': 800},
                'L3': {'id': 'L3', 'length': 800}
            }
        }
    }
    width = '100'
    material = 'KA125'
    required_length = 400
    used_roll_ids = set()
    last_used_roll_ids = {}
    roll_positions = {}
    positions = {}
    order_number = '123'
    group_id = '123-321'
    spec_key = 'front'

    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['KA125']['R1']['length'] == 600
    assert 'R1' in used_roll_ids
    assert 'R2' not in used_roll_ids
    position_key = (width, material, 0)
    assert 'R1' == last_used_roll_ids.get(position_key)
    assert 0 == positions[group_id]
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]

    spec_key = 'middle'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['KA125']['R2']['length'] == 400
    assert 'R2' in used_roll_ids
    position_key = (width, material, 1)
    assert 'R2' == last_used_roll_ids.get(position_key)
    assert 1 == positions[group_id]
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]

    material = 'LA125'
    spec_key = 'back'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['LA125']['L1']['length'] == 600
    assert 'L1' in used_roll_ids
    position_key = (width, material, positions[group_id])
    assert 'L1' == last_used_roll_ids.get(position_key)
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]


    order_number = '321'
    material = 'KA125'
    group_id = '123-321'
    spec_key = 'front'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert result == 'กลุ่มเดียวกัน'

    spec_key = 'middle'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert result == 'กลุ่มเดียวกัน'


    material = 'LA125'
    spec_key = 'back'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert result == 'กลุ่มเดียวกัน'

# ===================================


    order_number = '433'
    material = 'KA125'
    group_id = '433-322'
    spec_key = 'front'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['KA125']['R1']['length'] == 200
    assert 'R1' in used_roll_ids
    position_key = (width, material, 0)
    assert 'R1' == last_used_roll_ids.get(position_key)
    assert 0 == positions[group_id]
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]

    spec_key = 'middle'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['KA125']['R2']['length'] == 000
    assert 'R2' in used_roll_ids
    position_key = (width, material, 1)
    assert 'R2' == last_used_roll_ids.get(position_key)
    assert 1 == positions[group_id]
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]

    material = 'LA125'
    spec_key = 'back'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert roll_specs['100']['LA125']['L1']['length'] == 200
    assert 'L1' in used_roll_ids
    position_key = (width, material, positions[group_id])
    assert 'L1' == last_used_roll_ids.get(position_key)
    roll_key = (group_id, material, spec_key)
    assert 1 == roll_positions[roll_key]


    order_number = '322'
    material = 'KA125'
    group_id = '433-322'
    spec_key = 'front'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert result == 'กลุ่มเดียวกัน'

    spec_key = 'middle'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert result == 'กลุ่มเดียวกัน'


    material = 'LA125'
    spec_key = 'back'
    result = _find_and_update_roll(roll_specs, width, material, required_length, used_roll_ids, last_used_roll_ids,
                                material_specs=None, group_id=group_id, positions=positions, roll_positions=roll_positions,
                                spec_key=spec_key, order_number=order_number)

    assert result == 'กลุ่มเดียวกัน'
