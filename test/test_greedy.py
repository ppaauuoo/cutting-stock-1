import pytest
import polars as pl
from unittest.mock import patch, AsyncMock

from cuttingstock.core import _find_solution
from cuttingstock.mlmodel import try_xgboost_solution
from cuttingstock.grouping import greedy_nest, format_greedy_results

@pytest.mark.asyncio
async def test_find_solution_greedy_nest_w_pulp_success():
    """
    Tests that _find_solution correctly returns results from greedy_nest
    when it finds a complete solution.
    """
    # 1. Setup mock data
    orders = [
        {"order_number": "B", "width": 13, "type": "N", "demand": 5}, #can't be second pair -> high demand
        {"order_number": "D", "width": 13, "type": "W", "demand": 3}, #grouping run top-down, so D overwrite B as a first
        {"order_number": "G", "width": 17, "type": "W", "demand": 6},
    ]
    original_orders_df = pl.from_dicts(orders).rename({"demand": "quantity"})
    original_orders_df = original_orders_df.with_columns(
        pl.lit(100).alias("length"),
        pl.lit("").alias("component_type"),
        pl.arange(0, len(original_orders_df)).alias("original_idx")
    )
    orders_to_process = original_orders_df.clone()
    roll = {'width': 80, 'length': 1000}

    mock_greedy_results =  [
        {
            'group_id': 'D-G',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 3,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 3,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 2,
                'demand_per_cut': 3.81,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 1,
                'order_l': 100,
                'order_qty': 3,
                'order_w': 13,
                'rem_roll_l': 996.19,
                'roll_w': 80,
                'trim': 3,
                'type': 'W',
            },
        },
        {
            'group_id': 'D-G',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 4,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 3,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 3,
                'demand_per_cut': 3.3867,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 2,
                'order_l': 100,
                'order_qty': 4,
                'order_w': 17,
                'rem_roll_l': 996.19,
                'roll_w': 80,
                'trim': 3,
                'type': 'W',
            },
        },
        {
            'material_specs': {
                'b_type': None,
                'c_type': None,
            },
            'message': 'PuLP problem solved successfully.',
            'objective_value': 2.0,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 6.0,
                'demand_per_cut': 2.1167,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 0,
                'order_l': 100,
                'order_qty': 5,
                'order_w': 13,
                'rem_roll_l': 997.8833,
                'roll_w': 80,
                'trim': 2.0,
                'type': 'N',
            },
        },
     ]    # 3. Call the function
    results, solution = await _find_solution(
        orders_to_process, roll, None, None, None, original_orders_df
    )

    # 4. Assertions
    assert results == mock_greedy_results
    assert solution["status"] == "Optimal"


@pytest.mark.asyncio
async def test_find_solution_greedy_nest_max_out_success():
    """
    Test that X and Y sum out would be 6 and other sum out would be 5
    """
    # 1. Setup mock data
    orders = [
        {"order_number": "A", "width": 17, "type": "D", "demand": 5},
        {"order_number": "B", "width": 9, "type": "D", "demand": 7},
        {"order_number": "C", "width": 17, "type": "X", "demand": 5},
        {"order_number": "D", "width": 9, "type": "Y", "demand": 7},
    ]
    original_orders_df = pl.from_dicts(orders).rename({"demand": "quantity"})
    original_orders_df = original_orders_df.with_columns(
        pl.lit(100).alias("length"),
        pl.lit("").alias("component_type"),
        pl.arange(0, len(original_orders_df)).alias("original_idx")
    )
    orders_to_process = original_orders_df.clone()
    roll = {'width': 80, 'length': 1000}

    mock_greedy_results =  [
        {
            'group_id': 'C-D',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 5,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 2,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 3,
                'demand_per_cut': 4.2333,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 2,
                'order_l': 100,
                'order_qty': 5,
                'order_w': 17,
                'rem_roll_l': 995.7667,
                'roll_w': 80,
                'trim': 2,
                'type': 'X',
            },
        },
        {
            'group_id': 'C-D',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 5,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 2,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 3,
                'demand_per_cut': 4.2333,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 3,
                'order_l': 100,
                'order_qty': 5,
                'order_w': 9,
                'rem_roll_l': 995.7667,
                'roll_w': 80,
                'trim': 2,
                'type': 'Y',
            },
        },
        {
            'group_id': 'A-B',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 5,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 3,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 4,
                'demand_per_cut': 3.175,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 0,
                'order_l': 100,
                'order_qty': 5,
                'order_w': 17,
                'rem_roll_l': 996.825,
                'roll_w': 80,
                'trim': 3,
                'type': 'D',
            },
        },
        {
            'group_id': 'A-B',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 1,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 3,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 1,
                'demand_per_cut': 2.54,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 1,
                'order_l': 100,
                'order_qty': 1,
                'order_w': 9,
                'rem_roll_l': 996.825,
                'roll_w': 80,
                'trim': 3,
                'type': 'D',
            },
        },
     ]    # 3. Call the function
    results, solution = await _find_solution(
        orders_to_process, roll, None, None, None, original_orders_df
    )

    # 4. Assertions
    assert results == mock_greedy_results
    assert solution["status"] == "GreedyNestingSuccess"


def test_format_greedy_results():
    """
    Tests that format_greedy_results correctly transforms grouped data
    into the final result format.
    """
    # 1. Setup mock data
    nested_groups = [
        [
            [{"order_number": "A", "width": 20, "out": 2, "roll": 82, "group_id": "A-B", "quantity": 10}],
            [{"order_number": "B", "width": 30, "out": 1, "roll": 82, "group_id": "A-B", "quantity": 5}]
        ]
    ]
    updated_orders = [
        {"order_number": "A", "width": 20, "quantity": 0, "length": 100, "demand": 1000, "type": "N", "component_type": "sleeve", "due_date": "2025-01-01", "front": "F1", "middle": "M1", "back": "B1", "c": "C1", "b": "B2", "die_cut": "D1"},
        {"order_number": "B", "width": 30, "quantity": 15, "length": 150, "demand": 3000, "type": "W", "component_type": "pad", "due_date": "2025-01-02", "front": "F2"},
    ]
    original_orders_list = [
        {"order_number": "A", "width": 20, "quantity": 10, "length": 100, "demand": 1000, "type": "N", "component_type": "sleeve", "due_date": "2025-01-01", "front": "F1", "middle": "M1", "back": "B1", "c": "C1", "b": "B2", "die_cut": "D1"},
        {"order_number": "B", "width": 30, "quantity": 20, "length": 150, "demand": 3000, "type": "W", "component_type": "pad", "due_date": "2025-01-02", "front": "F2"},
    ]
    original_orders_df = pl.from_dicts(original_orders_list).with_columns(
        pl.arange(0, len(original_orders_list)).alias("original_idx")
    )
    roll_length = 10000
    c_type = "C"
    b_type = "B"

    # 2. Call the function
    results = format_greedy_results(
        nested_groups, updated_orders, original_orders_df, roll_length, c_type, b_type
    )

    # 3. Define expected output
    expected_results = [
        {
            "status": "Optimal",
            "group_id": "A-B",
            "objective_value": 12,
            "variables": {
                "roll_w": 82,
                "rem_roll_l": 9972.3775,
                "demand_per_cut": 18.415,
                "order_w": 20,
                "order_l": 100,
                "order_qty": 10,
                "order_dmd": 1000,
                "cuts": 2,
                "trim": 12,
                "order_idx": 0,
                "type": "N",
                "component_type": "sleeve",
                "due_date": "2025-01-01",
            },
            "material_specs": {
                "quantity": 10,
                "front": "F1",
                "middle": "M1",
                "back": "B1",
                "c": "C1",
                "b": "B2",
                "die_cut": "D1",
                "c_type": "C",
                "b_type": "B",
            },
            "message": "Greedy Nesting solution found.",
        },
        {
            "status": "Optimal",
            "group_id": "A-B",
            "objective_value": 12,
            "variables": {
                "roll_w": 82,
                "rem_roll_l": 9972.3775,
                "demand_per_cut": 27.6225,
                "order_w": 30,
                "order_l": 150,
                "order_qty": 5,
                "order_dmd": 3000,
                "cuts": 1,
                "trim": 12,
                "order_idx": 1,
                "type": "W",
                "component_type": "pad",
                "due_date": "2025-01-02",
            },
            "material_specs": {
                "quantity": 5,
                "front": "F2",
                "c_type": "C",
                "b_type": "B",
            },
            "message": "Greedy Nesting solution found.",
        },
    ]

    # 4. Assertions
    results.sort(key=lambda x: x['variables']['order_idx'])
    assert results == expected_results
