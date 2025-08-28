import pytest
import polars as pl
from unittest.mock import patch, AsyncMock

from cuttingstock.core import _find_solution
from cuttingstock.mlmodel import try_xgboost_solution
from cuttingstock.grouping import greedy_nest


@pytest.mark.asyncio
async def test_find_solution_greedy_nest_w_pulp_success():
    """
    Tests that _find_solution correctly returns results from greedy_nest
    when it finds a complete solution.
    """
    # 1. Setup mock data
    orders = [
        {"order_number": "A", "width": 40, "type": "X", "demand": 10},
        {"order_number": "B", "width": 13.2, "type": "N", "demand": 5},
        {"order_number": "C", "width": 40, "type": "X", "demand": 8},  # Example with same width as A
        {"order_number": "D", "width": 13, "type": "W", "demand": 3},
        {"order_number": "E", "width": 51, "type": "X", "demand": 7},
        {"order_number": "F", "width": 16, "type": "N", "demand": 2},
        {"order_number": "G", "width": 17, "type": "W", "demand": 4},
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
            'group_id': 'B-E',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 5,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 2.5999999999999943,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 2,
                'demand_per_cut': 6.35,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 1,
                'order_l': 100,
                'order_qty': 5,
                'order_w': 13.2,
                'rem_roll_l': 982.22,
                'roll_w': 80,
                'trim': 2.5999999999999943,
                'type': 'N',
            },
        },
        {
            'group_id': 'B-E',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 7,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 2.5999999999999943,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 1,
                'demand_per_cut': 17.78,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 4,
                'order_l': 100,
                'order_qty': 7,
                'order_w': 51.0,
                'rem_roll_l': 982.22,
                'roll_w': 80,
                'trim': 2.5999999999999943,
                'type': 'X',
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
                'demand_per_cut': 1.27,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 3,
                'order_l': 100,
                'order_qty': 3,
                'order_w': 13.0,
                'rem_roll_l': 998.73,
                'roll_w': 80,
                'trim': 2.0,
                'type': 'W',
            },
        }
     ]    # 3. Call the function
    results, solution = await _find_solution(
        orders_to_process, roll, None, None, None, original_orders_df
    )

    # 4. Assertions
    assert results == mock_greedy_results
    assert solution["status"] == "GreedyNestingSuccess"


@pytest.mark.asyncio
async def test_find_solution_greedy_nest_double_success():
    """
    Tests that _find_solution correctly returns results from greedy_nest
    when it finds a complete solution.
    """
    # 1. Setup mock data
    orders = [
        {"order_number": "A", "width": 40, "type": "X", "demand": 10},
        {"order_number": "B", "width": 13, "type": "N", "demand": 5},
        {"order_number": "C", "width": 40, "type": "X", "demand": 8},  # Example with same width as A
        {"order_number": "D", "width": 13, "type": "W", "demand": 3},
        {"order_number": "E", "width": 51, "type": "X", "demand": 7},
        {"order_number": "F", "width": 16, "type": "N", "demand": 2},
        {"order_number": "G", "width": 17, "type": "W", "demand": 4},
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
            'group_id': 'B-C',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 5,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 1,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 3,
                'demand_per_cut': 4.2333,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 1,
                'order_l': 100,
                'order_qty': 5,
                'order_w': 13,
                'rem_roll_l': 979.68,
                'roll_w': 80,
                'trim': 1,
                'type': 'N',
            },
        },
        {
            'group_id': 'B-C',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 8,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 1,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 1,
                'demand_per_cut': 20.32,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 2,
                'order_l': 100,
                'order_qty': 8,
                'order_w': 40,
                'rem_roll_l': 979.68,
                'roll_w': 80,
                'trim': 1,
                'type': 'X',
            },
        },
        {
            'group_id': 'D-E',
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
                'order_idx': 3,
                'order_l': 100,
                'order_qty': 3,
                'order_w': 13,
                'rem_roll_l': 982.22,
                'roll_w': 80,
                'trim': 3,
                'type': 'W',
            },
        },
        {
            'group_id': 'D-E',
            'material_specs': {
                'b_type': None,
                'c_type': None,
                'quantity': 7,
            },
            'message': 'Greedy Nesting solution found.',
            'objective_value': 3,
            'status': 'Optimal',
            'variables': {
                'component_type': '',
                'cuts': 1,
                'demand_per_cut': 17.78,
                'due_date': None,
                'order_dmd': None,
                'order_idx': 4,
                'order_l': 100,
                'order_qty': 7,
                'order_w': 51,
                'rem_roll_l': 982.22,
                'roll_w': 80,
                'trim': 3,
                'type': 'X',
            },
        },
     ]    # 3. Call the function
    results, solution = await _find_solution(
        orders_to_process, roll, None, None, None, original_orders_df
    )

    # 4. Assertions
    assert results == mock_greedy_results
    assert solution["status"] == "GreedyNestingSuccess"
