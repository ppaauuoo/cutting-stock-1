import pytest
import polars as pl
from unittest.mock import patch, AsyncMock

from cuttingstock.core import _find_solution, _try_xgboost_solution
from cuttingstock.grouping import greedy_nest


@pytest.mark.asyncio
async def test_find_solution_greedy_nest_success():
    """
    Tests that _find_solution correctly returns results from greedy_nest
    when it finds a complete solution.
    """
    # 1. Setup mock data
    orders = [
        {"order_number": "A", "width": 40, "type": "X", "demand": 10},
        {"order_number": "B", "width": 13, "type": "N", "demand": 5},
        {"order_number": "C", "width": 40, "type": "X", "demand": 8},  # Example with same width as A
        {"order_number": "D", "width": 4, "type": "W", "demand": 3},
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
     ]
    # 3. Call the function
    results, solution = await _find_solution(
        orders_to_process, roll, None, None, None, original_orders_df
    )

    # 4. Assertions
    assert results == mock_greedy_results
    assert solution["status"] == "GreedyNestingSuccess"



@pytest.mark.asyncio
async def test_find_solution_greedy_nest_xgboost_success():
    """
    Tests that _find_solution correctly attempts greedy_nest,
    and if it doesn't find a complete solution, falls back to _try_xgboost_solution
    and returns its results.
    """
    # 1. Setup mock data
    orders = [
        {"order_number": "A", "width": 40, "type": "X", "quantity": 10},
        {"order_number": "B", "width": 13, "type": "N", "quantity": 5},
        {"order_number": "C", "width": 40, "type": "X", "quantity": 8},
        {"order_number": "D", "width": 16, "type": "W", "quantity": 3},
        {"order_number": "E", "width": 16, "type": "X", "quantity": 7},
        {"order_number": "F", "width": 26, "type": "N", "quantity": 2},
        {"order_number": "G", "width": 15, "type": "W", "quantity": 4},
    ]
    original_orders_df = pl.from_dicts(orders)
    original_orders_df = original_orders_df.with_columns(
        pl.lit(100).alias("length"),
        pl.lit("").alias("component_type"),
        pl.arange(0, len(original_orders_df)).alias("original_idx")
    )
    orders_to_process = original_orders_df.clone()
    roll = {'width': 80, 'length': 1000}

    # Define the expected results from _try_xgboost_solution
    mock_xgboost_results =  [
         {
             'group_id': 'XGBoost-1',
             'material_specs': {
                 'b_type': None,
                 'c_type': None,
                 'quantity': 5,
             },
             'message': 'XGBoost solution found.',
             'objective_value': 0.95, # Changed to reflect XGBoost
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
             'group_id': 'XGBoost-2',
             'material_specs': {
                 'b_type': None,
                 'c_type': None,
                 'quantity': 8,
             },
             'message': 'XGBoost solution found.',
             'objective_value': 0.95, # Changed to reflect XGBoost
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
     ]
    mock_xgboost_solution_status = {"status": "XGBoostSuccess"}

    # 2. Patch the functions
    with patch('cuttingstock.grouping.greedy_nest') as mock_greedy_nest, \
         patch('cuttingstock.core._try_xgboost_solution', new_callable=AsyncMock) as mock_xgboost_solution:

        # Configure greedy_nest to return a non-success status to trigger XGBoost
        mock_greedy_nest.return_value = ([], {"status": "NoGreedySolution"})

        # Configure _try_xgboost_solution to return the expected results
        mock_xgboost_solution.return_value = (mock_xgboost_results, mock_xgboost_solution_status)

        # 3. Call the function
        results, solution = await _find_solution(
            orders_to_process, roll, None, None, None, original_orders_df
        )

        # 4. Assertions
        mock_greedy_nest.assert_called_once()
        mock_xgboost_solution.assert_called_once()

        assert results == mock_xgboost_results
        assert solution == mock_xgboost_solution_status
