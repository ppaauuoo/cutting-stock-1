import pytest
import polars as pl
from unittest.mock import patch

from cuttingstock.core import _find_solution


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

    mock_greedy_results = [
        {"status": "Optimal", "variables": {"order_idx": i}} for i in range(len(orders))
    ]

    # 2. Mock the dependencies
    with patch('cuttingstock.core.greedy_nest') as mock_greedy, \
         patch('cuttingstock.core.format_greedy_results') as mock_format:

        mock_greedy.return_value = (["some_group"], ["some_updated_order"])
        mock_format.return_value = mock_greedy_results

        # 3. Call the function
        results, solution = await _find_solution(
            orders_to_process, roll, None, None, None, original_orders_df
        )

        # 4. Assertions
        mock_greedy.assert_called_once()
        mock_format.assert_called_once()

        assert results == mock_greedy_results
        assert solution["status"] == "GreedyNestingSuccess"
