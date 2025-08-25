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
    original_orders_df = pl.DataFrame({
        "order_number": ["A", "B"],
        "width": [30, 40],
        "length": [100, 200],
        "quantity": [1, 1],
        "type": ["", ""],
        "component_type": ["", ""],
        "original_idx": [0, 1]
    })
    orders_to_process = original_orders_df.clone()
    roll = {'width': 80, 'length': 1000}

    mock_greedy_results = [
        {"status": "Optimal", "variables": {"order_idx": 0}},
        {"status": "Optimal", "variables": {"order_idx": 1}}
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
