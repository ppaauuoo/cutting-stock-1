from unittest.mock import patch

import polars as pl
import pytest

from cuttingstock.core import OutOfStockError, main_algorithm


@pytest.mark.asyncio
async def test_main_algorithm_out_of_stock_with_substitution():
    """
    Tests that the main algorithm correctly handles an out-of-stock situation
    by calling the handler and using the substitute material, and that the
    substitution is remembered for subsequent orders.
    """
    # 1. Mock data
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-1", "ORDER-2"], "order_idx": [0, 1],
        "front": ["KA125", "KA125"],
    })
    mock_roll_specs = {"85": {"KA150": {1: {"id": "R-KA150-1", "length": 50000}}}}

    # Mock LP solutions that require the out-of-stock material
    mock_lp_solutions = [
        {
            "status": "Optimal",
            "variables": {"roll_w": 85, "demand_per_cut": 1000, "order_idx": 0},
            "material_specs": {"front": "KA125"},
        },
        {
            "status": "Optimal",
            "variables": {"roll_w": 85, "demand_per_cut": 1000, "order_idx": 1},
            "material_specs": {"front": "KA125"},
        },
    ]

    handler_calls = []
    def mock_out_of_stock_handler(e: OutOfStockError):
        handler_calls.append(e)
        if e.material == "KA125":
            return "KA150"  # Substitute with KA150
        return None

    # 2. Patch dependencies to isolate the algorithm
    with patch('cuttingstock.core.load_data'), \
         patch('cuttingstock.core.clean_data', return_value=mock_orders_df), \
         patch('os.path.exists', return_value=False), \
         patch('polars.DataFrame.write_database'), \
         patch('cuttingstock.core.solve_linear_program', side_effect=mock_lp_solutions):

        # 3. Run the algorithm
        results = await main_algorithm(
            roll_width=85,
            roll_length=100000,
            file_path="dummy.csv",
            roll_specs=mock_roll_specs,
            out_of_stock_handler=mock_out_of_stock_handler,
            processed_orders=set(),
            front="KA125",
        )

        # 4. Assertions
        assert len(results) == 2, "Both orders should have been processed"
        assert len(handler_calls) == 1, "Handler should only be called once"
        results.sort(key=lambda x: x["order_number"])

        res1 = results[0]
        assert res1["order_number"] == "ORDER-1"
        assert res1["front"] == "KA150", "Material should be substituted to KA150"
        assert "R-KA150-1" in res1["front_roll_info"]

        res2 = results[1]
        assert res2["order_number"] == "ORDER-2"
        assert res2["front"] == "KA150", "Substitution should be remembered"
        assert "R-KA150-1" in res2["front_roll_info"]


@pytest.mark.asyncio
async def test_main_algorithm_out_of_stock_user_cancel():
    """
    Tests that the main algorithm correctly handles an out-of-stock situation
    when the user cancels the substitution, failing the order.
    """
    # 1. Mock data
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-3"], "order_idx": [0], "front": ["KA125"]
    })
    mock_roll_specs = {"85": {"KA150": {1: {"id": "R-KA150-1", "length": 50000}}}}

    mock_lp_solution = {
        "status": "Optimal",
        "variables": {"roll_w": 85, "demand_per_cut": 1000, "order_idx": 0},
        "material_specs": {"front": "KA125"},
    }

    handler_calls = []
    def mock_out_of_stock_handler_cancel(e: OutOfStockError):
        handler_calls.append(e)
        return None  # User cancels the substitution

    # 2. Patch dependencies
    with patch('cuttingstock.core.load_data'), \
         patch('cuttingstock.core.clean_data', return_value=mock_orders_df), \
         patch('os.path.exists', return_value=False), \
         patch('polars.DataFrame.write_database'), \
         patch('cuttingstock.core.solve_linear_program', return_value=mock_lp_solution):

        # 3. Run the algorithm
        results = await main_algorithm(
            roll_width=85,
            roll_length=100000,
            file_path="dummy.csv",
            roll_specs=mock_roll_specs,
            out_of_stock_handler=mock_out_of_stock_handler_cancel,
            processed_orders=set(),
            front="KA125",
        )

        # 4. Assertions
        assert len(results) == 1, "One unprocessed order should be in the results"
        assert len(handler_calls) == 1, "Handler should be called once"

        res = results[0]
        assert res["order_number"] == "ORDER-3"
        assert res["roll_w"].startswith("Failed"), "Roll processing should be marked as failed"
        assert "ผู้ใช้ยกเลิก" in res["front_roll_info"], "Roll info should indicate user cancellation"
