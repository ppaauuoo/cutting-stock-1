import pytest
import polars as pl

from cuttingstock.core import OutOfStockError, main_algorithm


@pytest.mark.asyncio
async def test_main_algorithm_out_of_stock_with_substitution(tmp_path):
    """
    Tests that the main algorithm correctly handles an out-of-stock situation
    by calling the handler and using the substitute material. It also checks
    that the substitution is remembered for subsequent orders.
    """
    order_file = tmp_path / "orders.csv"
    # Both orders require 'KA125', which is out of stock.
    orders_df = pl.DataFrame(
        {
            "order_number": ["ORDER-1", "ORDER-2"],
            "order_idx": [1, 2],
            "width": [20, 20],
            "length": [100, 100],
            "quantity": [10, 10],
            "demand": [1000, 1000],
            "front": ["KA125", "KA125"],
            "due_date": ["2025-01-01", "2025-01-01"],
            "type": ["A", "A"],
            "component_type": ["sheet", "sheet"],
            "die_cut": [None, None],
            "c": [None, None],
            "middle": [None, None],
            "b": [None, None],
            "back": [None, None],
        }
    )
    orders_df.write_csv(order_file)

    # 'KA125' is missing for width 85, but 'KA150' is available.
    roll_specs = {"85": {"KA150": {1: {"id": "R-KA150-1", "length": 50000}}}}

    handler_calls = []

    def mock_out_of_stock_handler(e: OutOfStockError):
        handler_calls.append(e)
        if e.material == "KA125":
            return "KA150"  # Substitute with KA150
        return None

    results = await main_algorithm(
        roll_width=85,
        roll_length=100000,
        file_path=str(order_file),
        roll_specs=roll_specs,
        out_of_stock_handler=mock_out_of_stock_handler,
        processed_orders=set(),
        front="KA125",
    )

    assert len(results) == 2, "Both orders should have been processed"
    assert len(handler_calls) == 1, "Handler should only be called once"

    # The results are not guaranteed to be in order, so we sort them.
    results.sort(key=lambda x: x["order_number"])

    # Check first result
    res1 = results[0]
    assert res1["order_number"] == "ORDER-1"
    assert res1["front"] == "KA150", "Material should be substituted to KA150"
    assert "R-KA150-1" in res1["front_roll_info"], "Should use the substitute roll"
    assert "ผู้ใช้ยกเลิก" not in res1["front_roll_info"]

    # Check second result
    res2 = results[1]
    assert res2["order_number"] == "ORDER-2"
    assert res2["front"] == "KA150", "Substitution should be remembered"
    assert "R-KA150-1" in res2["front_roll_info"], "Should use the same substitute roll"


@pytest.mark.asyncio
async def test_main_algorithm_out_of_stock_user_cancel(tmp_path):
    """
    Tests that the main algorithm correctly handles an out-of-stock situation
    when the user cancels the substitution, failing the order.
    """
    order_file = tmp_path / "orders.csv"
    orders_df = pl.DataFrame(
        {
            "order_number": ["ORDER-3"],
            "order_idx": ["3"],
            "width": ["20"],
            "length": ["100"],
            "quantity": ["10"],
            "demand": ["1000"],
            "front": ["KA125"],
            "due_date": ["2025-01-01"],
            "type": ["A"],
            "component_type": ["sheet"],
            "die_cut": ["1"],
            "c": [""],
            "middle": [""],
            "b": [""],
            "back": [""],
        }
    )
    orders_df.write_csv(order_file)

    # Stock does not contain the required 'KA125' material.
    roll_specs = {"85": {"KA150": {1: {"id": "R-KA150-1", "length": 50000}}}}

    handler_calls = []

    def mock_out_of_stock_handler_cancel(e: OutOfStockError):
        handler_calls.append(e)
        return None  # User cancels the substitution

    results = await main_algorithm(
        roll_width=85,
        roll_length=100000,
        file_path=str(order_file),
        roll_specs=roll_specs,
        out_of_stock_handler=mock_out_of_stock_handler_cancel,
        processed_orders=set(),
        front="KA125",
    )

    assert len(results) == 1, "One unprocessed order should be in the results"
    assert len(handler_calls) == 1, "Handler should be called once"

    res = results[0]
    assert res["order_number"] == "ORDER-3"
    assert res["roll_w"].startswith("Failed"), (
        "Roll processing should be marked as failed"
    )
    assert "ผู้ใช้ยกเลิก" in res["front_roll_info"], (
        "Roll info should indicate user cancellation"
    )
