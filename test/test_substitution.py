from unittest.mock import patch

import polars as pl
import pytest

from cuttingstock.core import OutOfStockError, main_algorithm


# mock the greedy nested function AI!
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
        "width": [12, 12],
        "type": ["A", "B"],
        "demand": [1000, 1000]
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
            # Return the full new spec
            new_spec = e.material_specs.copy()
            new_spec['front'] = 'KA150'
            return new_spec
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
async def test_main_algorithm_multiple_out_of_stock_with_substitution():
    """
    Tests that the main algorithm correctly handles an out-of-stock situation
    by calling the handler and using multiple substitute materials, and that the
    substitutions is remembered for subsequent orders.
    """
    # 1. Mock data
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-1", "ORDER-2"], "order_idx": [0, 1],
        "front": ["KA125", "KA125"],
        "middle": ["KA125", "KA125"],
        "back": ["KA125", "KA125"],
        "width": [12, 12],
        "type": ["A", "B"],
        "demand": [1000, 1000]
    })
    mock_roll_specs = {"85": {"KA150": {1: {"id": "R-KA150-1", "length": 50000}, 2: {"id": "R-KA150-2", "length": 50000}, 3: {"id": "R-KA150-3", "length": 50000}}}}

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
            # Return the full new spec
            new_spec = e.material_specs.copy()
            new_spec['front'] = 'KA150'
            new_spec['middle'] = 'KA150'
            new_spec['back'] = 'KA150'
            return new_spec
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
        assert res1["middle"] == "KA150", "Material should be substituted to KA150"
        assert res1["back"] == "KA150", "Material should be substituted to KA150"
        assert "R-KA150-1" in res1["front_roll_info"]
        assert "R-KA150-2" in res1["middle_roll_info"]
        assert "R-KA150-3" in res1["back_roll_info"]

        res2 = results[1]
        assert res2["order_number"] == "ORDER-2"
        assert res2["front"] == "KA150", "Substitution should be remembered"
        assert res2["middle"] == "KA150", "Substitution should be remembered"
        assert res2["back"] == "KA150", "Substitution should be remembered"
        assert "R-KA150-1" in res2["front_roll_info"]
        assert "R-KA150-2" in res2["middle_roll_info"]
        assert "R-KA150-3" in res2["back_roll_info"]



@pytest.mark.asyncio
async def test_main_algorithm_multiple_out_of_stock_with_some_substitution():
    """
    Tests that the main algorithm correctly handles an out-of-stock situation
    by calling the handler and using some substitute materials, and that the
    substitutions is remembered for subsequent orders.
    """
    # 1. Mock data
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-1", "ORDER-2"], "order_idx": [0, 1],
        "front": ["KA125", "KA125"],
        "middle": ["KA125", "KA125"],
        "back": ["KA125", "KA125"],
        "width": [12, 12],
        "type": ["A", "B"],
        "demand": [1000, 1000]
    })
    mock_roll_specs = {"85": {
        "KA150": {1: {"id": "R-KA150-1", "length": 50000}},
        "KA125": {1: {"id": "R-KA125-1", "length": 50000}, 2: {"id": "R-KA125-2", "length": 50}, 3: {"id": "R-KA125-3", "length": 50000}}
    }}


    # Mock LP solutions that require the out-of-stock material
    mock_lp_solutions = [
        {
            "status": "Optimal",
            "variables": {"roll_w": 85, "demand_per_cut": 1000, "order_idx": 0},
            "material_specs": {"front": "KA125","middle": "KA125", "back": "KA125"},
        },
        {
            "status": "Optimal",
            "variables": {"roll_w": 85, "demand_per_cut": 1000, "order_idx": 1},
            "material_specs": {"front": "KA125","middle": "KA125", "back": "KA125"},
        },
    ]

    handler_calls = []
    def mock_out_of_stock_handler(e: OutOfStockError):
        handler_calls.append(e)
        if e.material == "KA125":
            # Return the full new spec
            new_spec = e.material_specs.copy()
            new_spec['middle'] = 'KA150'
            return new_spec
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
        assert res1["front"] == "KA125", "Material should be the same"
        assert res1["middle"] == "KA150", "Material should be substituted to KA150"
        assert res1["back"] == "KA125", "Material should be the same"
        assert "R-KA125-1" in res1["front_roll_info"]
        assert "R-KA150-1" in res1["middle_roll_info"]
        assert "R-KA125-3" in res1["back_roll_info"]

        res2 = results[1]
        assert res2["order_number"] == "ORDER-2"
        assert res2["front"] == "KA125", "Material should be the same"
        assert res2["middle"] == "KA150", "Material should be substituted to KA150"
        assert res2["back"] == "KA125", "Material should be the same"
        assert "R-KA125-1" in res2["front_roll_info"]
        assert "R-KA150-1" in res2["middle_roll_info"]
        assert "R-KA125-3" in res2["back_roll_info"]



@pytest.mark.asyncio
async def test_main_algorithm_out_of_stock_user_cancel():
    """
    Tests that the main algorithm correctly handles an out-of-stock situation
    when the user cancels the substitution, failing the order.
    """
    # 1. Mock data
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-3"], "order_idx": [0], "front": ["KA125"],
        "width": [12, 12],
        "type": ["A", "B"],
        "demand": [1000, 1000]
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
        return None  # User cancels the substitution by returning None

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


@pytest.mark.asyncio
async def test_get_roll_for_material_atomic_substitution():
    """
    Tests that if a material substitution happens mid-order, the entire order's
    materials are re-calculated from scratch to ensure consistency.
    This validates the core retry logic within get_roll_for_material.
    """
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-MULTI"], "order_idx": [0],
        "front": ["MAT_A"], "back": ["MAT_B"],
        "width": [12, 12],
        "type": ["A", "B"],
        "demand": [1000, 1000]
    })
    # MAT_B will run out of stock, forcing a substitution for the whole spec.
    mock_roll_specs = {
        "80": {
            "MAT_A": {1: {"id": "R-A-1", "length": 5000}},
            "MAT_B": {1: {"id": "R-B-1", "length": 50}}, # Not enough
            "MAT_C": {1: {"id": "R-C-1", "length": 5000}}, # Substitute for A
            "MAT_D": {1: {"id": "R-D-1", "length": 5000}}, # Substitute for B
        }
    }
    mock_lp_solution = {
        "status": "Optimal",
        "variables": {"roll_w": 80, "demand_per_cut": 100, "order_idx": 0},
        "material_specs": {"front": "MAT_A", "back": "MAT_B"},
    }

    handler_calls = []
    def mock_out_of_stock_handler(e: OutOfStockError):
        handler_calls.append(e)
        assert e.material == "MAT_B"
        # The user decides to substitute BOTH materials
        return {"front": "MAT_C", "back": "MAT_D"}

    with patch('cuttingstock.core.load_data'), \
         patch('cuttingstock.core.clean_data', return_value=mock_orders_df), \
         patch('os.path.exists', return_value=False), \
         patch('polars.DataFrame.write_database'), \
         patch('cuttingstock.core.solve_linear_program', return_value=mock_lp_solution):

        results = await main_algorithm(
            roll_width=80,
            roll_length=100000,
            file_path="dummy.csv",
            roll_specs=mock_roll_specs,
            out_of_stock_handler=mock_out_of_stock_handler,
            processed_orders=set(),
            front="MAT_A", back="MAT_B",
        )

        assert len(results) == 1
        assert len(handler_calls) == 1, "Handler should be called for MAT_B"

        res = results[0]
        # Crucially, check that 'front' was re-allocated to MAT_C's roll,
        # even though it was 'found' successfully on the first pass.
        assert res["front"] == "MAT_C"
        assert "R-C-1" in res["front_roll_info"]
        assert res["back"] == "MAT_D"
        assert "R-D-1" in res["back_roll_info"]


@pytest.mark.asyncio
async def test_roll_specs_length_deduction_on_substitution():
    """
    Tests that roll_specs length is correctly deducted from the NEW material
    after a substitution, not the original out-of-stock one.
    """
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-SUB"], "order_idx": [0], "front": ["MAT_A"],
        "width": [12, 12],
        "type": ["A", "B"],
        "demand": [1000, 1000]
    })
    # MAT_A has no stock. MAT_B is the substitute.
    mock_roll_specs = {
        "80": {
            "MAT_A": {},
            "MAT_B": {1: {"id": "R-B-1", "length": 1000}},
        }
    }
    mock_lp_solution = {
        "status": "Optimal",
        "variables": {"roll_w": 80, "demand_per_cut": 300, "order_idx": 0},
        "material_specs": {"front": "MAT_A"},
    }

    def mock_out_of_stock_handler(e: OutOfStockError):
        assert e.material == "MAT_A"
        return {"front": "MAT_B"}

    with patch('cuttingstock.core.load_data'), \
         patch('cuttingstock.core.clean_data', return_value=mock_orders_df), \
         patch('os.path.exists', return_value=False), \
         patch('polars.DataFrame.write_database'), \
         patch('cuttingstock.core.solve_linear_program', return_value=mock_lp_solution):

        await main_algorithm(
            roll_width=80,
            roll_length=100000,
            file_path="dummy.csv",
            roll_specs=mock_roll_specs,
            out_of_stock_handler=mock_out_of_stock_handler,
            processed_orders=set(),
            front="MAT_A",
        )

        # The key assertion: Check that the length was deducted from the substitute material's roll.
        assert mock_roll_specs["80"]["MAT_B"][1]["length"] == 700


@pytest.mark.asyncio
async def test_roll_specs_deduction_on_multi_material_substitution():
    """
    Tests that roll_specs are correctly updated when multiple materials are
    substituted at once in response to a single out-of-stock event.
    """
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-MULTI-SUB"], "order_idx": [0],
        "front": ["MAT_A"], "middle": ["MAT_B"], "back": ["MAT_C"],
        "width": [12, 12],
        "type": ["A", "B"],
        "demand": [1000, 1000]
    })
    # MAT_A is out of stock. User will change all three materials.
    mock_roll_specs = {
        "80": {
            "MAT_A": {}, # Empty
            "MAT_B": {1: {"id": "R-B-1", "length": 1000}},
            "MAT_C": {1: {"id": "R-C-1", "length": 1000}},
            "SUB_A": {1: {"id": "R-SA-1", "length": 1000}}, # Substitute for A
            "SUB_B": {1: {"id": "R-SB-1", "length": 1000}}, # Substitute for B
            "SUB_C": {1: {"id": "R-SC-1", "length": 1000}}, # Substitute for C
        }
    }
    mock_lp_solution = {
        "status": "Optimal",
        "variables": {"roll_w": 80, "demand_per_cut": 200, "order_idx": 0},
        "material_specs": {"front": "MAT_A", "middle": "MAT_B", "back": "MAT_C"},
    }

    def mock_out_of_stock_handler(e: OutOfStockError):
        assert e.material == "MAT_A"
        # User changes all three materials
        return {"front": "SUB_A", "middle": "SUB_B", "back": "SUB_C"}

    with patch('cuttingstock.core.load_data'), \
         patch('cuttingstock.core.clean_data', return_value=mock_orders_df), \
         patch('os.path.exists', return_value=False), \
         patch('polars.DataFrame.write_database'), \
         patch('cuttingstock.core.solve_linear_program', return_value=mock_lp_solution):

        await main_algorithm(
            roll_width=80,
            roll_length=100000,
            file_path="dummy.csv",
            roll_specs=mock_roll_specs,
            out_of_stock_handler=mock_out_of_stock_handler,
            processed_orders=set(),
            front="MAT_A", middle="MAT_B", back="MAT_C",
        )

        # Assert that lengths were deducted from the NEW substitute materials
        assert mock_roll_specs["80"]["SUB_A"][1]["length"] == 800
        assert mock_roll_specs["80"]["SUB_B"][1]["length"] == 800
        assert mock_roll_specs["80"]["SUB_C"][1]["length"] == 800
        # Assert that original materials (that had stock) were untouched
        assert mock_roll_specs["80"]["MAT_B"][1]["length"] == 1000
        assert mock_roll_specs["80"]["MAT_C"][1]["length"] == 1000
