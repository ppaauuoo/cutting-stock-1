import os
import sys
from unittest.mock import patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import polars as pl
import pytest

from cuttingstock.core import main_algorithm


@pytest.mark.asyncio
async def test_main_algorithm_simple_success_case():
    """
    Tests the main_algorithm with a simple case where an order can be fulfilled.
    This test mocks the data cleaning and file I/O to focus on the algorithm's logic.
    """
    # 1. Mock data for a single order
    mock_orders_df = pl.DataFrame({
        "order_number": ["ORDER-001"],
        "width": [10.0],
        "length": [100.0],
        "quantity": [10],
        "type": ["A"],
        "component_type": ["None"],
        "demand": [1000.0],
        "front": ["FPAPER"],
        "c": ["CPAPER"],
        "middle": [None],
        "b": ["BPAPER"],
        "back": [None],
        "due_date": ["2025-01-01"],
        "die_cut": [None],
    })

    mock_cleaned_df = mock_orders_df.clone()

    mock_roll_specs = {
        "48": {
            "FPAPER": {"R1": {"id": "R1", "length": 5000}},
            "CPAPER": {"R2": {"id": "R2", "length": 5000}},
            "BPAPER": {"R3": {"id": "R3", "length": 5000}},
        }
    }

    # 2. Patch dependencies
    # We patch clean_data to return our controlled test data, file/db access,
    # and mock the linear program solver to return a feasible solution.
    mock_lp_solution = {
        "status": "Optimal",
        "variables": {
            "roll_w": 48,
            "rem_roll_l": 99000.0,
            "demand_per_cut": 1450.0,
            "order_w": 10.0,
            "order_l": 100.0,
            "order_qty": 10,
            "order_dmd": 1000.0,
            "cuts": 4.0,
            "trim": 8.0,
            "order_idx": 0,
            "type": "A",
            "component_type": "None",
            "due_date": "2025-01-01",
        },
        "material_specs": {
            "demand": 1000.0,
            "front": "FPAPER",
            "c": "CPAPER",
            "middle": None,
            "back": None,
            "b": "BPAPER",
            "die_cut": None,
            "c_type": "C",
            "b_type": "B",
        },
    }

    with patch('cuttingstock.core.load_data'), \
         patch('cuttingstock.core.clean_data', return_value=mock_cleaned_df), \
         patch('os.path.exists', return_value=False), \
         patch('polars.DataFrame.write_database'), \
         patch('cuttingstock.core.solve_linear_program', return_value=mock_lp_solution):

        # 3. Run the algorithm
        results = await main_algorithm(
            roll_width=48,
            roll_length=100000,
            file_path="dummy.csv",  # This will be ignored due to mocks
            roll_specs=mock_roll_specs,
            front="FPAPER",  # Filter criteria
            c="CPAPER",     # Filter criteria
            c_type="C",
            b="BPAPER",     # Filter criteria
            b_type="B",
            max_records=1,
        )

        # 4. Assertions
        assert isinstance(results, list)
        assert len(results) == 1
        result = results[0]

        assert result["roll_w"] == 48
        assert result["order_number"] == "ORDER-001"
        assert result["cuts"] > 0
        assert "front_roll_info" in result
        assert "c_roll_info" in result
        assert "b_roll_info" in result

        assert result["front_roll_info"] == "-> เปิดม้วนใหม่: R1 (ยาว 5000 ม., เหลือ 4000 ม.)"
        assert result["c_roll_info"] == "-> เปิดม้วนใหม่: R2 (ยาว 5000 ม., เหลือ 3550 ม.)"
        assert result["b_roll_info"] == "-> เปิดม้วนใหม่: R3 (ยาว 5000 ม., เหลือ 3650 ม.)"
