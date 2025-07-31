import asyncio
from unittest.mock import patch

import polars as pl
import pytest

from core import main_algorithm


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
        "b": [None],
        "back": [None],
        "due_date": ["2025-01-01"],
        "die_cut": [None],
    })

    mock_cleaned_df = mock_orders_df.clone()

    mock_roll_specs = {
        "48": {
            "FPAPER": {"R1": {"id": "R1", "length": 5000}},
            "CPAPER": {"R2": {"id": "R2", "length": 5000}}
        }
    }

    # 2. Patch dependencies
    # We patch clean_data to return our controlled test data.
    # We patch file/db access to prevent side effects.
    with patch('cleaning.load_data'), \
         patch('cleaning.clean_data', return_value=mock_cleaned_df), \
         patch('os.path.exists', return_value=False), \
         patch('polars.DataFrame.write_database'):

        # 3. Run the algorithm
        results = await main_algorithm(
            roll_width=48,
            roll_length=100000,
            file_path="dummy.csv",  # This will be ignored due to mocks
            roll_specs=mock_roll_specs,
            front="FPAPER",  # Filter criteria
            c="CPAPER",     # Filter criteria
            c_type="C",
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
        # Check that new rolls were opened, using the Thai text from core.py
        assert result["front_roll_info"].startswith("-> เปิดม้วนใหม่:")
        assert result["c_roll_info"].startswith("-> เปิดม้วนใหม่:")
