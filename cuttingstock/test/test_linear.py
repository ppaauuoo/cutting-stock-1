import pytest
import polars as pl
from cuttingstock.linear import solve_linear_program, STATUS_OPTIMAL, STATUS_INFEASIBLE_NO_ORDERS, INCH_TO_M, CORRUGATE_MULTIPLIERS

# Basic test data
orders_data_optimal = [
    # This one should be selected, width 48 * 2 = 96, trim = 4
    {'width': 48, 'length': 100, 'quantity': 10, 'type': 'A', 'component_type': 'A', 'original_idx': 0, 'demand': 1000, 'front': 'K186', 'due_date': '2025-01-01'},
    # This one doesn't fit well for trim waste
    {'width': 30, 'length': 120, 'quantity': 5, 'type': 'B', 'component_type': 'B', 'original_idx': 1, 'demand': 600, 'front': 'K186', 'due_date': '2025-01-02'},
]

@pytest.mark.asyncio
async def test_solve_linear_program_optimal_solution():
    """ Test that the solver finds an optimal solution. """
    orders_df = pl.DataFrame(orders_data_optimal)
    roll_width = 100
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)

    assert result['status'] == STATUS_OPTIMAL
    assert result['objective_value'] == 4.0
    variables = result['variables']
    assert variables['cuts'] == 2
    assert variables['order_w'] == 48
    assert variables['trim'] == 4.0
    assert variables['order_idx'] == 0

@pytest.mark.asyncio
async def test_solve_linear_program_no_orders():
    """ Test behavior with no orders. """
    orders_df = pl.DataFrame()
    roll_width = 100
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)

    assert result['status'] == STATUS_INFEASIBLE_NO_ORDERS
    assert "No available orders" in result['message']

@pytest.mark.asyncio
async def test_solve_linear_program_infeasible():
    """ Test a scenario where no order can satisfy the trim constraints. """
    orders_data = [
        # 100 - 30*z in [1,5] -> no integer z
        {'width': 30, 'length': 120, 'quantity': 5, 'type': 'B', 'component_type': 'B', 'original_idx': 1},
        # 100 - 47*z in [1,5] -> no integer z
        {'width': 47, 'length': 100, 'quantity': 10, 'type': 'A', 'component_type': 'A', 'original_idx': 0},
    ]
    orders_df = pl.DataFrame(orders_data)
    roll_width = 100
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)
    assert result['status'] == 'Infeasible'

@pytest.mark.asyncio
async def test_solve_linear_program_with_type_x_order():
    """ Test that 'X' type orders are limited to MAX_Z_FOR_TYPE_X cuts. """
    orders_data = [
        # type 'X' order, optimal z would be 6 (trim 4), but constrained to 5, making trim 20 (invalid).
        {'width': 16, 'length': 100, 'quantity': 10, 'type': 'X', 'component_type': 'X', 'original_idx': 0},
        # regular order, z=5 gives trim 5. This should be chosen.
        {'width': 19, 'length': 120, 'quantity': 5, 'type': 'A', 'component_type': 'A', 'original_idx': 1},
    ]
    orders_df = pl.DataFrame(orders_data)
    roll_width = 100
    roll_length = 10000

    result = await solve_linear_program(roll_width, roll_length, orders_df)
    
    assert result['status'] == STATUS_OPTIMAL
    variables = result['variables']
    assert variables['order_w'] == 19
    assert variables['cuts'] == 5
    assert variables['trim'] == 5.0
    assert variables['order_idx'] == 1

@pytest.mark.asyncio
async def test_solve_linear_program_with_corrugate_type():
    """ Test that corrugate type correctly affects length calculation. """
    orders_df = pl.DataFrame(orders_data_optimal)
    roll_width = 100
    roll_length = 10000
    c_type = 'C' # Multiplier of 1.45

    result = await solve_linear_program(roll_width, roll_length, orders_df, c_type=c_type)

    assert result['status'] == STATUS_OPTIMAL
    variables = result['variables']
    
    # Check that remaining length is calculated correctly with multiplier
    selected_order = orders_data_optimal[0]
    total_len_val = (
        selected_order['length'] * INCH_TO_M * selected_order['quantity'] * CORRUGATE_MULTIPLIERS[c_type]
    )
    demand_per_cut = total_len_val / variables['cuts']
    expected_rem_len = round(roll_length - demand_per_cut, 4)

    assert variables['rem_roll_l'] == expected_rem_len
