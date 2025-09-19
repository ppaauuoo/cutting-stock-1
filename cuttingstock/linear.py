import polars as pl
from typing import Optional
from pulp import (
    PULP_CBC_CMD,
    LpBinary,
    LpInteger,
    LpMinimize,
    LpProblem,
    LpStatus,
    LpVariable,
    lpSum,
    value,
)
from cuttingstock.utils import log_message

# Constants
INCH_TO_M = 25.4 / 1000  # Conversion factor from inches
MIN_TRIM_WASTE = 1
MAX_TRIM_WASTE = 5
MAX_CUTS_ACROSS_WIDTH = 6
MAX_Z_FOR_TYPE_X = 5

# Status Messages
STATUS_OPTIMAL = "Optimal"
STATUS_INFEASIBLE_NO_ORDERS = "Infeasible - No orders left"
STATUS_SOLVER_ERROR = "Solver Error"
STATUS_INFEASIBLE = "Infeasible"
STATUS_FAILED = "Failed"

CORRUGATE_MULTIPLIERS = {
    "C": 1.45,
    "B": 1.35,
    "E": 1.25,
}


async def solve_linear_program(
    roll_width: int,
    roll_length: int,
    orders_df: pl.DataFrame,
    c_type: Optional[str] = None,  # New parameters for corrugate types
    b_type: Optional[str] = None,  # New parameters for corrugate types
) -> dict:
    """
    Solve a simple Linear Programming problem using PuLP for a given roll paper width
    and available orders DataFrame, considering different corrugate types.
    """
    # 1. Create the LP problem
    # The original objective seems to be related to minimizing trim waste
    # Therefore, change to LpMinimize
    prob = LpProblem(f"LP_Roll_{roll_width}x{roll_length}", LpMinimize)

    # 2. Create decision variables
    if orders_df.is_empty():
        return {"status": STATUS_INFEASIBLE_NO_ORDERS, "message": "No available orders to cut from."}

    most_demand_type = None
    if c_type == 'C':
        most_demand_type = 'C'
    elif b_type == 'B':
        most_demand_type = 'B'
    elif 'E' in (c_type, b_type):
        most_demand_type = 'E'

    widths = orders_df['width'].to_list()
    lengths = orders_df['length'].to_list()
    quantities = orders_df['quantity'].to_list()
    types = orders_df['type'].to_list()
    component_types = orders_df['component_type'].to_list()

    # Define existing_cols based on available columns in orders_df
    material_cols = ['front', 'c', 'middle', 'b', 'back']
    existing_cols = [col for col in material_cols if col in orders_df.columns]

    num_orders = len(widths)
    # y[j] = 1 if order j is selected, 0 otherwise
    y = LpVariable.dicts("select_order", range(num_orders), cat=LpBinary)

    # z = number of times the selected order is cut across the width
    non_zero_widths = [w for w in widths if w > 0]
    min_width = min(non_zero_widths) if non_zero_widths else 1
    max_z = int(roll_width / min_width) if min_width > 0 else 1000
    z = LpVariable("num_cuts", 0, max_z, LpInteger)

    # M for Big-M method linearization
    M = max_z

    # z_width[j] = z if order j is selected, otherwise 0
    z_width = LpVariable.dicts("z_width_part", range(num_orders), 0, None)

    # 3. Define constraints
    # Select only one order
    prob += lpSum(y[j] for j in range(num_orders)) == 1, "SelectOneOrder"

    # Linearization constraints for z_width[j] = y[j] * z
    for j in range(num_orders):
        prob += z_width[j] <= z, f"Linearize_Z_1_{j}"
        prob += z_width[j] <= M * y[j], f"Linearize_Z_2_{j}"
        prob += z_width[j] >= z - M * (1 - y[j]), f"Linearize_Z_3_{j}"
        prob += z_width[j] >= 0, f"Linearize_Z_4_{j}"
        prob += z_width[j] <= MAX_CUTS_ACROSS_WIDTH, f"MaxCutsAcrossWidth_{j}" # TODO: This seems to be a hardcoded business rule.

        # If order type is 'X', limit z to 5 cuts
        if 'X' in (types[j], component_types[j]):
            prob += z <= MAX_Z_FOR_TYPE_X + M * (1 - y[j]), f"MaxZ_TypeX_{j}"


    total_cut_width = lpSum(widths[j] * z_width[j] for j in range(num_orders))

    # 4. Define objective function and related constraints
    corr_multiplier = CORRUGATE_MULTIPLIERS.get(most_demand_type, 1.0)

    # Total length of material required for the selected order (using .get with a default value)
    # The sum will effectively pick the one order where y[j]=1
    total_order_len = lpSum(
        (lengths[j] * INCH_TO_M * quantities[j] * corr_multiplier * y[j])
        for j in range(num_orders)
    )

    # Objective: Minimize trim waste
    trim_waste = roll_width - total_cut_width
    prob += trim_waste, "MinimizeTrim"

    # Constraints
    prob += trim_waste >= MIN_TRIM_WASTE, "TrimLowerBound"
    prob += trim_waste <= MAX_TRIM_WASTE, "TrimUpperBound"

    # Remaining length on roll must be at least 100
    # not used now because we are not using roll length in the objective function
    # prob += roll_length * z - total_order_len >= 100, "RemainingLengthLowerBound"

    # 5. Solve the problem
    try:
        prob.solve(PULP_CBC_CMD(msg=False))
    except Exception as e:
        log_message("error", "PuLP solver failed", {"error": str(e)})
        return {"status": STATUS_SOLVER_ERROR, "message": f"Solver failed: {str(e)}"}

    # 6. Retrieve and format results
    return await _format_lp_solution(
        prob, y, z, orders_df, roll_width, roll_length, total_order_len, c_type, b_type
    )

async def _format_lp_solution(
    prob: LpProblem, y: dict, z: LpVariable, orders_df: pl.DataFrame,
    roll_width: int, roll_length: int, total_order_len: LpVariable,
    c_type: Optional[str], b_type: Optional[str]
) -> dict:
    """Formats the results from the solved PuLP problem."""
    status = LpStatus[prob.status]
    obj_val = round(value(prob.objective), 4) if value(prob.objective) is not None else None

    sel_idx = next((j for j, v in y.items() if v.varValue == 1), -1)

    if sel_idx == -1:
        return {"status": status, "message": "No optimal solution found or order selected."}

    z_val = z.varValue or 0
    orders_data = orders_df.to_dicts()
    sel_order = orders_data[sel_idx]

    sel_order_w = sel_order.get('width')

    total_len_val = value(total_order_len) or 0
    demand_per_cut = round(total_len_val / z_val, 4) if z_val > 0 else 0
    rem_roll_len = round(roll_length - demand_per_cut, 4)
    trim = round(roll_width - (sel_order_w * z_val), 4) if sel_order_w else None

    material_keys = ['demand', 'front', 'middle', 'back', 'c', 'b', 'die_cut']
    material_specs = {key: sel_order.get(key) for key in material_keys if sel_order.get(key)}
    material_specs.update({'c_type': c_type, 'b_type': b_type})

    return {
        "status": status,
        "objective_value": obj_val,
        "variables": {
            "roll_w": roll_width,
            "rem_roll_l": rem_roll_len,
            "demand_per_cut": demand_per_cut,
            "order_w": sel_order_w,
            "order_l": sel_order.get('length'),
            "order_qty": sel_order.get('quantity'),
            "order_dmd": sel_order.get('demand'),
            "cuts": z_val,
            "trim": trim,
            "order_idx": sel_order.get('original_idx'),
            "type": sel_order.get('type'),
            "component_type": sel_order.get('component_type'),
            "due_date": sel_order.get('due_date'),
        },
        "material_specs": material_specs,
        "message": "PuLP problem solved successfully."
    }
