import copy
import logging
import os
import re
import time
from typing import Callable, Optional

import polars as pl
from fastapi import FastAPI
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

from cuttingstock.cleaning import clean_data, load_data
from cuttingstock.grouping import format_greedy_results, greedy_nest
from cuttingstock.mlmodel import predict_with_xgboost

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/cuttingstock.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def _spec_to_key(spec: dict) -> tuple:
    """Converts a spec dictionary to a hashable tuple key."""
    if not spec:
        return tuple()
    # Filter out None values and keys that are not material names
    valid_keys = {'front', 'c', 'middle', 'b', 'back'}
    return tuple(sorted((k, v) for k, v in spec.items() if k in valid_keys and v))

def verify_stock_availability(width: int, material: str, required_length: float, stock_data: pl.DataFrame) -> bool:
    """
    Verifies if the required length of material is truly available in stock.

    Args:
        width: The width of the roll
        material: The material name
        required_length: The required length
        stock_data: The stock data DataFrame

    Returns:
        bool: True if stock is sufficient, False otherwise
    """
    # Filter stock data for matching width and material
    matching_stock = stock_data.filter(
        (pl.col("width") == width) &
        (pl.col("material") == material)
    )

    # Calculate total available length
    total_available = matching_stock.select(pl.col("length").sum()).item() if not matching_stock.is_empty() else 0

    return total_available >= required_length


class OutOfStockError(Exception):
    """Custom exception for out-of-stock events."""
    def __init__(self, message, width, material, required_length, material_specs=None, known_out_of_stock=None):
        super().__init__(message)
        self.width = width
        self.material = material
        self.required_length = required_length
        self.material_specs = material_specs or {}
        self.known_out_of_stock = known_out_of_stock or []

def log_message(level: str, message: str, details: dict = None):
    """
    Log a message with the specified level.

    Args:
        level (str): The log level ('debug', 'info', 'warning', 'error', 'critical')
        message (str): The main message to log
        details (dict, optional): Additional details to include in the log
    """
    log_func = getattr(logger, level.lower(), logger.info)

    if details:
        detail_str = ", ".join([f"{k}: {v}" for k, v in details.items()])
        log_func(f"{message} | Details: {detail_str}")
    else:
        log_func(message)

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


# Global variable to store stock data
_stock_data = None

def _try_partial_roll_for_new_order(
    material_rolls_dict: dict, required_length: float, width: str, material: str,
    used_roll_ids: set, last_used_roll_ids: dict, order_number: Optional[str], position: int
) -> Optional[str]:
    """For a new order, try to find any partially used roll first to minimize waste."""
    used_ids_for_this_width = {
        v for k, v in last_used_roll_ids.items()
        if isinstance(k, tuple) and len(k) == 3 and k[0] == width
    }

    partial_rolls = sorted(
        [(k, r) for k, r in material_rolls_dict.items() if r.get('id') in used_ids_for_this_width and r.get('id') not in used_roll_ids and r.get('length', 0) > 0],
        key=lambda item: item[0]
    )

    for _roll_key, roll in partial_rolls:
        if roll.get('length', 0) >= required_length:
            roll_id = roll.get('id')
            original_length = roll['length']
            roll['length'] -= required_length
            position_key = ('_position', width, material)
            last_order_key = ('_last_order', width, material)
            last_used_roll_ids[(width, material, position)] = roll_id
            last_used_roll_ids[position_key] = position
            last_used_roll_ids[last_order_key] = order_number
            return f"-> ใช้ม้วนต่อเนื่อง: {roll_id} (ยาว {int(original_length)} ม., เหลือ {int(roll['length'])} ม.)"
    return None

def _try_continuation_roll(
    material_rolls_dict: dict, required_length: float, width: str, material: str,
    used_roll_ids: set, last_used_roll_ids: dict, order_number: Optional[str],
    unused_rolls: list, position: int, last_roll_id: Optional[str]
) -> Optional[str]:
    """Try to continue using the last used roll for this material spec."""
    if not last_roll_id:
        return None

    last_roll_data = next(((k, r) for k, r in material_rolls_dict.items() if r.get('id') == last_roll_id), None)

    if last_roll_data:
        _last_roll_key, last_roll = last_roll_data
        position_key = ('_position', width, material)
        last_order_key = ('_last_order', width, material)

        if last_roll['length'] >= required_length:
            original_length = last_roll['length']
            last_roll['length'] -= required_length
            used_roll_ids.add(last_roll_id)
            last_used_roll_ids[(width, material, position)] = last_roll_id
            last_used_roll_ids[position_key] = position
            last_used_roll_ids[last_order_key] = order_number
            return f"-> ใช้ม้วนต่อเนื่อง: {last_roll_id} (ยาว {int(original_length)} ม., เหลือ {int(last_roll['length'])} ม.)"
        else:
            # Not enough length, try to combine with other rolls
            needed_from_another = required_length - last_roll['length']
            original_len_roll1 = last_roll['length']

            supplement_rolls = sorted(
                [(k, r) for k, r in unused_rolls if r.get('id') != last_roll_id],
                key=lambda item: item[1]['length'], reverse=True
            )

            rolls_for_combination = []
            length_from_supplements = 0
            for supp_key, supp_roll in supplement_rolls:
                rolls_for_combination.append((supp_key, supp_roll))
                length_from_supplements += supp_roll.get('length', 0)
                if length_from_supplements >= needed_from_another:
                    break

            if length_from_supplements >= needed_from_another:
                last_roll['length'] = 0
                used_roll_ids.add(last_roll_id)
                message_parts = [f"-> ใช้ม้วนต่อเนื่อง: {last_roll_id} (ยาว {int(original_len_roll1)} ม., ใช้หมด)"]
                remaining_needed = needed_from_another
                new_last_used_roll_id = None

                for i, (supp_key, supp_roll) in enumerate(rolls_for_combination):
                    supp_id = supp_roll.get('id')
                    original_supp_length = supp_roll['length']
                    used_roll_ids.add(supp_id)
                    if remaining_needed > 0:
                        if original_supp_length >= remaining_needed:
                            supp_roll['length'] -= remaining_needed
                            message_parts.append(f"{supp_id} (ยาว {int(original_supp_length)} ม., เหลือ {int(supp_roll['length'])} ม.)")
                            new_last_used_roll_id = supp_id
                            remaining_needed = 0
                        else:
                            supp_roll['length'] = 0
                            message_parts.append(f"{supp_id} (ยาว {int(original_supp_length)} ม., ใช้หมด)")
                            remaining_needed -= original_supp_length
                            if i == len(rolls_for_combination) - 1:
                                new_last_used_roll_id = supp_id

                if new_last_used_roll_id:
                    last_used_roll_ids[(width, material, position)] = new_last_used_roll_id
                    last_used_roll_ids[position_key] = position
                    last_used_roll_ids[last_order_key] = order_number

                return " + ".join(message_parts)
    return None

def _try_new_rolls(
    unused_rolls: list, required_length: float, width: str, material: str,
    used_roll_ids: set, last_used_roll_ids: dict, order_number: Optional[str], position: int
) -> Optional[str]:
    """Fallback to combining one or more new rolls to meet the required length."""
    sorted_unused_rolls = sorted(unused_rolls, key=lambda item: item[1]['length'], reverse=True)

    rolls_for_combination = []
    combined_length = 0
    for roll_key, roll in sorted_unused_rolls:
        rolls_for_combination.append((roll_key, roll))
        combined_length += roll.get('length', 0)
        if combined_length >= required_length:
            break

    if combined_length >= required_length:
        message_parts = []
        remaining_needed = required_length
        new_last_used_roll_id = None

        for i, (roll_key, roll) in enumerate(rolls_for_combination):
            roll_id = roll.get('id')
            original_length = roll.get('length', 0)
            used_roll_ids.add(roll_id)

            if remaining_needed > 0:
                if original_length >= remaining_needed:
                    roll['length'] -= remaining_needed
                    message_parts.append(f"{roll_id} (ยาว {int(original_length)} ม., เหลือ {int(roll['length'])} ม.)")
                    new_last_used_roll_id = roll_id
                    remaining_needed = 0
                else:
                    roll['length'] = 0
                    message_parts.append(f"{roll_id} (ยาว {int(original_length)} ม., ใช้หมด)")
                    remaining_needed -= original_length
                    if i == len(rolls_for_combination) - 1:
                        new_last_used_roll_id = roll_id

        if new_last_used_roll_id:
            position_key = ('_position', width, material)
            last_order_key = ('_last_order', width, material)
            last_used_roll_ids[(width, material, position)] = new_last_used_roll_id
            last_used_roll_ids[position_key] = position
            last_used_roll_ids[last_order_key] = order_number

        return f"-> เปิดม้วนใหม่: " + " + ".join(message_parts)
    return None

def _find_and_update_roll(roll_specs: dict, width: str, material: str, required_length: float, used_roll_ids: set, last_used_roll_ids: dict, order_number: Optional[str] = None, material_specs: Optional[dict] = None, material_substitutions: Optional[dict] = None, known_out_of_stock: Optional[list] = None) -> str:
    """
    Finds a suitable roll by trying different strategies in order of priority:
    1. Use a partially used roll for a new order.
    2. Continue using the last used roll for the same material.
    3. Open a new roll (or combination of rolls).
    """
    if not material or not width:
        return ""

    material_rolls_dict = roll_specs.get(str(width), {}).get(material, {})
    if not material_rolls_dict:
        log_message("error", "Out of stock: No stock data available for material.", {"width": width, "material": material, "required_length": required_length, "material_specs": material_specs, "known_out_of_stock": known_out_of_stock})
        raise OutOfStockError("ไม่มีข้อมูลสต็อก", width, material, required_length, material_specs, known_out_of_stock=known_out_of_stock)

    all_available_rolls = sorted(material_rolls_dict.items(), key=lambda item: item[1]['length'])
    unused_rolls = [(k, r) for k, r in all_available_rolls if r.get('id') not in used_roll_ids]

    # --- State management for roll usage ---
    seen_orders = last_used_roll_ids.setdefault('_seen_orders', set())
    position_key = ('_position', width, material)
    last_order_key = ('_last_order', width, material)
    last_order_number = last_used_roll_ids.get(last_order_key)

    if order_number and order_number != last_order_number and not (order_number and (order_number, material) in seen_orders):
        position = 0
    else:
        position = last_used_roll_ids.get(position_key, 0)

    last_roll_id = last_used_roll_ids.get((width, material, position))
    if order_number and (order_number, material) in seen_orders:
        position = last_used_roll_ids.get(position_key, 0) + 1
        last_roll_id = last_used_roll_ids.get((width, material, position))

    if order_number:
        seen_orders.add((order_number, material))
    is_new_order = (order_number and order_number != last_order_number)

    # --- Strategy 1: For a new order, try to find any partially used roll first ---
    if is_new_order:
        message = _try_partial_roll_for_new_order(
            material_rolls_dict, required_length, width, material,
            used_roll_ids, last_used_roll_ids, order_number, position
        )
        if message:
            return message

    # --- Strategy 2: Try to continue using the last used roll ---
    if last_roll_id and last_roll_id in used_roll_ids and order_number == last_order_number and not (order_number and (order_number, material) in seen_orders):
        position += 1
        last_roll_id = last_used_roll_ids.get((width, material, position))

    message = _try_continuation_roll(
        material_rolls_dict, required_length, width, material, used_roll_ids,
        last_used_roll_ids, order_number, unused_rolls, position, last_roll_id
    )
    if message:
        return message

    # --- Strategy 3: Fallback to opening a new roll ---
    message = _try_new_rolls(
        unused_rolls, required_length, width, material, used_roll_ids,
        last_used_roll_ids, order_number, position
    )
    if message:
        return message

    log_message("error", "Out of stock: Not enough stock length available for material.", {"width": width, "material": material, "required_length": required_length, "material_specs": material_specs, "known_out_of_stock": known_out_of_stock})
    raise OutOfStockError("ไม่มีสต็อกที่พอ", width, material, required_length, material_specs, known_out_of_stock=known_out_of_stock)


def generate_suggestions(orders_df: pl.DataFrame, roll_specs: dict, selected_factory: str) -> list:
    """
    Generates a list of all possible calculation settings based on order frequency and stock.
    """
    if orders_df is None or orders_df.is_empty():
        return []

    # Filter orders based on factory selection
    if "order_number" in orders_df.columns:
        # Use a more robust numeric check for order number prefixes.
        # Cast to string, strip whitespace, then check the numeric value of the prefix.
        order_num_col = pl.col("order_number").cast(pl.Utf8).str.strip_chars()

        if selected_factory == "1&2":
            orders_df = orders_df.filter(
                order_num_col.str.slice(0, 4).str.to_integer(strict=False) == 1218
            )
        elif selected_factory in ["3", "4", "5"]:
            orders_df = orders_df.filter(
                order_num_col.str.slice(0, 1).str.to_integer(strict=False) == int(selected_factory)
            )

    material_cols = ['front', 'c', 'middle', 'b', 'back']
    existing_cols = [col for col in material_cols if col in orders_df.columns]

    if not existing_cols:
        return []

    spec_df = orders_df.with_columns(
        [pl.col(c).fill_null("").str.strip_chars() for c in existing_cols]
    )

    all_specs_df = spec_df.group_by(existing_cols).len().sort("len", descending=True)

    if all_specs_df.is_empty():
        return []

    suggestions = []
    for spec_row in all_specs_df.iter_rows(named=True):
        spec_materials = {m for k, m in spec_row.items() if k != 'len' and m}

        if not spec_materials:
            continue

        available_widths = []
        if roll_specs:
            for width, materials_in_stock in roll_specs.items():
                if spec_materials.issubset(materials_in_stock.keys()):
                    available_widths.append(width)

        if available_widths:
            # make this work, i want to sort the available widths by their count AI!
            sorted_widths = sorted(available_widths, key=lambda x: count(x))

            if selected_factory == "1&2":
                def sort_key_factory_1_2(width_str):
                    width_int = int(re.sub(r'\D', '', width_str) or 0)
                    if 82 <= width_int <= 97:
                        return (0, width_int)
                    elif 73 <= width_int <= 79:
                        return (1, width_int)
                    else:
                        return (2, width_int)
                sorted_widths = sorted(available_widths, key=sort_key_factory_1_2)
            else:
                sorted_widths = sorted(available_widths, key=lambda x: int(re.sub(r'\D', '', x) or 0))



            for width in sorted_widths:
                full_spec = {k: v for k, v in spec_row.items() if k != 'len'}
                suggestion = {'width': width, 'spec': full_spec}
                suggestions.append(suggestion)

    return suggestions


app = FastAPI()

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

async def _try_xgboost_solution(
    orders_to_process: pl.DataFrame, roll: dict, c_type: Optional[str], b_type: Optional[str],
    progress_callback: Optional[Callable[[str], None]]
) -> Optional[dict]:
    """Tries to find a quick solution using the pre-trained XGBoost model."""
    try:
        if progress_callback:
            progress_callback("    🤖 Trying XGBoost for a quick solution...")

        xgb_cuts_preds, xgb_roll_w_preds = predict_with_xgboost(orders_to_process)
        orders_with_preds = orders_to_process.with_columns(
            pl.Series("xgb_cuts", xgb_cuts_preds, dtype=pl.Int64),
            pl.Series("xgb_roll_w", xgb_roll_w_preds, dtype=pl.Int64),
        )
        candidate_orders = orders_with_preds.filter(pl.col("xgb_roll_w") == roll['width'])

        if not candidate_orders.is_empty():
            if progress_callback:
                progress_callback(f"    Found {len(candidate_orders)} candidates from XGBoost for roll {roll['width']}\".")
            for order in candidate_orders.iter_rows(named=True):
                cuts = order.get('xgb_cuts')
                order_w = order.get('width')
                if not cuts or not order_w:
                    continue
                trim = roll['width'] - (order_w * cuts)
                if MIN_TRIM_WASTE <= trim <= MAX_TRIM_WASTE:
                    if progress_callback:
                        progress_callback(f"    ✅ XGBoost found a valid solution for order_idx {order.get('original_idx')}.")
                    sel_order, z_val = order, cuts
                    corr_multiplier = CORRUGATE_MULTIPLIERS.get(c_type or b_type) or 1.0
                    total_len_val = sel_order.get('length') * INCH_TO_M * sel_order.get('quantity') * corr_multiplier
                    demand_per_cut = round(total_len_val / z_val, 4) if z_val > 0 else 0
                    rem_roll_len = round(roll['length'] - demand_per_cut, 4)
                    material_keys = ['demand', 'front', 'middle', 'back', 'c', 'b', 'die_cut']
                    material_specs = {key: sel_order.get(key) for key in material_keys if sel_order.get(key)}
                    material_specs.update({'c_type': c_type, 'b_type': b_type})
                    return {
                        "status": STATUS_OPTIMAL, "objective_value": trim,
                        "variables": {
                            "roll_w": roll['width'], "rem_roll_l": rem_roll_len, "demand_per_cut": demand_per_cut,
                            "order_w": sel_order.get('width'), "order_l": sel_order.get('length'),
                            "order_qty": sel_order.get('quantity'), "order_dmd": sel_order.get('demand'),
                            "cuts": z_val, "trim": trim, "order_idx": sel_order.get('original_idx'),
                            "type": sel_order.get('type'), "component_type": sel_order.get('component_type'),
                            "due_date": sel_order.get('due_date'),
                        },
                        "material_specs": material_specs, "message": "XGBoost solution found."
                    }
    except Exception as e:
        log_message("error", "XGBoost prediction failed.", {"error": str(e)})
        if progress_callback:
            progress_callback(f"    ⚠️ XGBoost prediction failed: {e}. Falling back to linear solver.")
    return None

async def _find_solution(
    orders_to_process: pl.DataFrame, roll: dict, c_type: Optional[str], b_type: Optional[str],
    progress_callback: Optional[Callable[[str], None]], original_orders_df: pl.DataFrame
) -> tuple[list, dict]:
    """
    Attempts to find a cutting solution using a sequence of methods:
    1. Greedy Nesting as a heuristic for a quick, full-roll solution.
    2. XGBoost for a quick, pattern-based solution.
    3. Linear Programming for an optimal single-cut solution if XGBoost also fails.
    """
    # 1. Attempt with Greedy Nesting first, as it can produce a full plan for the roll.
    if progress_callback:
        progress_callback(f"    Trying greedy nesting for roll {roll['width']}...")

    orders_for_greedy = orders_to_process.to_dicts()
    nested_groups, updated_orders = greedy_nest(orders_for_greedy, materials=[roll['width']])
    greedy_results_to_return = []

    if nested_groups:
        greedy_results = format_greedy_results(
            nested_groups, updated_orders, original_orders_df, roll['length'],
            c_type=c_type, b_type=b_type
        )
        if greedy_results:
            if progress_callback:
                progress_callback(f"    ✅ Greedy nesting found a solution with {len(greedy_results)} cuts.")
            greedy_results_to_return = greedy_results

    orders_for_solvers = orders_to_process
    if greedy_results_to_return:
        processed_indices = {
            res.get("variables", {}).get("order_idx") for res in greedy_results_to_return
        }
        processed_indices.discard(None)  # Remove None if it exists
        if processed_indices:
            orders_for_solvers = orders_to_process.filter(
                ~pl.col("original_idx").is_in(list(processed_indices))
            )

    log_message("info", "Orders after greedy", {'orders': orders_for_solvers})
    if not greedy_results_to_return and progress_callback:
        progress_callback("    Greedy nesting did not find a solution. Falling back to XGBoost.")

    # 2. Attempt to find a solution for remaining orders with XGBoost.
    solution = None
    if not orders_for_solvers.is_empty():
        solution = await _try_xgboost_solution(orders_for_solvers, roll, c_type, b_type, progress_callback)

        # 3. If XGBoost fails, fall back to the linear programming solver.
        if solution is None:
            if progress_callback:
                progress_callback("    XGBoost did not find a solution. Falling back to linear solver.")
            solution = await solve_linear_program(
                roll['width'], roll['length'], orders_for_solvers, c_type=c_type, b_type=b_type
            )
    else:
        # This case handles when greedy nesting processes all orders.
        solution = {"status": "NoOrdersLeft", "message": "No orders left for solvers."}

    # If greedy nesting was successful, prioritize its result as it's a complete plan.
    if greedy_results_to_return:
        return greedy_results_to_return, {"status": "GreedyNestingSuccess", "message": "Greedy nesting found a solution."}

    # If the solution from XGBoost or LP is optimal, we can use it directly.
    if solution.get("status") == STATUS_OPTIMAL:
        return [solution], solution

    # If no method produced a viable solution, return an empty list with the last failed solution.
    return [], solution

async def _process_single_order(
    result: dict, orders_df: pl.DataFrame, order_num_col_idx: int, material_substitutions: dict,
    progress_callback: Optional[Callable[[str], None]], out_of_stock_handler: Optional[Callable],
    roll_specs: dict, used_roll_ids_for_cut: set, last_used_roll_ids: dict, roll_width: int
) -> tuple[Optional[dict], Optional[int], Optional[str]]:
    """
    Processes a single order solution, handling stock checks, material substitutions, and roll allocation.
    Returns the final cut information, the processed order index, and any failure reason.
    """
    variables = result.get("variables", {})
    order_idx = variables.get("order_idx")

    if progress_callback:
        progress_callback(f"    Optimal solution found. Trim: {variables.get('trim', 0):.4f}")
        progress_callback(f"    Selected order width: {variables.get('order_w')} (Index: {order_idx}), Cuts: {variables.get('cuts')}")

    order_number = orders_df.row(int(order_idx))[order_num_col_idx] if order_idx is not None else None

    material_specs_for_order = result.get("material_specs", {}).copy()
    order_processed_successfully = False
    final_roll_info = {}
    calculation_failed_reason = None
    material_specs = {}  # Will be set to the final successful spec
    known_out_of_stock_materials = []

    while not order_processed_successfully:
        spec_key_for_lookup = _spec_to_key(material_specs_for_order)
        if material_substitutions and spec_key_for_lookup in material_substitutions:
            substituted_spec = material_substitutions[spec_key_for_lookup]
            if substituted_spec is None:
                if progress_callback: progress_callback("    ❌ User previously cancelled substitution for this spec. Failing order.")
                calculation_failed_reason = "ผู้ใช้ยกเลิกสำหรับสเปคนี้"
                break
            if progress_callback:
                changes_str = ", ".join([f"{k.title()}: {v}" for k, v in substituted_spec.items() if material_specs_for_order.get(k) != v])
                progress_callback(f"    🔄 Applying stored substitution for spec: {changes_str}")
            material_specs_for_order = substituted_spec.copy()

        current_attempt_specs = material_specs_for_order.copy()
        variables = result.get("variables", {})
        roll_info_this_attempt = {}
        spec_changed_this_attempt = False
        calculation_failed_reason = None
        roll_specs_backup = copy.deepcopy(roll_specs)
        last_used_roll_ids_backup = copy.deepcopy(last_used_roll_ids)

        def get_roll_for_material(spec_key: str, value_calculator: Callable[[], float]):
            nonlocal calculation_failed_reason, spec_changed_this_attempt, material_specs_for_order
            if calculation_failed_reason or not current_attempt_specs.get(spec_key): return
            material = str(current_attempt_specs.get(spec_key)).strip()
            try:
                value = value_calculator()
                roll_w_str = str(variables.get("roll_w", "")).strip()
                info = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids, order_number, current_attempt_specs, material_substitutions=material_substitutions, known_out_of_stock=known_out_of_stock_materials)
                roll_info_this_attempt[f'{spec_key}_roll_info'] = info
            except OutOfStockError as e:
                known_out_of_stock_materials.append((e.width, e.material))
                if out_of_stock_handler:
                    if progress_callback: progress_callback(f"    ⚠️ สต็อกสำหรับ '{e.material}' (หน้ากว้าง {e.width}) ไม่พอ, รอการตัดสินใจจากผู้ใช้...")
                    log_message("info", "Out of stock, awaiting user interaction.", {"width": e.width, "material": e.material, "required_length": e.required_length, "material_specs": e.material_specs})
                    new_material_specs = out_of_stock_handler(e)
                    if new_material_specs:
                        changes = {k: v for k, v in new_material_specs.items() if current_attempt_specs.get(k) != v}
                        log_message("info", "User provided material substitution.", {"original_specs": current_attempt_specs, "new_specs": new_material_specs, "changes": changes})
                        if progress_callback:
                            changes_str = ", ".join([f"{k.title()}: {v}" for k, v in changes.items()])
                            progress_callback(f"    ✅ User chose: {changes_str}. Will retry order.")
                        original_spec_key = _spec_to_key(current_attempt_specs)
                        material_substitutions[original_spec_key] = new_material_specs
                        for key, value in list(material_substitutions.items()):
                            if value is not None and _spec_to_key(value) == original_spec_key: material_substitutions[key] = new_material_specs
                        material_specs_for_order = new_material_specs
                        spec_changed_this_attempt = True
                        calculation_failed_reason = "SPEC_CHANGED"
                    else:
                        log_message("warning", "User cancelled material substitution.", {"original_specs": current_attempt_specs, "out_of_stock_material": e.material})
                        if progress_callback: progress_callback(f"    ❌ ผู้ใช้ยกเลิก, ไม่สามารถหาวัสดุสำหรับ '{e.material}' ได้")
                        original_spec_key = _spec_to_key(current_attempt_specs)
                        material_substitutions[original_spec_key] = None
                        for key, value in list(material_substitutions.items()):
                            if value is not None and _spec_to_key(value) == original_spec_key: material_substitutions[key] = None
                        roll_info_this_attempt[f'{spec_key}_roll_info'] = "-> (ผู้ใช้ยกเลิก)"
                        calculation_failed_reason = "ผู้ใช้ยกเลิก"
                else:
                    fail_reason_msg = e.args[0]
                    roll_info_this_attempt[f'{spec_key}_roll_info'] = f"-> ({fail_reason_msg})"
                    calculation_failed_reason = fail_reason_msg

        if roll_specs:
            demand_per_cut = variables.get("demand_per_cut", 0)
            c_type_spec = current_attempt_specs.get('c_type'); b_type_spec = current_attempt_specs.get('b_type')
            type_demand_divisor = CORRUGATE_MULTIPLIERS.get(c_type_spec or b_type_spec) or 1.0
            get_roll_for_material('front', lambda: demand_per_cut / type_demand_divisor)
            if c_type_spec == 'C': get_roll_for_material('c', lambda: demand_per_cut)
            elif c_type_spec == 'E': get_roll_for_material('c', (lambda: demand_per_cut / CORRUGATE_MULTIPLIERS['B'] * CORRUGATE_MULTIPLIERS['E']) if b_type_spec == 'B' else (lambda: demand_per_cut))
            get_roll_for_material('middle', lambda: demand_per_cut / type_demand_divisor)
            if b_type_spec == 'B': get_roll_for_material('b', (lambda: (demand_per_cut / CORRUGATE_MULTIPLIERS['C']) * CORRUGATE_MULTIPLIERS['B']) if c_type_spec == 'C' else (lambda: demand_per_cut))
            elif b_type_spec == 'E': get_roll_for_material('b', (lambda: (demand_per_cut / CORRUGATE_MULTIPLIERS['C']) * CORRUGATE_MULTIPLIERS['E']) if c_type_spec == 'C' else (lambda: demand_per_cut))
            get_roll_for_material('back', lambda: demand_per_cut / type_demand_divisor)

        if spec_changed_this_attempt:
            roll_specs.clear(); roll_specs.update(roll_specs_backup)
            last_used_roll_ids.clear(); last_used_roll_ids.update(last_used_roll_ids_backup)
            used_roll_ids_for_cut.clear()
            if calculation_failed_reason == "SPEC_CHANGED": calculation_failed_reason = None
            if progress_callback: progress_callback("    🔄 Spec changed, restarting roll allocation for this order...")
            continue
        if calculation_failed_reason: break
        if _stock_data is not None:
            insufficient_materials = [
                m for spec_key in ['front', 'c', 'middle', 'b', 'back']
                if (m := current_attempt_specs.get(spec_key)) and not verify_stock_availability(roll_width, m, variables.get("demand_per_cut", 0), _stock_data)
            ]
            if insufficient_materials:
                if progress_callback: progress_callback(f"    ❌ ตรวจพบว่าสต็อกสำหรับ {', '.join(insufficient_materials)} ไม่พอจริงๆ หลังตรวจสอบระบบสต็อก")
                log_message("error", "Confirmed out of stock", {"materials": insufficient_materials})
                calculation_failed_reason = "Confirmed out of stock"
                break
        if calculation_failed_reason: break
        final_roll_info = roll_info_this_attempt
        material_specs = current_attempt_specs
        order_processed_successfully = True

    if not order_processed_successfully:
        if progress_callback: progress_callback(f"    ❌ การคำนวณสำหรับ {order_number} ล้มเหลวเนื่องจาก: {calculation_failed_reason}")
        return None, order_idx, calculation_failed_reason

    cut_info = {
        "roll_w": variables.get("roll_w"), "rem_roll_l": variables.get("rem_roll_l"),
        "demand_per_cut": variables.get("demand_per_cut"), "order_number": order_number,
        "order_w": variables.get("order_w"), "order_l": variables.get("order_l"),
        "order_qty": variables.get("order_qty"), "order_dmd": variables.get("order_dmd"),
        "cuts": variables.get("cuts"), "trim": variables.get("trim"),
        "type": variables.get("type"), "component_type": variables.get("component_type"),
        "due_date": variables.get("due_date"),
    }
    if "group_id" in result: cut_info["group_id"] = result.get("group_id")
    cut_info.update(material_specs)
    cut_info.update(final_roll_info)
    return cut_info, order_idx, None

def _create_unprocessed_result(order: dict, status: str, reason: str) -> dict:
    """Creates a dictionary representing an unprocessed order for the results table."""
    fail_msg = f"-> (ประมวลผลไม่สำเร็จ: {reason})"
    status_with_reason = f"{status} (Reason: {reason})"
    return {
        "roll_w": status_with_reason,
        "rem_roll_l": 0,
        "demand_per_cut": 0,
        "order_number": order.get("order_number"),
        "order_w": order.get("width"),
        "order_l": order.get("length"),
        "order_qty": order.get("quantity"),
        "order_dmd": order.get("demand"),
        "due_date": order.get("due_date"),
        "cuts": 0,
        "trim": 0,
        "type": order.get("type"),
        "component_type": order.get("component_type"),
        "die_cut": order.get("die_cut"),
        "front": order.get("front"),
        "c": order.get("c"),
        "middle": order.get("middle"),
        "b": order.get("b"),
        "back": order.get("back"),
        "front_roll_info": fail_msg,
        "c_roll_info": fail_msg,
        "middle_roll_info": fail_msg,
        "b_roll_info": fail_msg,
        "back_roll_info": fail_msg,
    }

async def main_algorithm(
    roll_width: int,
    roll_length: int,
    file_path: str = "order2024.csv",
    max_records: Optional[int] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    out_of_stock_handler: Optional[Callable[[OutOfStockError], Optional[dict]]] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    front: Optional[str] = None,
    c_type: Optional[str] = None,
    c: Optional[str] = None,
    middle: Optional[str] = None,
    b_type: Optional[str] = None,
    b: Optional[str] = None,
    back: Optional[str] = None,
    roll_specs: Optional[dict] = None,
    processed_orders: Optional[set] = None,
    material_substitutions: Optional[dict] = None,
    chunk_size: Optional[int] = 100,
):
    output_dir = "cache"
    os.makedirs(output_dir, exist_ok=True)

    log_message("info", "Starting calculation process")

    base_filename = os.path.splitext(os.path.basename(file_path))[0]
    cache_db_path = os.path.join(output_dir, f"{base_filename}.db")
    table_name = base_filename
    if os.path.exists(cache_db_path):
        conn_str = f"sqlite:///{os.path.abspath(cache_db_path)}"
        query = f'SELECT * FROM "{table_name}"'
        raw_orders_df = pl.read_database_uri(query, conn_str)
        log_message("info", "Loaded order data from cache", {"cache_path": cache_db_path})
    else:
        raw_orders_df = load_data(file_path)
        log_message("info", "No cache found, loading order data from CSV file", {"file_path": file_path})

    orders_df = clean_data(
        raw_orders_df,
        start_date,
        end_date,
        front=front,
        c=c if c_type in ["C", "E"] else None,
        middle=middle,
        b=b if b_type in ["B", "E"] else None,
        back=back,
    )

    if processed_orders:
        orders_df = orders_df.filter(
            ~pl.col("order_number").is_in(list(processed_orders))
        )

    log_message("info", "Loaded and sorted data successfully")

    if max_records:
        orders_df = orders_df.head(max_records)
    orders_df = orders_df.with_row_index("original_idx")

    # rolls per specs (set of materials)
    rolls = [{"width": roll_width, "length": roll_length}]
    all_results = []

    order_num_col_idx = orders_df.columns.index("order_number")

    if material_substitutions is None:
        material_substitutions = {}

    for roll in rolls:
        last_used_roll_ids = {}
        used_roll_ids_for_cut = set()
        if progress_callback:
            progress_callback(f"🔧 กำลังประมวลผลม้วน {roll['width']} นิ้ว")

        rem_orders_df = orders_df.clone()
        roll_cuts = []
        iteration = 0
        failure_reason = "ไม่สามารถหาผลลัพธ์ที่เหมาะสมได้"
        final_status = None
        while not rem_orders_df.is_empty():
            iteration += 1
            if progress_callback:
                progress_callback(f"  Iteration {iteration}: Remaining orders: {rem_orders_df.shape[0]} items")

            orders_to_process = rem_orders_df
            if chunk_size and rem_orders_df.shape[0] > chunk_size:
                orders_to_process = rem_orders_df.sample(n=chunk_size, with_replacement=False, shuffle=True, seed=iteration)
                if progress_callback:
                    progress_callback(f"    Sampling {chunk_size} orders out of {rem_orders_df.shape[0]} for processing.")

            if c is None : c_type = None
            if b is None : b_type = None

            results_to_process, result = await _find_solution(
                orders_to_process, roll, c_type, b_type, progress_callback, orders_df
            )
            final_status = result.get("status")

            if not results_to_process:
                if progress_callback:
                    progress_callback(f"    ❌ {result.get('message', 'Non-optimal status')}")
                failure_reason = result.get('message', f'สถานะไม่เหมาะสม: {final_status}')
                break

            processed_order_indices_this_iteration = set()
            cut_infos_this_iteration = []
            last_rem_roll_l = roll['length']
            all_successful = True

            for res in results_to_process:
                cut_info, order_idx, failure_reason_from_process = await _process_single_order(
                    res, orders_df, order_num_col_idx, material_substitutions,
                    progress_callback, out_of_stock_handler, roll_specs,
                    used_roll_ids_for_cut, last_used_roll_ids, roll_width
                )

                if cut_info:
                    cut_infos_this_iteration.append(cut_info)
                    if order_idx is not None:
                        processed_order_indices_this_iteration.add(order_idx)
                    last_rem_roll_l = cut_info.get("rem_roll_l", last_rem_roll_l)
                else:
                    all_successful = False
                    failure_reason = failure_reason_from_process

            all_results.extend(cut_infos_this_iteration)
            roll_cuts.extend(cut_infos_this_iteration)
            roll['length'] = last_rem_roll_l

            if processed_order_indices_this_iteration:
                rem_orders_df = rem_orders_df.filter(~pl.col("original_idx").is_in(list(processed_order_indices_this_iteration)))

            if not all_successful:
                failure_reason = failure_reason or "One or more cuts in the group failed."
                break

            # If we used greedy nesting, it creates a full plan, so we can break from the main loop for this roll.
            if result.get("status") != STATUS_OPTIMAL and results_to_process:
                 break

        # Save results for the current roll to a sqlite file
        if roll_cuts:
            output_df = pl.DataFrame(roll_cuts)
            db_path = os.path.join(output_dir, "cache.db")
            conn_str = f"sqlite:///{os.path.abspath(db_path)}"
            table_name = f"roll_cut_results_{roll['width']}"
            output_df.write_database(table_name, connection=conn_str, if_table_exists="replace")
            if progress_callback:
                progress_callback(f"--- Saved {len(roll_cuts)} cuts for roll {roll['width']} to database table '{table_name}' ---")
        elif progress_callback:
            progress_callback(f"--- No cuts made for roll {roll['width']} ---")

        if not rem_orders_df.is_empty():
            roll_w_status = STATUS_FAILED
            if final_status == STATUS_INFEASIBLE:
                roll_w_status = STATUS_INFEASIBLE

            if progress_callback:
                progress_callback(
                    f"    Adding {rem_orders_df.shape[0]} {roll_w_status.lower()} orders to the results."
                )

            unprocessed_orders = rem_orders_df.to_dicts()
            for order in unprocessed_orders:
                unprocessed_result = _create_unprocessed_result(order, roll_w_status, failure_reason)
                all_results.append(unprocessed_result)

    # Save all cutting results to a single summary table in sqlite
    if all_results:
        final_output_df = pl.DataFrame(all_results)
        db_path = os.path.join(output_dir, "cache.db")
        conn_str = f"sqlite:///{os.path.abspath(db_path)}"
        final_output_df.write_database("all_cutting_plan_summary", connection=conn_str, if_table_exists="replace")
        if progress_callback:
            progress_callback("💾 บันทึกผลลัพธ์ลงฐานข้อมูลเรียบร้อย")

    return all_results
