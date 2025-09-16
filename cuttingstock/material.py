from typing import Optional, Callable
import copy
from cuttingstock.utils import log_message
import polars as pl

STATUS_FAILED = "Failed"
CORRUGATE_MULTIPLIERS = {
    "C": 1.45,
    "B": 1.35,
    "E": 1.25,
}

class OutOfStockError(Exception):
    """Custom exception for out-of-stock events."""
    def __init__(self, message, width, material, required_length, material_specs=None, known_out_of_stock=None):
        super().__init__(message)
        self.width = width
        self.material = material
        self.required_length = required_length
        self.material_specs = material_specs or {}
        self.known_out_of_stock = known_out_of_stock or []


def _spec_to_key(spec: dict) -> tuple:
    """Converts a spec dictionary to a hashable tuple key."""
    if not spec:
        return tuple()
    # Filter out None values and keys that are not material names
    valid_keys = {'front', 'c', 'middle', 'b', 'back'}
    return tuple(sorted((k, v) for k, v in spec.items() if k in valid_keys and v))

def _format_roll_message(roll_id: str, original_length: float, remaining_length: float) -> str:
    """Format a consistent roll usage message."""
    if remaining_length > 0:
        return f"{roll_id} (ยาว {int(original_length)} ม., เหลือ {int(remaining_length)} ม.)"
    else:
        return f"{roll_id} (ยาว {int(original_length)} ม., ใช้หมด)"

def _process_roll_usage(rolls: list, required_length: float, used_roll_ids: set) -> tuple[list[str], float, Optional[str]]:
    """Process roll usage and return message parts, remaining length, and final roll ID."""
    message_parts = []
    remaining_needed = required_length
    final_roll_id = None

    for roll_key, roll in rolls:
        roll_id = roll.get('id')
        roll_length = roll.get('length', 0)
        used_roll_ids.add(roll_id)

        if remaining_needed > 0:
            if roll_length >= remaining_needed:
                roll['length'] -= remaining_needed
                message_parts.append(_format_roll_message(roll_id, roll_length, roll['length']))
                final_roll_id = roll_id
                remaining_needed = 0
            else:
                roll['length'] = 0
                message_parts.append(_format_roll_message(roll_id, roll_length, 0))
                remaining_needed -= roll_length
                if remaining_needed == 0:
                    final_roll_id = roll_id
        else:
            break

    return message_parts, remaining_needed, final_roll_id

def _find_and_update_roll(roll_specs: dict, width: str, material: str, required_length: float, used_roll_ids: set, last_used_roll_ids: dict, order_number: Optional[str] = None, material_specs: Optional[dict] = None, group_id: Optional[str] = None, positions: Optional[dict[str, int]] = None, roll_positions: Optional[dict[str, tuple[str, str, str]]] = None, spec_key: Optional[str] = None) -> str:
    """
    Args:
        roll_specs (dict): Dictionary of roll specifications and stock.
        width (str): Width of the roll.
        material (str): Material of the roll.
        required_length (float): Required length of the roll.
        used_roll_ids (set): Set of all used roll IDs.
        last_used_roll_ids (dict): Dictionary of last used roll IDs.
        order_number (Optional[str]): Order number.
        material_specs (Optional[dict]): Dictionary of material specifications.
        group_id (Optional[str]): Group ID for order grouping.
        positions (Optional[dict]): Dictionary tracking order/group positions.
        roll_positions (Optional[dict]): Dictionary tracking roll positions per material.
        spec_key (Optional[str]): Specification key (front, c, middle, b, back) for material context.

    Returns:
        str: Message indicating the roll to be used.

    Raises:
        OutOfStockError: If no suitable roll is found.

    Keep track of sequences of orders base on order_number, width and material and
    finds a suitable roll by trying different strategies in order of priority:
    1. Continue using the last used roll for the same material on the same width and position.
    2. Open a new roll if the last used roll if out of stock or not exists.
    3. Reset position if material is different from last used material.

    Roll position tracking counts the number of rolls used for each material combination.
    """
    if not material or not width:
        return ""

    material_rolls_dict = roll_specs.get(str(width), {}).get(material, {})
    if not material_rolls_dict:
        log_message("error", "Out of stock: No stock data available for material.", {"width": width, "material": material, "required_length": required_length, "material_specs": material_specs})
        raise OutOfStockError("ไม่มีข้อมูลสต็อก", width, material, required_length, material_specs)

    # Get available rolls sorted by length
    all_available_rolls = sorted(material_rolls_dict.items(), key=lambda item: item[1]['length'], reverse=True)
    unused_rolls = [(k, r) for k, r in all_available_rolls if r.get('id') not in used_roll_ids]

    # Try to find a suitable roll combination
    rolls_for_combination = []
    combined_length = 0

    # --- State management for roll usage ---
    current_order_id = group_id if group_id else order_number

    if positions is None:
        positions = {}
    if roll_positions is None:
        roll_positions = {}

    spec_key_for_tracking = spec_key or 'unknown'
    # (group_id, CM127, front)
    material_key = (current_order_id, material, spec_key_for_tracking)
    # eg. same group, same material and spec -> new order within same group
    if material_key in roll_positions:
        return 'กลุ่มเดียวกัน'
    else:
        roll_positions[material_key] = 1

    if current_order_id in positions:
        position = positions[current_order_id] + 1
    else:
        position = 0

    position_key = (width, material, position)
    last_used_roll_id = last_used_roll_ids.get(position_key)

    # Check if we can continue using the last roll (either alone or in combination)
    if last_used_roll_id:
        last_roll_data = next(((k, r) for k, r in material_rolls_dict.items() if r.get('id') == last_used_roll_id), None)
        if last_roll_data:
            last_roll = last_roll_data[1]
            last_roll_length = last_roll['length']

            if last_roll_length >= required_length:
                # Use the existing roll alone
                last_roll['length'] -= required_length
                used_roll_ids.add(last_used_roll_id)
                positions[current_order_id] = position
                # Update roll position (reusing existing roll, no increment)
                return f"-> ใช้ม้วนต่อเนื่อง: {_format_roll_message(last_used_roll_id, last_roll_length, last_roll['length'])}"
            elif last_roll_length > 0:
                # Use last roll in combination with new rolls
                remaining_needed = required_length - last_roll_length

                # Find additional rolls to meet the remaining need
                supplement_rolls = []
                supplement_length = 0

                for roll_key, roll in unused_rolls:
                    if roll.get('id') != last_used_roll_id:
                        supplement_rolls.append((roll_key, roll))
                        supplement_length += roll.get('length', 0)
                        if supplement_length >= remaining_needed:
                            break

                if supplement_length >= remaining_needed:
                    # We have enough with combination
                    message_parts = [f"-> ใช้ม้วนต่อเนื่อง: {_format_roll_message(last_used_roll_id, last_roll_length, 0)}"]
                    last_roll['length'] = 0
                    used_roll_ids.add(last_used_roll_id)

                    # Process supplement rolls using helper function
                    supp_message_parts, remaining_supplement_needed, final_roll_id = _process_roll_usage(
                        supplement_rolls, remaining_needed, used_roll_ids
                    )
                    message_parts.extend(supp_message_parts)

                    # Update tracking
                    last_used_roll_ids[position_key] = final_roll_id
                    positions[current_order_id] = position
                    # Update roll position (increment because we used new rolls)

                    return " + ".join(message_parts)

    # Find combination of new rolls (excluding the last used roll which was already considered)
    for roll_key, roll in unused_rolls:
        if roll.get('id') != last_used_roll_id:
            rolls_for_combination.append((roll_key, roll))
            combined_length += roll.get('length', 0)
            if combined_length >= required_length:
                break

    # Check if we have enough length
    if combined_length < required_length:
        log_message("error", "Out of stock: Not enough stock length available for material.", {"width": width, "material": material, "required_length": required_length, "material_specs": material_specs})
        raise OutOfStockError("ไม่มีสต็อกที่พอ", width, material, required_length, material_specs)

    # Process the selected rolls using helper function
    message_parts, remaining_needed, new_last_used_roll_id = _process_roll_usage(
        rolls_for_combination, required_length, used_roll_ids
    )

    # Update last used roll tracking and positions
    if new_last_used_roll_id:
        last_used_roll_ids[position_key] = new_last_used_roll_id
        positions[current_order_id] = position
        # Update roll position (increment because we used new rolls)

    return f"-> เปิดม้วนใหม่: " + " + ".join(message_parts)

async def process_single_order(
    result: dict, orders_df: pl.DataFrame, order_num_col_idx: int, material_substitutions: dict,
    progress_callback: Optional[Callable[[str], None]], out_of_stock_handler: Optional[Callable],
    roll_specs: dict, used_roll_ids_for_cut: set, last_used_roll_ids: dict, roll_positions: dict, roll_width: int
) -> tuple[Optional[dict], Optional[int], Optional[str]]:
    """
    Processes a single order solution, handling stock checks, material substitutions, and roll allocation.
    Returns the final cut information, the processed order index, and any failure reason.
    """
    variables = result.get("variables", {})
    group_id = result.get("group_id", None)
    order_idx = variables.get("order_idx")
    roll_w_str = str(variables.get("roll_w", "")).strip()
    positions = {}  # Initialize positions tracking for order grouping

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

        # look for front,c,middle,b,back in order
        spec_key_for_lookup = _spec_to_key(material_specs_for_order)

        # if OutOfStockError happened in the previous round, material_substitutions will be set
        if material_substitutions and spec_key_for_lookup in material_substitutions:
            # check if this is a valid substitution for this material.
            substituted_spec = material_substitutions[spec_key_for_lookup]
            if substituted_spec is None:
                if progress_callback: progress_callback("    ❌ User previously cancelled substitution for this spec. Failing order.")
                break
            if progress_callback:
                changes_str = ", ".join([f"{k.title()}: {v}" for k, v in substituted_spec.items() if material_specs_for_order.get(k) != v])
                progress_callback(f"    🔄 Applying stored substitution for spec: {changes_str}")
            # set new spec.
            material_specs_for_order = substituted_spec.copy()

        current_attempt_specs = material_specs_for_order.copy()
        roll_info_this_attempt = {}
        spec_changed_this_attempt = False
        calculation_failed_reason = None
        roll_specs_backup = copy.deepcopy(roll_specs)
        last_used_roll_ids_backup = copy.deepcopy(last_used_roll_ids)
        roll_positions_backup = copy.deepcopy(roll_positions)

        def get_roll_for_material(spec_key: str, value_calculator: Callable[[], float]):
            nonlocal calculation_failed_reason, spec_changed_this_attempt, material_specs_for_order
            nonlocal roll_w_str
            value = value_calculator()
            material = str(current_attempt_specs.get(spec_key, "")).strip()
            if calculation_failed_reason or not material or not value: return
            try:
                info = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids, order_number, current_attempt_specs, group_id=group_id, positions=positions, roll_positions=roll_positions, spec_key=spec_key)
                roll_info_this_attempt[f'{spec_key}_roll_info'] = info
            except OutOfStockError as e:
                known_out_of_stock_materials.append((e.width, e.material))
                if out_of_stock_handler:
                    if progress_callback: progress_callback(f"    ⚠️ สต็อกสำหรับ '{e.material}' (หน้ากว้าง {e.width}) ไม่พอ, รอการตัดสินใจจากผู้ใช้...")
                    log_message("info", "Out of stock, awaiting user interaction.", {"width": e.width, "material": e.material, "required_length": e.required_length, "material_specs": e.material_specs})
                    # get new set of materials from handler(user)
                    new_material_specs = out_of_stock_handler(e)
                    if new_material_specs:
                        changes = {k: v for k, v in new_material_specs.items()}
                        log_message("info", "User provided material substitution.", {"original_specs": current_attempt_specs, "new_specs": new_material_specs, "changes": changes})
                        if progress_callback:
                            changes_str = ", ".join([f"{k.title()}: {v}" for k, v in changes.items()])
                            progress_callback(f"    ✅ User chose: {changes_str}. Will retry order.")
                        original_spec_key = _spec_to_key(current_attempt_specs)
                        material_substitutions[original_spec_key] = new_material_specs
                        for key, value in list(material_substitutions.items()):
                            if value is not None and _spec_to_key(value) == original_spec_key:
                                material_substitutions[key] = new_material_specs
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
                        calculation_failed_reason = "ผู้ใช้ยกเลิก"
                else:
                    fail_reason_msg = e.args[0]
                    calculation_failed_reason = fail_reason_msg

        if roll_specs:
            demand_per_cut = variables.get("demand_per_cut", 0)
            c_type_spec = current_attempt_specs.get('c_type')
            b_type_spec = current_attempt_specs.get('b_type')
            type_demand_divisor = CORRUGATE_MULTIPLIERS.get(c_type_spec or b_type_spec) or 1.0
            get_roll_for_material('front', lambda: demand_per_cut / type_demand_divisor)
            if c_type_spec == 'C': get_roll_for_material('c', lambda: demand_per_cut)
            elif c_type_spec == 'E': get_roll_for_material('c', (lambda: demand_per_cut / CORRUGATE_MULTIPLIERS['B'] * CORRUGATE_MULTIPLIERS['E']) if b_type_spec == 'B' else (lambda: demand_per_cut))
            get_roll_for_material('middle', lambda: demand_per_cut / type_demand_divisor)
            if b_type_spec == 'B': get_roll_for_material('b', (lambda: (demand_per_cut / CORRUGATE_MULTIPLIERS['C']) * CORRUGATE_MULTIPLIERS['B']) if c_type_spec == 'C' else (lambda: demand_per_cut))
            elif b_type_spec == 'E': get_roll_for_material('b', (lambda: (demand_per_cut / CORRUGATE_MULTIPLIERS['C']) * CORRUGATE_MULTIPLIERS['E']) if c_type_spec == 'C' else (lambda: demand_per_cut))
            get_roll_for_material('back', lambda: demand_per_cut / type_demand_divisor)

        if spec_changed_this_attempt:
            roll_specs.clear()
            roll_specs.update(roll_specs_backup)
            last_used_roll_ids.clear()
            last_used_roll_ids.update(last_used_roll_ids_backup)
            roll_positions.clear()
            roll_positions.update(roll_positions_backup)
            used_roll_ids_for_cut.clear()
            if calculation_failed_reason == "SPEC_CHANGED": calculation_failed_reason = None
            if progress_callback: progress_callback("    🔄 Spec changed, restarting roll allocation for this order...")
            continue
        if calculation_failed_reason: break
        final_roll_info = roll_info_this_attempt
        material_specs = current_attempt_specs
        order_processed_successfully = True

    if not order_processed_successfully:
        if progress_callback: progress_callback(f"    ❌ การคำนวณสำหรับ {order_number} ล้มเหลวเนื่องจาก: {calculation_failed_reason}")
        log_message("error", "Calculation failed", {"order_number": order_number, "reason": calculation_failed_reason})
        return None, order_idx

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
    return cut_info, order_idx

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


def handle_unprocessed_orders(
    rem_orders_df: pl.DataFrame,
    progress_callback: Optional[Callable[[str], None]]
) -> list:
    if rem_orders_df.is_empty():
        return []

    roll_w_status = failure_reason = STATUS_FAILED

    if progress_callback:
        progress_callback(
            f"    Adding {rem_orders_df.shape[0]} {roll_w_status.lower()} orders to the results."
        )

    unprocessed_results = []
    unprocessed_orders = rem_orders_df.to_dicts()
    for order in unprocessed_orders:
        unprocessed_result = _create_unprocessed_result(order, roll_w_status, failure_reason)
        unprocessed_results.append(unprocessed_result)
    return unprocessed_results
