import asyncio
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

import cleaning


def _find_and_update_roll(roll_specs: dict, width: str, material: str, required_length: float, used_roll_ids: set, last_used_roll_ids: dict) -> str:
    """
    Finds a suitable roll, prioritizing the last used roll for the same material to ensure sequential use.
    If one roll is not enough, it tries to combine with another available roll.
    """
    if not material or not width:
        return ""

    material_rolls_dict = roll_specs.get(str(width), {}).get(material, {})
    if not material_rolls_dict:
        return "-> (ไม่มีข้อมูลสต็อก)"

    # Get available rolls, sorted by length.
    all_available_rolls = sorted(material_rolls_dict.items(), key=lambda item: item[1]['length'])
    
    # Filter out already used rolls for this specific cut.
    unused_rolls = [
        (k, r) for k, r in all_available_rolls if r.get('id') not in used_roll_ids
    ]

    # --- New Logic: Prioritize last used roll for this material ---
    last_roll_id = last_used_roll_ids.get((width, material))
    if last_roll_id:
        # Find the last roll in the list of all rolls for this material.
        last_roll_data = next(((k, r) for k, r in material_rolls_dict.items() if r.get('id') == last_roll_id), None)

        if last_roll_data and last_roll_data[1]['length'] > 0 and last_roll_id not in used_roll_ids:
            _last_roll_key, last_roll = last_roll_data
            
            # Case 1: The last used roll is sufficient by itself.
            if last_roll['length'] >= required_length:
                original_length = last_roll['length']
                last_roll['length'] -= required_length
                used_roll_ids.add(last_roll_id)
                # The last used roll remains the same.
                return f"-> ใช้ม้วนต่อเนื่อง: {last_roll_id} (ยาว {int(original_length)} ม., เหลือ {int(last_roll['length'])} ม.)"
            
            # Case 2: The last used roll is not sufficient, try to combine it with another.
            else:
                needed_from_another = required_length - last_roll['length']
                original_len_roll1 = last_roll['length']
                
                # Find a supplementary roll from the unused rolls (excluding the last_roll itself).
                supplement_rolls = [(k, r) for k, r in unused_rolls if r.get('id') != last_roll_id]
                
                # Find the smallest roll that can cover the remaining requirement.
                best_supplement = next(( (k, r) for k, r in sorted(supplement_rolls, key=lambda item: item[1]['length']) if r['length'] >= needed_from_another), None)

                if best_supplement:
                    _supp_key, supp_roll = best_supplement
                    original_len_roll2 = supp_roll['length']
                    supp_roll_id = supp_roll['id']

                    # Update rolls: use last_roll completely, take remainder from supplement.
                    last_roll['length'] = 0
                    supp_roll['length'] -= needed_from_another
                    
                    used_roll_ids.add(last_roll_id)
                    used_roll_ids.add(supp_roll_id)
                    
                    # The new active roll is the supplement roll.
                    last_used_roll_ids[(width, material)] = supp_roll_id
                    
                    return (f"-> ใช้ม้วนต่อเนื่อง: {last_roll_id} (ยาว {int(original_len_roll1)} ม., ใช้หมด) "
                            f"+ {supp_roll_id} (ยาว {int(original_len_roll2)} ม., เหลือ {int(supp_roll['length'])} ม.)")

    # --- Fallback to original logic if last used roll wasn't applicable ---
    # 1. Try to find a single new roll that is sufficient.
    for roll_key, roll in unused_rolls:
        roll_id = roll.get('id')
        roll_length = roll.get('length', 0)
        if roll_id and roll_length >= required_length:
            roll['length'] -= required_length
            used_roll_ids.add(roll_id)
            # Set this as the new last used roll for this material.
            last_used_roll_ids[(width, material)] = roll_id
            return f"-> เปิดม้วนใหม่: {roll_id} (ยาว {int(roll_length)} ม., เหลือ {int(roll['length'])} ม.)"

    # 2. If no single roll is sufficient, try to combine two new rolls.
    if len(unused_rolls) >= 2:
        for i in range(len(unused_rolls) - 1):
            roll1_key, roll1 = unused_rolls[i]
            roll2_key, roll2 = unused_rolls[i+1]
            
            roll1_id = roll1.get('id')
            roll1_length = roll1.get('length', 0)
            roll2_id = roll2.get('id')
            roll2_length = roll2.get('length', 0)

            if roll1_id and roll2_id:
                combined_length = roll1_length + roll2_length
                if combined_length >= required_length:
                    needed_from_roll2 = required_length - roll1_length
                    
                    roll1['length'] = 0
                    roll2['length'] -= needed_from_roll2
                    
                    # The second roll in the combination becomes the new active roll.
                    last_used_roll_ids[(width, material)] = roll2_id
                    
                    used_roll_ids.add(roll1_id)
                    used_roll_ids.add(roll2_id)
                    
                    return (f"-> เปิดม้วนใหม่: {roll1_id} (ยาว {int(roll1_length)} ม., ใช้หมด) "
                            f"+ {roll2_id} (ยาว {int(roll2_length)} ม., เหลือ {int(roll2['length'])} ม.)")

    return "-> (ไม่มีสต็อกที่พอ)"


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
        return {"status": "Infeasible - No orders left", "message": "No available orders to cut from."}

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
        prob += z_width[j] <= 6, f"MaxCutsAcrossWidth_{j}" # TODO: This seems to be a hardcoded business rule.

        # If order type is 'X', limit z to 5 cuts
        if 'X' in (types[j], component_types[j]):
            prob += z <= 5 + M * (1 - y[j]), f"MaxZ_TypeX_{j}"

    total_cut_width = lpSum(widths[j] * z_width[j] for j in range(num_orders))

    # 4. Define objective function and related constraints
    corr_multiplier = CORRUGATE_MULTIPLIERS.get(most_demand_type, 1.0)
    
    # Total length of material required for the selected order (using .get with a default value)
    # The sum will effectively pick the one order where y[j]=1
    total_order_len = lpSum(
        (lengths[j] * 25.4 / 100 * quantities[j] * corr_multiplier * y[j])
        for j in range(num_orders)
    )

    # Objective: Minimize trim waste
    trim_waste = roll_width - total_cut_width
    prob += trim_waste, "MinimizeTrim"

    # Constraints
    prob += trim_waste >= 1, "TrimLowerBound"
    prob += trim_waste <= 5, "TrimUpperBound"
    
    # Remaining length on roll must be at least 100
    prob += roll_length * z - total_order_len >= 100, "RemainingLengthLowerBound"

    # 5. Solve the problem
    try:
        prob.solve(PULP_CBC_CMD(msg=False))
    except Exception as e:
        return {"status": "Solver Error", "message": f"Solver failed: {str(e)}"}
    
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
        },
        "material_specs": material_specs,
        "message": "PuLP problem solved successfully."
    }

async def main_algorithm(
    roll_width: int,
    roll_length: int,
    file_path: str = "order2024.csv",
    max_records: Optional[int] = 2000,
    progress_callback: Optional[Callable[[str], None]] = None,
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
):
    if progress_callback:
        progress_callback("⚙️ กำลังเริ่มการคำนวณ")

    orders_df = cleaning.clean_data(
        cleaning.load_data(file_path), 
        start_date,
        end_date,
        front=front,
        c=c if c_type in ["C", "E"] else None,
        middle=middle,
        b=b if b_type in ["B", "E"] else None,
        back=back,
    )
    if progress_callback:
        progress_callback("📁 โหลดและจัดเรียงข้อมูลเรียบร้อย")

    if max_records:
        orders_df = orders_df.head(max_records)
    orders_df = orders_df.with_row_index("original_idx")
    
    rolls = [{"width": roll_width, "length": roll_length}]
    all_results = []
    last_used_roll_ids = {}

    order_num_col_idx = orders_df.columns.index("order_number")

    for roll in rolls:
        if progress_callback:
            progress_callback(f"🔧 กำลังประมวลผลม้วน {roll['width']} นิ้ว")
        
        rem_orders_df = orders_df.clone()
        roll_cuts = []
        iteration = 0
        while not rem_orders_df.is_empty():
            iteration += 1
            if progress_callback:
                progress_callback(f"  Iteration {iteration}: Remaining orders: {rem_orders_df.shape[0]} items")

            if c is None : 
                c_type = None
            if b is None : 
                b_type = None         

            result = await solve_linear_program(
                roll['width'],
                roll['length'],
                rem_orders_df,
                c_type=c_type,
                b_type=b_type,
            )

            status = result.get("status")
            if status != "Optimal":
                if progress_callback:
                    progress_callback(f"    ❌ {result.get('message', 'Non-optimal status')}")
                break
            
            variables = result.get("variables", {})
            order_idx = variables.get("order_idx")

            if progress_callback:
                progress_callback(f"    Optimal solution found. Trim: {variables.get('trim', 0):.4f}")
                progress_callback(f"    Selected order width: {variables.get('order_w')} (Index: {order_idx}), Cuts: {variables.get('cuts')}")

            order_number = orders_df.row(int(order_idx))[order_num_col_idx] if order_idx is not None else None
            
            material_specs = result.get("material_specs", {})
            variables = result.get("variables", {})
            roll_info = {}
            if roll_specs:
                used_roll_ids_for_cut = set()
                roll_w_str = str(variables.get("roll_w", "")).strip()
                demand_per_cut = variables.get("demand_per_cut", 0)
                c_type_spec = material_specs.get('c_type')
                b_type_spec = material_specs.get('b_type')

                type_demand_divisor = 1.0
                if c_type_spec == 'C':
                    type_demand_divisor = CORRUGATE_MULTIPLIERS['C']
                elif b_type_spec == 'B':
                    type_demand_divisor = CORRUGATE_MULTIPLIERS['B']
                elif c_type_spec == 'E' or b_type_spec == 'E':
                    type_demand_divisor = CORRUGATE_MULTIPLIERS['E']

                if material_specs.get('front'):
                    material = str(material_specs.get('front')).strip()
                    value = demand_per_cut / type_demand_divisor
                    roll_info['front_roll_info'] = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids)

                if material_specs.get('c') and c_type_spec == 'C':
                    material = str(material_specs.get('c')).strip()
                    value = demand_per_cut
                    roll_info['c_roll_info'] = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids)
                elif material_specs.get('c') and c_type_spec == 'E':
                    material = str(material_specs.get('c')).strip()
                    value = demand_per_cut
                    if b_type_spec == 'B':
                        value = value / CORRUGATE_MULTIPLIERS['B'] * CORRUGATE_MULTIPLIERS['E']
                    roll_info['c_roll_info'] = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids)

                if material_specs.get('middle'):
                    material = str(material_specs.get('middle')).strip()
                    value = demand_per_cut / type_demand_divisor
                    roll_info['middle_roll_info'] = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids)

                if material_specs.get('b') and b_type_spec == 'B':
                    material = str(material_specs.get('b')).strip()
                    value = demand_per_cut
                    if c_type_spec == 'C':
                        value = (value / CORRUGATE_MULTIPLIERS['C']) * CORRUGATE_MULTIPLIERS['B']
                    roll_info['b_roll_info'] = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids)
                elif material_specs.get('b') and b_type_spec == 'E':
                    material = str(material_specs.get('b')).strip()
                    value = demand_per_cut
                    if c_type_spec == 'C':
                        value = (value / CORRUGATE_MULTIPLIERS['C']) * CORRUGATE_MULTIPLIERS['E']
                    roll_info['b_roll_info'] = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids)

                if material_specs.get('back'):
                    material = str(material_specs.get('back')).strip()
                    value = demand_per_cut / type_demand_divisor
                    roll_info['back_roll_info'] = _find_and_update_roll(roll_specs, roll_w_str, material, value, used_roll_ids_for_cut, last_used_roll_ids)

            cut_info = {
                "roll_w": variables.get("roll_w"),
                "rem_roll_l": variables.get("rem_roll_l"),
                "demand_per_cut": variables.get("demand_per_cut"),
                "order_number": order_number,
                "order_w": variables.get("order_w"),
                "order_l": variables.get("order_l"),
                "order_qty": variables.get("order_qty"),
                "order_dmd": variables.get("order_dmd"),
                "cuts": variables.get("cuts"),
                "trim": variables.get("trim"),
                "type": variables.get("type"),
                "component_type": variables.get("component_type"),
            }
            cut_info.update(material_specs)  # Add all material specs
            cut_info.update(roll_info)
            all_results.append(cut_info)
            roll_cuts.append(cut_info)

            roll['length'] = variables.get("rem_roll_l")

            if order_idx is not None:
                rem_orders_df = rem_orders_df.filter(pl.col("original_idx") != order_idx)
            else:
                if progress_callback:
                    progress_callback("    Warning: order_idx is None, cannot remove order. Stopping.")
                break

        # Save results for the current roll to a CSV file
        if roll_cuts:
            output_df = pl.DataFrame(roll_cuts)
            output_filename = f"roll_cut_results_{roll['width']}.csv"
            output_df.write_csv(output_filename)
            if progress_callback:
                progress_callback(f"--- Saved {len(roll_cuts)} cuts for roll {roll['width']} to {output_filename} ---")
        elif progress_callback:
            progress_callback(f"--- No cuts made for roll {roll['width']} ---")

    # Save all cutting results to a single summary CSV file
    if all_results:
        final_output_df = pl.DataFrame(all_results)
        final_output_df.write_csv("all_cutting_plan_summary.csv")
        if progress_callback:
            progress_callback("💾 บันทึกผลลัพธ์ลงไฟล์ CSV เรียบร้อย")
    
    return all_results
