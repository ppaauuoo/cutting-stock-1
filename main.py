import asyncio
import sys
from pathlib import Path
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

app = FastAPI()

async def solve_linear_program(roll_paper_width: int, roll_paper_length: int, orders_df: pl.DataFrame, C: Optional[str] = None, B: Optional[str] = None) -> dict:
    """
    Solve a simple Linear Programming problem using PuLP for a given roll paper width
    and available orders DataFrame.
    """
    # 1. Create the LP problem
    # The original objective seems to be related to minimizing trim waste
    # Therefore, change to LpMinimize
    prob = LpProblem(f"LP Problem for Roll {roll_paper_width} with {roll_paper_length} length", LpMinimize)

    # 2. Create decision variables
    if orders_df.is_empty():
        return {
            "status": "Infeasible - No orders left",
            "objective_value": None,
            "variables": {},
            "message": "No available orders to cut from."
        }

    orders_widths = orders_df['width'].to_list()
    orders_lengths = orders_df['length'].to_list()
    orders_quantity = orders_df['quantity'].to_list()
    orders_types = orders_df['type'].to_list() # ดึงข้อมูลประเภททับเส้น
    orders_comopnent_types = orders_df['component_type'].to_list() # ดึงข้อมูลประเภททับเส้น

    # Binary variable for selecting order width
    # y_order_selection[j] = 1 if orders_widths[j] is selected, 0 otherwise
    y_order_selection = LpVariable.dicts("select_order", range(len(orders_widths)), 0, 1, LpBinary)

    # Variable z represents the number of times the selected order is cut
    # Variable z represents the number of times the selected order is cut (across the width)
    # Determine a suitable upper bound for z based on roll_paper_width and minimum order width
    # แก้ไขการหา min_order_width_positive ให้เปลี่ยนเปอร์เซ็นต์เป็นจำนวนจริง
    non_zero_widths = [w for w in orders_widths if w > 0]
    min_order_width_positive = min(non_zero_widths) if non_zero_widths else 1
    max_possible_cuts_z = int(roll_paper_width / min_order_width_positive) if min_order_width_positive > 0 else 1000
    z = LpVariable("num_cuts", 0, max_possible_cuts_z, LpInteger) # z >= 0, should be an integer

    # M_UPPER_BOUND_Z for linearization of product z * y_order_selection[j]
    # This M should be the upper bound of z, so we use max_possible_cuts_z
    M_UPPER_BOUND_Z_FOR_LINEARIZATION = max_possible_cuts_z

    # Auxiliary variables for linearization of product z * y_order_selection[j]
    # z_effective_cut_width_part[j] will be equal to z if y_order_selection[j] is 1, otherwise it will be 0
    z_effective_cut_width_part = LpVariable.dicts("z_effective_cut_part", range(len(orders_widths)), 0, None)

    # 3. Define constraints
    # Select only one order length
    prob += lpSum(y_order_selection[j] for j in range(len(orders_widths))) == 1, "Constraint_SelectOneOrder"

    # Constraints for linearization of z * y_order_selection[j]
    # For each order j, z_effective_cut_width_part[j] = y_order_selection[j] * z
    for j in range(len(orders_widths)):
        prob += z_effective_cut_width_part[j] <= z, f"Linearize_Wj_1_{j}"
        prob += z_effective_cut_width_part[j] <= M_UPPER_BOUND_Z_FOR_LINEARIZATION * y_order_selection[j], f"Linearize_Wj_2_{j}"
        prob += z_effective_cut_width_part[j] >= z - M_UPPER_BOUND_Z_FOR_LINEARIZATION * (1 - y_order_selection[j]), f"Linearize_Wj_3_{j}"
        prob += z_effective_cut_width_part[j] >= 0, f"Linearize_Wj_4_{j}" # Ensure it's non-negative
        prob += z_effective_cut_width_part[j] <= 6, f"Linearize_Wj_5_{j}" 
        # prob += (orders_lengths[j]*orders_quantity[j])/z_effective_cut_width_part[j] <= 100, f"Linearize_Wj_5_{j}" 

        # เพิ่มเงื่อนไขการจำกัด z เมื่อประเภททับเส้นเป็น 'X'
        if orders_types[j] == 'X' or orders_comopnent_types[j] == 'X':
            # สร้าง constraint: หากเลือกออเดอร์นี้ (y==1) ให้ z <=5
            # ใช้ Big-M method: z <= 5 + M * (1 - y_order_selection[j])
            prob += (
                z <= 5 + M_UPPER_BOUND_Z_FOR_LINEARIZATION * (1 - y_order_selection[j]),
                f"MaxZConstraint_TypeX_Order_{j}"
            )

    # Define the selected roll paper width and the total length of cut orders
    selected_roll_width = roll_paper_width # Use the roll paper width passed as input directly
    selected_roll_length = roll_paper_length # Use the roll paper length passed as input directly
    effective_order_cut_width = lpSum(orders_widths[j] * z_effective_cut_width_part[j] for j in range(len(orders_widths)))
    effective_order_cut_length = lpSum(
        (
            (orders_lengths[j] * 25.4 / 100)
            * orders_quantity[j]
            * (
                (1.45 if C else 1.35 if B else 1.0)
            )
            * y_order_selection[j]
        )
        for j in range(len(orders_widths))
    )

    # 4. Define objective function
    # Objective: Minimize trim waste
    # Waste = selected_roll_width - effective_order_cut_width
    prob += selected_roll_width - effective_order_cut_width, "Objective_MinimizeTrim"

    # 5. Define constraints (originally: -1 < trim < 5)
    # Since trim should not be negative (cut length should not exceed roll paper length)
    # And to be a Linear Program, it must be split into 2 constraints
    prob += selected_roll_width - effective_order_cut_width >= 1, "Constraint_Trim_LowerBound"
    prob += selected_roll_width - effective_order_cut_width <= 5, "Constraint_Trim_UpperBound"
    
    prob += selected_roll_length * z - effective_order_cut_length  >= 100, "Constraint_Length_LowerBound"
    
    # Add constraint that cut length must not exceed roll paper length (waste must not be negative)
    prob += effective_order_cut_width <= selected_roll_width, "Constraint_CutsMustFit"

    # 6. Solve the problem (แก้ไขส่วนนี้)
    try:
        solver = PULP_CBC_CMD(msg=False)  # ระบุ solver อย่างชัดเจน
        prob.solve(solver)  # แก้จาก prob.solve() เป็น prob.solve(solver)
        
        # ตรวจสอบสถานะการแก้ปัญหา
        # if prob.status != LpStatus.Optimal:  # 1 = Optimal
        #     status_str = LpStatus[prob.status]
        #     return {
        #         "status": f"Solution Status: {status_str}",
        #         "objective_value": None,
        #         "variables": {},
        #         "message": f"Failed to find optimal solution: {status_str}"
        #     }
    except Exception as e:
        return {
            "status": "Solver Error",
            "objective_value": None,
            "variables": {},
            "message": f"Solver failed: {str(e)}"
        }
    
    # 7. Retrieve and format results
    return await _get_lp_solution_details(
        prob,
        y_order_selection,
        z,
        orders_df,
        roll_paper_width,
        roll_paper_length,
        effective_order_cut_length
    )

async def _get_lp_solution_details(
    prob: LpProblem,
    y_order_selection: dict,
    z: LpVariable,
    orders_df: pl.DataFrame,
    roll_paper_width: int,
    roll_paper_length: int,
    effective_order_cut_length: LpVariable
) -> dict:
    """
    Extracts and formats the results from the solved PuLP problem.
    """
    status = LpStatus[prob.status]
    objective_value = round(value(prob.objective), 4) if value(prob.objective) is not None else None

    selected_order_idx = -1
    # orders_df.shape[0] is the total number of orders in the DataFrame, which is equivalent to len(orders_widths)
    for j in range(orders_df.shape[0]):
        if y_order_selection[j].varValue == 1:
            selected_order_idx = j
            break

    actual_z_value = z.varValue if z.varValue is not None else 0

    actual_selected_roll_width = roll_paper_width
    # Access order details using orders_df directly
    orders_data = orders_df.to_dicts() # Convert to list of dictionaries for easier access by index
    actual_selected_order_width = orders_data[selected_order_idx]['width'] if selected_order_idx != -1 else None
    actual_selected_order_length = orders_data[selected_order_idx]['length'] if selected_order_idx != -1 else None
    actual_selected_order_quantity = orders_data[selected_order_idx]['quantity'] if selected_order_idx != -1 else None

    # Calculate actual_selected_order_demand from effective_order_cut_length
    # This represents the total length (length * quantity) for the selected order
    actual_selected_order_total_length = value(effective_order_cut_length) if selected_order_idx != -1 else 0

    # Calculate the demand per effective cut (as per user's request for calculation)
    # This calculation should only happen if actual_z_value is valid and not zero
    calculated_effective_demand_per_cut = None
    if actual_selected_order_length is not None and actual_selected_order_quantity is not None and actual_z_value is not None and actual_z_value > 0:
        calculated_effective_demand_per_cut = round(actual_selected_order_total_length / actual_z_value, 4)

    actual_remaining_roll_length = round(roll_paper_length - (calculated_effective_demand_per_cut if calculated_effective_demand_per_cut is not None else 0), 4)


    selected_order_original_index = None
    if selected_order_idx != -1:
        selected_order_original_index = orders_data[selected_order_idx]['original_idx']

    actual_trim = None
    if actual_selected_roll_width is not None and actual_selected_order_width is not None and actual_z_value is not None:
        actual_trim = round(actual_selected_roll_width - (actual_selected_order_width * actual_z_value), 4)

    # เพิ่มการดึงข้อมูลวัสดุจาก orders_df
    material_specs = {}
    if selected_order_idx != -1 and orders_data[selected_order_idx]:
        for key in ['front', 'C', 'middle', 'B', 'back']:
            # เปลี่ยนชื่อตัวแปร 'value' เป็น 'material_value' เพื่อหลีกเลี่ยงการชนกับฟังก์ชัน pulp.value
            material_value = orders_data[selected_order_idx].get(key)
            if material_value:
                material_specs[key] = material_value

    # เพิ่มการดึงข้อมูลประเภททับเส้นและชนิดส่วนประกอบ
    type_value = orders_data[selected_order_idx].get('type') if selected_order_idx != -1 else None
    component_type_value = orders_data[selected_order_idx].get('component_type') if selected_order_idx != -1 else None

    return {
        "status": status,
        "objective_value": objective_value,
        "variables": {
            "selected_roll_width": actual_selected_roll_width,
            "selected_roll_length": actual_remaining_roll_length, # This is the remaining length of the roll
            "demand": calculated_effective_demand_per_cut, # ความยาวที่ใช้ไปต่อการตัด 1 ครั้ง
            "selected_order_width": actual_selected_order_width,
            "selected_order_length": actual_selected_order_length,
            "selected_order_quantity": actual_selected_order_quantity,
            "num_cuts_z": actual_z_value,
            "calculated_trim": actual_trim, # Display calculated trim value
            "selected_order_original_index": selected_order_original_index, # Add original_idx of the selected order
            "type": type_value,  # เพิ่มฟิลด์นี้
            "component_type": component_type_value  # เพิ่มฟิลด์นี้
        },
        # เพิ่มฟิลด์ material_specs
        "material_specs": material_specs,
        "message": "PuLP problem solved successfully. Note: This is a linearized formulation for a mixed-integer linear program."
    }

async def main_algorithm(
    roll_width: int,
    roll_length: int,
    file_path: str = "order2024.csv", # เปลี่ยนเป็น order2024.csv เพราะจะโหลดและคลีนเอง
    max_records: Optional[int] = 2000,
    progress_callback: Optional[Callable[[str], None]] = None,
    start_date: Optional[str] = None,  # เพิ่มพารามิเตอร์นี้
    end_date: Optional[str] = None,    # เพิ่มพารามิเตอร์นี้
    front: Optional[str] = None,       # เพิ่มพารามิเตอร์สำหรับกรองวัสดุ
    C: Optional[str] = None,
    middle: Optional[str] = None,
    B: Optional[str] = None,
    back: Optional[str] = None,
):
    if progress_callback:
        progress_callback("⚙️ กำลังเริ่มการคำนวณ")

    # โหลดและคลีนข้อมูลทั้งหมดพร้อมเพิ่ม original index column
    full_orders_df = cleaning.clean_data(
        cleaning.load_data(file_path), 
        start_date,
        end_date,
        front=front, # ส่งพารามิเตอร์วัสดุไปยัง cleaning.clean_data
        C=C,
        middle=middle,
        B=B,
        back=back,
    )
    if progress_callback:
        progress_callback("📁 โหลดและจัดเรียงข้อมูลเรียบร้อย")

    if max_records:
        full_orders_df = full_orders_df.head(max_records)
    full_orders_df = full_orders_df.with_row_index("original_idx")
    
    # The widths of the paper rolls we have (can be changed or received from input)
    ROLL_PAPER = [{"width": roll_width, "length": roll_length}]

    all_cut_results = [] # Store all results from cutting all rolls

    for roll in ROLL_PAPER:
        if progress_callback:
            progress_callback(f"🔧 กำลังประมวลผลม้วน {roll['width']} นิ้ว")
        
        # Create a copy of the full orders DataFrame for the current roll
        remaining_orders_df = full_orders_df.clone()
        current_roll_cuts = [] # Store cutting results for the current roll

        iteration = 0
        # Loop until no more orders can be cut from the current roll
        while not remaining_orders_df.is_empty():
            iteration += 1
            if progress_callback:
                progress_callback(f"  Iteration {iteration}: Remaining orders: {remaining_orders_df.shape[0]} items")

            # Call the LP solving function
            result = await solve_linear_program(roll['width'], roll['length'], remaining_orders_df, C, B)

            status = result["status"]
            
            # แจ้งข้อผิดพลาดหากมี
            if "Error" in status or "Failed" in status or "Infeasible" in status:
                if progress_callback:
                    progress_callback(f"    ❌ {result['message']}")
                break
            
            selected_order_original_index = result["variables"].get("selected_order_original_index")

            if status == "Optimal":
                if progress_callback:
                    progress_callback(f"    Optimal solution found for roll {roll['width']}. Trim waste: {result['variables']['calculated_trim']:.4f}")
                    progress_callback(f"    Selected order width: {result['variables']['selected_order_width']} (Original Index: {selected_order_original_index}), Number of cuts: {result['variables']['num_cuts_z']}")

                # Store current cut information
                # Get the order_number using the selected_order_original_index
                order_number = None
                if selected_order_original_index is not None:
                    # Get the index of the "order_number" column
                    order_number_col_idx = full_orders_df.columns.index("order_number")
                    order_number = full_orders_df.row(int(selected_order_original_index))[order_number_col_idx]
                
                # ดึงค่า actual_z_value จาก result["variables"]
                actual_z_value = result["variables"].get("num_cuts_z")

                cut_info = {
                    "roll width": result["variables"]["selected_roll_width"],
                    "roll length": result["variables"]["selected_roll_length"], # ความยาวคงเหลือของม้วน
                    "demand": result["variables"]["demand"], # ความยาวที่ใช้ไปต่อการตัด 1 ครั้ง
                    "order_number": order_number,
                    "selected_order_width": result["variables"]["selected_order_width"],
                    "selected_order_length": result["variables"]["selected_order_length"],
                    "selected_order_quantity": result["variables"]["selected_order_quantity"],
                    "num_cuts_z": actual_z_value, # ใช้ actual_z_value ที่ดึงมา
                    "calculated_trim": result["variables"]["calculated_trim"],
                    # เพิ่มข้อมูลวัสดุจาก result
                    "front": result.get("material_specs", {}).get("front"),
                    "C": result.get("material_specs", {}).get("C"),
                    "middle": result.get("material_specs", {}).get("middle"),
                    "B": result.get("material_specs", {}).get("B"),
                    "back": result.get("material_specs", {}).get("back"),
                    "type": result["variables"].get("type"),  # เพิ่มฟิลด์นี้
                    "component_type": result["variables"].get("component_type")  # เพิ่มฟิลด์นี้
                }
                all_cut_results.append(cut_info)
                current_roll_cuts.append(cut_info)

                roll['length'] = result["variables"]["selected_roll_length"] # Update roll length after cuts

                # Remove the selected order from the remaining orders DataFrame
                if selected_order_original_index is not None:
                    remaining_orders_df = remaining_orders_df.filter(
                        pl.col("original_idx") != selected_order_original_index
                    )
                else:
                    if progress_callback:
                        progress_callback("    Warning: selected_order_original_index is None, cannot remove order.")
                    break # Exit loop if order cannot be identified and removed

            elif "No orders left" in status:
                if progress_callback:
                    progress_callback(f"    No orders left for roll {roll['width']}.")
                break # No more orders to process
            else:
                if progress_callback:
                    progress_callback(f"    Problem is infeasible or status is undefined for roll {roll['width']}.")
                    progress_callback(f"    Status: {status}")
                break # Cannot find solution for remaining orders with this roll

        # Save results for the current roll to a CSV file if cuts occurred
        if current_roll_cuts:
            output_df = pl.DataFrame(current_roll_cuts)
            output_filename = f"roll_cut_results_{roll['width']}.csv"
            output_df.write_csv(output_filename)
            if progress_callback:
                progress_callback(f"--- Saved {len(current_roll_cuts)} cut items for roll {roll['width']} to {output_filename} ---")
        else:
            if progress_callback:
                progress_callback(f"--- No cuts made for roll {roll['width']} ---")

        # Update the roll width based on the last selected order demand
        # จุดนี้อาจจะผิด ถ้าต้องการอัปเดตความกว้างของม้วนกระดาษที่เหลือ
        # แต่เดิมอัปเดตจาก selected_order_demand ซึ่งเป็นความยาว ไม่ใช่ความกว้าง
        # การเปลี่ยน width ของ roll_paper ในลูปนี้อาจไม่ถูกต้องตามวัตถุประสงค์
        # หากต้องการลดความยาวของม้วนที่เหลือ ต้องอัปเดต roll['length']
        # selected_order_demand = result["variables"].get("selected_order_demand")
        # roll['width'] = roll['width'] - selected_order_demand if selected_order_demand is not None else roll['width']
        # ความยาวที่เหลือของม้วนจะถูกคำนวณและส่งกลับใน result['variables']['selected_roll_length'] แล้ว
        # ดังนั้นไม่จำเป็นต้องอัปเดต roll['width'] ตรงนี้

    # Optional: Save all cutting results from all rolls to a single CSV file
    if all_cut_results:
        final_output_df = pl.DataFrame(all_cut_results)
        final_output_df.write_csv("all_cutting_plan_summary.csv")
        if progress_callback:
            progress_callback("💾 บันทึกผลลัพธ์ลงไฟล์ CSV เรียบร้อย")
    
    return all_cut_results

if __name__ == "__main__":
    # Default values for CLI execution
    default_roll_width = 75
    default_roll_length = 111175
    default_file_path = "order2024.csv" # เปลี่ยนเป็น order2024.csv
    default_max_records = 200
    default_start_date = None # เพิ่มค่าเริ่มต้น
    default_end_date = None   # เพิ่มค่าเริ่มต้น
    default_front = None      # เพิ่มค่าเริ่มต้นสำหรับวัสดุ
    default_C = None
    default_middle = None
    default_B = None
    default_back = None

    print(f"Running cutting optimization with default parameters:")
    print(f"  Roll Width: {default_roll_width}")
    print(f"  Roll Length: {default_roll_length}")
    print(f"  Order File: {default_file_path}")
    print(f"  Max Records: {default_max_records}")
    print(f"  Start Date: {default_start_date}")
    print(f"  End Date: {default_end_date}")
    print(f"  Front Material: {default_front}") # แสดงค่าเริ่มต้นสำหรับวัสดุ
    print(f"  C Material: {default_C}")
    print(f"  Middle Material: {default_middle}")
    print(f"  B Material: {default_B}")
    print(f"  Back Material: {default_back}")

    asyncio.run(main_algorithm(
        roll_width=default_roll_width,
        roll_length=default_roll_length,
        file_path=default_file_path,
        max_records=default_max_records,
        progress_callback=print, # Use print for CLI output
        start_date=default_start_date,
        end_date=default_end_date,
        front=default_front, # ส่งค่าเริ่มต้นสำหรับวัสดุ
        C=default_C,
        middle=default_middle,
        B=default_B,
        back=default_back,
    ))
