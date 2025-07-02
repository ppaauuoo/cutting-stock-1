import asyncio
from pathlib import Path

import polars as pl
from fastapi import FastAPI
from pulp import (
    LpBinary,
    LpInteger,
    LpMinimize,
    LpProblem,
    LpStatus,
    LpVariable,
    lpSum,
    value,
)

app = FastAPI()

# Remove @app.get("/solve_lp") because we will call it directly from main() not via a FastAPI endpoint
# @app.get("/solve_lp")
async def solve_linear_program(roll_paper_width: int, available_orders_df: pl.DataFrame):
    """
    Solve a simple Linear Programming problem using PuLP for a given roll paper width
    and available orders DataFrame.
    """
    # 1. Create the LP problem
    # The original objective seems to be related to minimizing trim waste
    # Therefore, change to LpMinimize
    prob = LpProblem(f"LP Problem for Roll {roll_paper_width}", LpMinimize)

    # 2. Create decision variables
    if available_orders_df.is_empty():
        return {
            "status": "Infeasible - No orders left",
            "objective_value": None,
            "variables": {},
            "message": "No available orders to cut from."
        }

    orders_widths = available_orders_df['width'].to_list()

    # Binary variable for selecting order width
    # y_order_selection[j] = 1 if orders_widths[j] is selected, 0 otherwise
    y_order_selection = LpVariable.dicts("select_order", range(len(orders_widths)), 0, 1, LpBinary)

    # Variable z represents the number of times the selected order is cut
    z = LpVariable("num_cuts", 0, 7, LpInteger) # z >= 0, should be an integer

    # # M_UPPER_BOUND_Z is a sufficiently large number for Big-M linearization
    # # Should be chosen appropriately for the maximum possible value of z
    # # Calculate M_UPPER_BOUND_Z from roll_paper_width divided by the minimum positive value of orders_widths
    # # help when the minimum order only need 4-5 cut
    min_order_len_positive = min(val for val in orders_widths if val > 0) if any(val > 0 for val in orders_widths) else 1
    M_UPPER_BOUND_Z = roll_paper_width / min_order_len_positive
    if M_UPPER_BOUND_Z == 0: M_UPPER_BOUND_Z = 1000 # Fallback for safety (e.g., if roll_paper_width is 0 or orders_widths are all 0)

    # # Auxiliary variables for linearization of product z * y_order_selection[j]
    # # z_effective_cut_width_part[j] will be equal to z if y_order_selection[j] is 1, otherwise it will be 0
    z_effective_cut_width_part = LpVariable.dicts("z_effective_cut_part", range(len(orders_widths)), 0, None)

    # 3. Define constraints
    # Select only one order length
    prob += lpSum(y_order_selection[j] for j in range(len(orders_widths))) == 1, "Constraint_SelectOneOrder"

    # Constraints for linearization of z * y_order_selection[j]
    # For each order j, z_effective_cut_width_part[j] = y_order_selection[j] * z
    for j in range(len(orders_widths)):
        prob += z_effective_cut_width_part[j] <= z, f"Linearize_Wj_1_{j}"
        prob += z_effective_cut_width_part[j] <= M_UPPER_BOUND_Z * y_order_selection[j], f"Linearize_Wj_2_{j}"
        prob += z_effective_cut_width_part[j] >= z - M_UPPER_BOUND_Z * (1 - y_order_selection[j]), f"Linearize_Wj_3_{j}"
        prob += z_effective_cut_width_part[j] >= 0, f"Linearize_Wj_4_{j}" # Ensure it's non-negative

    # Define the selected roll paper width and the total length of cut orders
    selected_roll_width = roll_paper_width # Use the roll paper width passed as input directly
    effective_order_cut_width = lpSum(orders_widths[j] * z_effective_cut_width_part[j] for j in range(len(orders_widths)))

    # 4. Define objective function
    # Objective: Minimize trim waste
    # Waste = selected_roll_width - effective_order_cut_width
    prob += selected_roll_width - effective_order_cut_width, "Objective_MinimizeTrim"

    # 5. Define constraints (originally: -1 < trim < 5)
    # Since trim should not be negative (cut length should not exceed roll paper length)
    # And to be a Linear Program, it must be split into 2 constraints
    prob += selected_roll_width - effective_order_cut_width >= 1, "Constraint_Trim_LowerBound"
    prob += selected_roll_width - effective_order_cut_width <= 5, "Constraint_Trim_UpperBound"
    
    # Add constraint that cut length must not exceed roll paper length (waste must not be negative)
    prob += effective_order_cut_width <= selected_roll_width, "Constraint_CutsMustFit"

    # 6. Solve the problem
    prob.solve()

    # 7. Retrieve results
    status = LpStatus[prob.status]
    objective_value = value(prob.objective)

    # Retrieve values of selected variables
    selected_order_idx = -1
    for j in range(len(orders_widths)):
        if y_order_selection[j].varValue == 1:
            selected_order_idx = j
            break

    actual_z_value = z.varValue if z.varValue is not None else 0

    # Calculate actual selected_roll_width and selected_order_width
    actual_selected_roll_width = roll_paper_width # Use the roll_paper_width passed as input
    actual_selected_order_width = orders_widths[selected_order_idx] if selected_order_idx != -1 else None
    
    # Retrieve original_idx of the selected order
    selected_order_original_index = None
    if selected_order_idx != -1:
        selected_order_original_index = available_orders_df['original_idx'].to_list()[selected_order_idx]

    # Calculate actual Trim
    actual_trim = None
    if actual_selected_roll_width is not None and actual_selected_order_width is not None and actual_z_value is not None:
        actual_trim = actual_selected_roll_width - (actual_selected_order_width * actual_z_value)

    return {
        "status": status,
        "objective_value": objective_value,
        "variables": {
            "selected_roll_width": actual_selected_roll_width,
            "selected_order_width": actual_selected_order_width,
            "num_cuts_z": actual_z_value,
            "calculated_trim": actual_trim, # Display calculated trim value
            "selected_order_original_index": selected_order_original_index # Add original_idx of the selected order
        },
        "message": "PuLP problem solved successfully. Note: This is a linearized formulation for a mixed-integer linear program."
    }

async def main():
    # Load all order data and add original index column
    full_orders_df = pl.read_csv("clean_order2024.csv").head(200)
    full_orders_df = full_orders_df.with_row_index("original_idx")
    
    # The widths of the paper rolls we have (can be changed or received from input)
    # ROLL_PAPER_WIDTHS = [66, 68, 70, 73, 74, 75, 79, 82, 85, 88, 91, 93, 95, 97]
    ROLL_PAPER_WIDTHS = [75]

    all_cut_results = [] # Store all results from cutting all rolls

    for roll_width in ROLL_PAPER_WIDTHS:
        print(f"\n--- Processing for roll paper width: {roll_width} ---")
        # Create a copy of the full orders DataFrame for the current roll
        remaining_orders_df = full_orders_df.clone()
        current_roll_cuts = [] # Store cutting results for the current roll

        iteration = 0
        # Loop until no more orders can be cut from the current roll
        while not remaining_orders_df.is_empty():
            iteration += 1
            print(f"  Iteration {iteration}: Remaining orders: {remaining_orders_df.shape[0]} items")

            # Call the LP solving function
            result = await solve_linear_program(roll_width, remaining_orders_df)

            status = result["status"]
            selected_order_original_index = result["variables"].get("selected_order_original_index")

            if status == "Optimal":
                print(f"    Optimal solution found for roll {roll_width}. Trim waste: {result['variables']['calculated_trim']:.2f}")
                print(f"    Selected order width: {result['variables']['selected_order_width']} (Original Index: {selected_order_original_index}), Number of cuts: {result['variables']['num_cuts_z']}")

                # Store current cut information
                # Get the order_number using the selected_order_original_index
                order_number = None
                if selected_order_original_index is not None:
                    # Get the index of the "order_number" column
                    order_number_col_idx = full_orders_df.columns.index("order_number")
                    order_number = full_orders_df.row(int(selected_order_original_index))[order_number_col_idx]
                cut_info = {
                    "roll_width": roll_width,
                    "order_number": order_number,
                    "selected_order_width": result["variables"]["selected_order_width"],
                    "num_cuts_z": result["variables"]["num_cuts_z"],
                    "calculated_trim": result["variables"]["calculated_trim"],
                }
                all_cut_results.append(cut_info)
                current_roll_cuts.append(cut_info)

                # Remove the selected order from the remaining orders DataFrame
                if selected_order_original_index is not None:
                    remaining_orders_df = remaining_orders_df.filter(
                        pl.col("original_idx") != selected_order_original_index
                    )
                else:
                    print("    Warning: selected_order_original_index is None, cannot remove order.")
                    break # Exit loop if order cannot be identified and removed

            elif "No orders left" in status:
                print(f"    No orders left for roll {roll_width}.")
                break # No more orders to process
            else:
                print(f"    Problem is infeasible or status is undefined for roll {roll_width}.")
                print(f"    Status: {status}")
                break # Cannot find solution for remaining orders with this roll

        # Save results for the current roll to a CSV file if cuts occurred
        if current_roll_cuts:
            output_df = pl.DataFrame(current_roll_cuts)
            output_filename = f"roll_cut_results_{roll_width}.csv"
            output_df.write_csv(output_filename)
            print(f"--- Saved {len(current_roll_cuts)} cut items for roll {roll_width} to {output_filename} ---")
        else:
            print(f"--- No cuts made for roll {roll_width} ---")

    # Optional: Save all cutting results from all rolls to a single CSV file
    if all_cut_results:
        final_output_df = pl.DataFrame(all_cut_results)
        final_output_df.write_csv("all_cutting_plan_summary.csv")
        print("\n--- Saved summary of all cutting plans to all_cutting_plan_summary.csv ---")

if __name__ == "__main__":
    asyncio.run(main())
