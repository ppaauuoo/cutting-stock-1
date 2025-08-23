import heapq
from itertools import product
from typing import Optional

import polars as pl

# Material list for logic test
MATERIAL_LIST = [82, 85, 87, 92, 95, 97]
MAX_OUT = 5  # Max 'out' value to try
MIN_COMPAT = 0.1  # Minimum compatibility threshold


INCH_TO_M = 25.4 / 1000  # Conversion factor from inches
CORRUGATE_MULTIPLIERS = {
    "C": 1.45,
    "B": 1.35,
    "E": 1.25,
}


def compute_result(order1: dict[str, int], order2: dict[str, int]) -> float:
    """Compute result for a pair: width1 * out1 + width2 * out2"""
    return order1["width"] * order1.get("out", 1) + order2["width"] * order2.get(
        "out", 1
    )


def passes_logic(order1: dict[str, int], order2: dict[str, int], materials: list[int] = MATERIAL_LIST) -> float:
    """Check if result is within 1 to 5 units of any material value"""
    result = compute_result(order1, order2)
    logic = any(1 <= m - result <= 5 for m in materials)

    EDGE_TYPE = {"X": 1, "N": 2, "W": 2}
    init_type = EDGE_TYPE.get(order1["type"], 0)
    if init_type:
        logic2 = init_type == 1 and order2["type"] not in ["X", "Y"]
        logic2 = init_type == 2 and order2["type"] == "X"

    logic3 = order1["demand"] < order2["demand"]

    logic = logic and logic2 and logic3

    return logic


def compat_score(
    order1: dict[str, int], order2: dict[str, int], materials: list[int] = MATERIAL_LIST
) -> float:
    """Compute compatibility: highest score for valid 'out' assignments"""
    best_score = 0
    best_material = 0
    best_outs = (0, 0)  # Default if none pass
    for out1, out2 in product(range(1, MAX_OUT + 1), repeat=2):
        order1["out"] = out1
        order2["out"] = out2
        if passes_logic(order1, order2, materials):
            result = compute_result(order1, order2)
            # Score inversely proportional to min distance to material
            # Find the closest material that satisfies the logic
            valid_materials = [m for m in materials if 1 <= m - result <= 5]
            min_dist = min(m - result for m in valid_materials)
            score = 1 / (1 + min_dist)  # Higher score for closer match
            if score > best_score:
                best_score = score
                best_outs = (out1, out2)
                best_material = min(valid_materials, key=lambda m: abs(result - m))
    order1["out"], order2["out"] = best_outs  # Assign best 'out' values
    order1["roll"] = best_material
    order2["roll"] = best_material
    return best_score


def greedy_nest(orders: list[dict[str, int|str]], materials: list[int] = None, min_compat: float = MIN_COMPAT):
    """Greedy algorithm to assign 'out' and nest orders"""
    if materials is None:
        materials = MATERIAL_LIST
    # Initialize groups as single orders
    groups: list[list[dict[str, int|str]]] = [
        [{"order_number": o["order_number"], "width": o["width"], "type": o["type"], "demand": o["demand"]}] for o in orders
    ]
    # Priority queue: (-score, i, j) for max-heap
    pairs: list[tuple[float, int, int]] = []
    for i in range(len(orders)):
        for j in range(i + 1, len(orders)):
            score = compat_score(orders[i].copy(), orders[j].copy(), materials=materials)
            if score >= min_compat:
                heapq.heappush(pairs, (-score, i, j))

    # Greedy pairing
    while pairs:
        score, i, j = heapq.heappop(pairs)
        if groups[i] and groups[j]:  # Not already merged
            # Recompute with current group state
            temp_i, temp_j = groups[i][0].copy(), groups[j][0].copy()
            score = compat_score(temp_i, temp_j, materials=materials)
            if score >= min_compat:
                # Assign 'out' values to original orders
                orders[i]["out"] = temp_i["out"]
                orders[j]["out"] = temp_j["out"]
                orders[i]["roll"] = temp_i["roll"]
                orders[j]["roll"] = temp_j["roll"]
                id = f'{orders[i]["order_number"]}-{orders[j]["order_number"]}'
                orders[i]["group_id"] = id
                orders[j]["group_id"] = id
                new_group: list[list[dict[str, int]]] = [groups[i], groups[j]]
                groups.append(new_group)
                groups[i] = groups[j] = None  # Mark as merged
                # Optionally: recompute pairs for new group (for deeper nesting)

    # Filter non-None groups and return
    result = [g for g in groups if g]
    return result, orders


def _flatten_group(group: list) -> list:
    """Recursively flattens a nested list structure representing a group."""
    flat_list = []
    for item in group:
        if isinstance(item, list):
            flat_list.extend(_flatten_group(item))
        else:
            flat_list.append(item)
    return flat_list


def format_greedy_results(
    nested_groups: list,
    updated_orders: list[dict],
    original_orders_df: pl.DataFrame,
    roll_length: int,
    c_type: Optional[str] = None,
    b_type: Optional[str] = None,
) -> list[dict]:
    """
    Formats the results from the greedy_nest function into a structure
    similar to the output of solve_linear_program.
    """
    all_results = []

    updated_orders_map = {o["order_number"]: o for o in updated_orders}
    original_orders_map = {
        o["order_number"]: o for o in original_orders_df.to_dicts()
    }

    for group in nested_groups:
        flat_group_base = _flatten_group(group)
        if not flat_group_base:
            continue

        group_order_numbers = [o["order_number"] for o in flat_group_base]
        group_updated_orders = [
            updated_orders_map.get(on) for on in group_order_numbers
        ]
        group_updated_orders = [o for o in group_updated_orders if o]

        if not group_updated_orders:
            continue

        roll_w = group_updated_orders[0].get("roll")
        if not roll_w or roll_w == 0:
            continue  # Skip groups that were not successfully matched to a roll

        total_cut_width = sum(
            o["width"] * o.get("out", 1) for o in group_updated_orders
        )
        trim = roll_w - total_cut_width

        most_demand_type = None
        if c_type == "C":
            most_demand_type = "C"
        elif b_type == "B":
            most_demand_type = "B"
        elif "E" in (c_type, b_type):
            most_demand_type = "E"
        corr_multiplier = CORRUGATE_MULTIPLIERS.get(most_demand_type, 1.0)

        group_demands = []
        for updated_order in group_updated_orders:
            original_order = original_orders_map.get(updated_order["order_number"])
            if not original_order:
                continue

            cuts = updated_order.get("out", 1)
            total_len_val = (
                original_order.get("length", 0)
                * INCH_TO_M
                * original_order.get("quantity", 0)
                * corr_multiplier
            )
            demand_per_cut = round(total_len_val / cuts, 4) if cuts > 0 else 0
            group_demands.append(demand_per_cut)

        max_demand_per_cut = max(group_demands) if group_demands else 0
        rem_roll_len = round(roll_length - max_demand_per_cut, 4)

        for i, updated_order in enumerate(group_updated_orders):
            original_order = original_orders_map.get(updated_order["order_number"])
            if not original_order:
                continue

            group_id = updated_order.get("group_id")
            cuts = updated_order.get("out", 1)
            demand_per_cut = group_demands[i]

            material_keys = ["demand", "front", "middle", "back", "c", "b", "die_cut"]
            material_specs = {
                key: original_order.get(key)
                for key in material_keys
                if original_order.get(key)
            }
            material_specs.update({"c_type": c_type, "b_type": b_type})

            result = {
                "status": "Optimal",
                "group_id": group_id,
                "objective_value": trim,
                "variables": {
                    "roll_w": roll_w,
                    "rem_roll_l": rem_roll_len,
                    "demand_per_cut": demand_per_cut,
                    "order_w": original_order.get("width"),
                    "order_l": original_order.get("length"),
                    "order_qty": original_order.get("quantity"),
                    "order_dmd": original_order.get("demand"),
                    "cuts": cuts,
                    "trim": trim,
                    "order_idx": original_order.get("original_idx"),
                    "type": original_order.get("type"),
                    "component_type": original_order.get("component_type"),
                    "due_date": original_order.get("due_date"),
                },
                "material_specs": material_specs,
                "message": "Greedy Nesting solution found.",
            }
            all_results.append(result)

    return all_results



def main():
    # Example usage
    orders = [
        {"order_number": "A", "width": 40, "type": "X", "demand": 10},
        {"order_number": "B", "width": 13, "type": "N", "demand": 5},
        {"order_number": "C", "width": 40, "type": "X", "demand": 8},  # Example with same width as A
        {"order_number": "D", "width": 4, "type": "W", "demand": 3},
        {"order_number": "E", "width": 51, "type": "X", "demand": 7},
        {"order_number": "F", "width": 16, "type": "N", "demand": 2},
        {"order_number": "G", "width": 17, "type": "W", "demand": 4},
        {"order_number": "H", "width": 40, "type": "X", "demand": 6},
        {"order_number": "I", "width": 13, "type": "N", "demand": 9},
        {"order_number": "J", "width": 40, "type": "X", "demand": 12},  # Example with same width as A
        {"order_number": "K", "width": 4, "type": "W", "demand": 1},
        {"order_number": "L", "width": 51, "type": "X", "demand": 5},
        {"order_number": "M", "width": 16, "type": "N", "demand": 3},
        {"order_number": "N", "width": 17, "type": "W", "demand": 2},
    ]
    nested, updated_orders = greedy_nest(orders, materials=MATERIAL_LIST)

    print("=== ORDER GROUPING RESULTS ===")
    print(f"Total original orders: {len(orders)}")
    print(f"Total groups formed: {len(nested)}")
    print("\nGROUP DETAILS:")
    for i, group in enumerate(nested, 1):
        if isinstance(group[0], list):  # This is a nested group
            order = next(
                o
                for o in updated_orders
                if o["order_number"] == group[0][0]["order_number"]
            )
            material = order.get("roll", "N/A")
            print(f"\nGroup {i} Roll {material}inch:")
            for j, subgroup in enumerate(group, 1):
                order = next(
                    o
                    for o in updated_orders
                    if o["order_number"] == subgroup[0]["order_number"]
                )
                out_val = order.get("out", 1)
                width = order["width"]
                print(f"  ├─ Order {order['order_number']}: {width}inch (out: {out_val})")
