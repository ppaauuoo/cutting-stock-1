import heapq
from itertools import product

# Material list for logic test
MATERIAL_LIST = [82, 85, 87, 92, 95, 97]
MAX_OUT = 5  # Max 'out' value to try
MIN_COMPAT = 0.1  # Minimum compatibility threshold


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


def greedy_nest(orders: list[dict[str, int|str]], min_compat: float = MIN_COMPAT):
    """Greedy algorithm to assign 'out' and nest orders"""
    # Initialize groups as single orders
    groups: list[list[dict[str, int|str]]] = [
        [{"order_number": o["order_number"], "width": o["width"], "type": o["type"], "demand": o["demand"]}] for o in orders
    ]
    # Priority queue: (-score, i, j) for max-heap
    pairs: list[tuple[float, int, int]] = []
    for i in range(len(orders)):
        for j in range(i + 1, len(orders)):
            score = compat_score(orders[i].copy(), orders[j].copy())
            if score >= min_compat:
                heapq.heappush(pairs, (-score, i, j))

    # Greedy pairing
    while pairs:
        score, i, j = heapq.heappop(pairs)
        if groups[i] and groups[j]:  # Not already merged
            # Recompute with current group state
            temp_i, temp_j = groups[i][0].copy(), groups[j][0].copy()
            score = compat_score(temp_i, temp_j)
            if score >= min_compat:
                # Assign 'out' values to original orders
                orders[i]["out"] = temp_i["out"]
                orders[j]["out"] = temp_j["out"]
                orders[i]["roll"] = temp_i["roll"]
                orders[j]["roll"] = temp_j["roll"]
                # Form new nested group
                new_group: list[list[dict[str, int]]] = [groups[i], groups[j]]
                groups.append(new_group)
                groups[i] = groups[j] = None  # Mark as merged
                # Optionally: recompute pairs for new group (for deeper nesting)

    # Filter non-None groups and return
    result = [g for g in groups if g]
    return result, orders


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
nested, updated_orders = greedy_nest(orders)

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
