import os
import re
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Set, Tuple

import polars as pl

from cuttingstock.cleaning import clean_data, load_data
from cuttingstock.grouping import format_greedy_results, greedy_nest
from cuttingstock.linear import solve_linear_program
from cuttingstock.material import (
    OutOfStockError,
    handle_unprocessed_orders,
    process_single_order,
)
from cuttingstock.mlmodel import try_xgboost_solution
from cuttingstock.order import filter_orders_by_factory
from cuttingstock.utils import log_message

# Status Messages
STATUS_OPTIMAL = "Optimal"
STATUS_INFEASIBLE = "Infeasible"
STATUS_FAILED = "Failed"

CORRUGATE_MULTIPLIERS = {
    "C": 1.45,
    "B": 1.35,
    "E": 1.25,
}

# Factory configuration constants
FACTORY_CONFIGS = {
    "1": {"min_width": 73, "max_width": 79},
    "2": {"min_width": 82, "max_width": 97},
}

# Material columns configuration
MATERIAL_COLUMNS = ['front', 'c', 'middle', 'b', 'back']

@dataclass
class WidthInfo:
    """Container for width information and numeric value."""
    width_str: str
    width_int: int
    
    @classmethod
    def from_string(cls, width_str: str) -> 'WidthInfo':
        """Create WidthInfo from string representation."""
        try:
            width_int = int(re.search(r'\d+', width_str).group() if re.search(r'\d+', width_str) else 0)
            return cls(width_str, width_int)
        except (AttributeError, ValueError):
            return cls(width_str, 0)

@dataclass
class AvailableWidth:
    """Container for available width data."""
    width: str
    length: float

def extract_width_from_string(width_str: str) -> int:
    """
    Extract numeric width from string representation.
    
    Args:
        width_str: String containing width (e.g., '79B', '82')
        
    Returns:
        int: Extracted numeric width, 0 if extraction fails
    """
    try:
        return int(re.search(r'\d+', width_str).group())
    except (AttributeError, ValueError):
        return 0

def is_width_valid_for_factory(width_str: str, factory: str) -> bool:
    """
    Check if width is valid for the specified factory.
    
    Args:
        width_str: Width string to validate
        factory: Factory identifier ('1', '2', etc.)
        
    Returns:
        bool: True if width is valid for factory
    """
    if factory not in FACTORY_CONFIGS:
        return True  # Allow all widths for unknown factories
    
    width_int = extract_width_from_string(width_str)
    config = FACTORY_CONFIGS[factory]
    
    return config["min_width"] <= width_int <= config["max_width"]

def create_factory_sort_key(factory: str) -> Callable[[str], Tuple[int, int]]:
    """
    Create a sort key function for factory-specific width sorting.
    
    Args:
        factory: Factory identifier
        
    Returns:
        callable: Sort key function for the specified factory
    """
    if factory not in FACTORY_CONFIGS:
        return lambda width_str: (0, extract_width_from_string(width_str))
    
    config = FACTORY_CONFIGS[factory]
    min_width, max_width = config["min_width"], config["max_width"]
    
    def sort_key(width_str: str) -> Tuple[int, int]:
        width_int = extract_width_from_string(width_str)
        if min_width <= width_int <= max_width:
            return (0, width_int)  # Prioritize valid widths
        return (1, width_int)      # Deprioritize invalid widths
    
    return sort_key

def calculate_total_relevant_length(spec_materials: Set[str], materials_in_stock: Dict[str, Dict]) -> float:
    """
    Calculate total relevant length for specified materials.
    
    Args:
        spec_materials: Set of materials to calculate for
        materials_in_stock: Dictionary of materials in stock
        
    Returns:
        float: Total length of relevant materials
    """
    try:
        return sum(
            roll.get('length', 0)
            for mat in spec_materials
            for roll in materials_in_stock.get(mat, {}).values()
        )
    except (AttributeError, TypeError):
        log_message("warning", "Error calculating total relevant length", {
            'spec_materials': spec_materials,
            'error': 'Invalid data structure'
        })
        return 0.0

def get_available_widths_for_spec(
    spec_materials: Set[str],
    roll_specs: Dict[str, Dict],
    selected_factory: str
) -> List[AvailableWidth]:
    """
    Get list of available widths that can accommodate the specified materials.
    
    Args:
        spec_materials: Set of materials needed
        roll_specs: Dictionary of roll specifications
        selected_factory: Selected factory identifier
        
    Returns:
        List[AvailableWidth]: List of available widths with their lengths
    """
    if not roll_specs:
        return []
    
    available_widths = []
    
    for width_str, materials_in_stock in roll_specs.items():
        # Skip widths not valid for the factory
        if not is_width_valid_for_factory(width_str, selected_factory):
            continue
        
        # Check if all required materials are available
        if spec_materials.issubset(materials_in_stock.keys()):
            total_length = calculate_total_relevant_length(spec_materials, materials_in_stock)
            if total_length > 0:
                available_widths.append(AvailableWidth(width_str, total_length))
    
    return available_widths

def sort_widths_by_priority(
    available_widths: List[AvailableWidth],
    selected_factory: str
) -> List[str]:
    """
    Sort available widths by factory-specific priority.
    
    Args:
        available_widths: List of available widths
        selected_factory: Selected factory identifier
        
    Returns:
        List[str]: Sorted list of width strings
    """
    if not available_widths:
        return []
    
    # Sort by length (descending) and factory-specific width priority
    sort_key_func = create_factory_sort_key(selected_factory)
    
    sorted_widths = sorted(
        available_widths,
        key=lambda w: (-w.length, sort_key_func(w.width))
    )
    
    return [w.width for w in sorted_widths]

def create_suggestion_from_spec(spec_row: Dict, width: str) -> Dict:
    """
    Create a suggestion dictionary from specification row and width.
    
    Args:
        spec_row: Specification row from grouped data
        width: Selected width string
        
    Returns:
        Dict: Formatted suggestion
    """
    full_spec = {k: v for k, v in spec_row.items() if k != 'len'}
    return {'width': width, 'spec': full_spec}

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

def generate_suggestions(
    orders_df: pl.DataFrame,
    roll_specs: Dict[str, Dict],
    selected_factory: str,
    test: bool = False
) -> List[Dict]:
    """
    Generates a list of all possible calculation settings based on order frequency and stock.
    
    Args:
        orders_df: DataFrame containing order information
        roll_specs: Dictionary of roll specifications with material availability
        selected_factory: Selected factory identifier
        test: Whether to enable test mode sorting
        
    Returns:
        List[Dict]: List of suggestion dictionaries with width and spec information
    """
    # Input validation
    if orders_df is None or orders_df.is_empty():
        log_message("info", "No orders provided for suggestion generation")
        return []

    # Get available material columns
    existing_cols = [col for col in MATERIAL_COLUMNS if col in orders_df.columns]
    if not existing_cols:
        log_message("warning", "No material columns found in orders data")
        return []

    try:
        # Group by material specifications and count frequencies
        all_specs_df = orders_df.group_by(existing_cols).len().sort("len", descending=True)
        
        if all_specs_df.is_empty():
            log_message("info", "No valid specifications found after grouping")
            return []

        suggestions = []
        
        for spec_row in all_specs_df.iter_rows(named=True):
            # Extract materials from specification (non-empty values)
            spec_materials = {m for k, m in spec_row.items() if k != 'len' and m}
            
            if not spec_materials:
                continue

            # Get available widths for this specification
            available_widths = get_available_widths_for_spec(
                spec_materials, roll_specs, selected_factory
            )
            
            if not available_widths:
                continue

            # Sort widths by priority
            sorted_widths = sort_widths_by_priority(available_widths, selected_factory)
            
            # Create suggestions for each valid width
            for width in sorted_widths:
                suggestion = create_suggestion_from_spec(spec_row, width)
                suggestions.append(suggestion)

        # Sort suggestions by complexity and material values
        suggestions.sort(key=lambda s: (
            len([v for v in s['spec'].values() if v]),  # Number of materials
            sorted([v for v in s['spec'].values() if v])  # Sorted material names
        ))

        # Test mode: sort by width numeric value
        if test:
            suggestions.sort(key=lambda w: extract_width_from_string(str(w['width'])))

        log_message("info", "Suggestions generated successfully", {
            'count': len(suggestions),
            'factory': selected_factory
        })
        
        return suggestions

    except Exception as e:
        log_message("error", "Error generating suggestions", {
            'error': str(e),
            'factory': selected_factory
        })
        return []

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
                log_message("info", "Greedy nesting results", {'results': greedy_results})
            greedy_results_to_return = greedy_results

    # orders_for_solvers = orders_to_process
    orders_for_solvers = pl.DataFrame(updated_orders)
    if greedy_results_to_return:
        orders_for_solvers = orders_for_solvers.filter(pl.col('quantity') > 0)
        # processed_indices = {
        #     res.get("variables", {}).get("order_idx") for res in greedy_results_to_return
        # }
        # processed_indices.discard(None)  # Remove None if it exists
        # if processed_indices:
        #     orders_for_solvers = orders_to_process.filter(
        #         ~pl.col("original_idx").is_in(list(processed_indices))
        #     )

    log_message("info", "Orders after greedy", {'remaining_orders_count': orders_for_solvers.shape[0]})
    if not greedy_results_to_return and progress_callback:
        progress_callback("    Greedy nesting did not find a solution. Falling back to XGBoost.")

    # 2. Attempt to find solutions for remaining orders with XGBoost and Linear Programming.
    solver_solutions = []
    solution = None # Represents the last or most relevant solution for status checks

    if not orders_for_solvers.is_empty():
        xgb_solution = await try_xgboost_solution(orders_for_solvers, roll, c_type, b_type, progress_callback)

        if xgb_solution:
            solver_solutions.append(xgb_solution)
            solution = xgb_solution
            order_idx = xgb_solution.get("variables", {}).get("order_idx")
            if order_idx is not None:
                orders_for_solvers = orders_for_solvers.filter(pl.col("original_idx") != order_idx)

        # Process remaining orders with linear solver
        if not orders_for_solvers.is_empty():
            if not xgb_solution and progress_callback:
                progress_callback("    XGBoost did not find a solution. Falling back to linear solver.")
                log_message("info", "XGBoost did not find a solution")

            linear_solution = await solve_linear_program(
                roll['width'], roll['length'], orders_for_solvers, c_type=c_type, b_type=b_type
            )

            if linear_solution:
                if linear_solution.get("status") == STATUS_OPTIMAL:
                    solver_solutions.append(linear_solution)
                # Use linear solution for final status, as it ran on the final set of orders
                solution = linear_solution
    else:
        # This case handles when greedy nesting processes all orders.
        solution = {"status": "NoOrdersLeft", "message": "No orders left for solvers."}
        log_message("info", "No orders left for solvers")

    # Combine greedy and solver results if applicable
    if greedy_results_to_return:
        if solver_solutions:
            combined_results = greedy_results_to_return + solver_solutions
            log_message("info", "Combined results", {'results': combined_results})
            return combined_results, solution
        else:
            log_message("info", "Greedy nesting successful, no further solver solutions.")
            return greedy_results_to_return, {"status": "GreedyNestingSuccess", "message": "Greedy nesting found a solution."}
    else:
        if solver_solutions:
            log_message("info", "Solver(s) found a solution.")
            return solver_solutions, solution
        else:
            log_message("error", "No solver found a solution.")
            return [], solution


def _load_and_prepare_data(
    file_path: str,
    start_date: Optional[str],
    end_date: Optional[str],
    front: Optional[str],
    c_type: Optional[str],
    c: Optional[str],
    middle: Optional[str],
    b_type: Optional[str],
    b: Optional[str],
    back: Optional[str],
    processed_orders: Optional[set],
    max_records: Optional[int],
    output_dir: str = "cache",
    selected_factory: str = "ALL"
) -> pl.DataFrame:
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
    orders_df = filter_orders_by_factory(orders_df, selected_factory)
    log_message("info", "Loaded and filtered data successfully")

    if processed_orders:
        orders_df = orders_df.filter(
            ~pl.col("order_number").is_in(list(processed_orders))
        )

    log_message("info", "Clear data successfully")

    if max_records:
        orders_df = orders_df.head(max_records)
    orders_df = orders_df.with_row_index("original_idx")
    return orders_df


async def _process_cuts_for_roll(
    roll: dict,
    rem_orders_df: pl.DataFrame,
    orders_df: pl.DataFrame,
    c_type: Optional[str],
    b_type: Optional[str],
    c: Optional[str],
    b: Optional[str],
    order_num_col_idx: int,
    material_substitutions: dict,
    progress_callback: Optional[Callable[[str], None]],
    out_of_stock_handler: Optional[Callable[[OutOfStockError], Optional[dict]]],
    roll_specs: Optional[dict],
) -> tuple[list, pl.DataFrame, str, Optional[str]]:
    last_used_roll_ids = {}
    used_roll_ids_for_cut = set()
    roll_positions = {}
    if progress_callback:
        progress_callback(f"🔧 กำลังประมวลผลม้วน {roll['width']} นิ้ว")

    roll_cuts = []
    iteration = 0
    while not rem_orders_df.is_empty():
        iteration += 1
        if progress_callback:
            progress_callback(f"  Iteration {iteration}: Remaining orders: {rem_orders_df.shape[0]} items")

        orders_to_process = rem_orders_df

        if c is None: c_type = None
        if b is None: b_type = None

        results_to_process, result = await _find_solution(
            orders_to_process, roll, c_type, b_type, progress_callback, orders_df
        )

        if not results_to_process:
            if progress_callback:
                progress_callback(f"    ❌ {result.get('message', 'Non-optimal status')}")
            break

        processed_order_indices_this_iteration = set()
        cut_infos_this_iteration = []
        last_rem_roll_l = roll['length']
        all_successful = True

        for res in results_to_process:
            cut_info, order_idx = await process_single_order(
                res, orders_df, order_num_col_idx, material_substitutions,
                progress_callback, out_of_stock_handler, roll_specs,
                used_roll_ids_for_cut, last_used_roll_ids, roll_positions, roll['width']
            )

            if cut_info:
                cut_infos_this_iteration.append(cut_info)
                if order_idx is not None:
                    processed_order_indices_this_iteration.add(order_idx)
                last_rem_roll_l = cut_info.get("rem_roll_l", last_rem_roll_l)
            else:
                all_successful = False

        roll_cuts.extend(cut_infos_this_iteration)
        roll['length'] = last_rem_roll_l

        if processed_order_indices_this_iteration:
            rem_orders_df = rem_orders_df.filter(~pl.col("original_idx").is_in(list(processed_order_indices_this_iteration)))

        if not all_successful:
            # failure_reason = failure_reason or "One or more cuts in the group failed."
            break

        # If we used greedy nesting, it creates a full plan, so we can break from the main loop for this roll.
        if result.get("status") != STATUS_OPTIMAL and results_to_process:
             break

    return roll_cuts, rem_orders_df

def _save_roll_results_to_db(
    roll_cuts: list,
    roll_width: int,
    progress_callback: Optional[Callable[[str], None]],
    output_dir: str = "cache",
):
    if roll_cuts:
        output_df = pl.DataFrame(roll_cuts)
        db_path = os.path.join(output_dir, "cache.db")
        conn_str = f"sqlite:///{os.path.abspath(db_path)}"
        table_name = f"roll_cut_results_{roll_width}"
        output_df.write_database(table_name, connection=conn_str, if_table_exists="replace")
        if progress_callback:
            progress_callback(f"--- Saved {len(roll_cuts)} cuts for roll {roll_width} to database table '{table_name}' ---")
    elif progress_callback:
        progress_callback(f"--- No cuts made for roll {roll_width} ---")


def _save_summary_results_to_db(
    all_results: list,
    progress_callback: Optional[Callable[[str], None]],
    output_dir: str = "cache",
):
    if all_results:
        final_output_df = pl.DataFrame(all_results)
        db_path = os.path.join(output_dir, "cache.db")
        conn_str = f"sqlite:///{os.path.abspath(db_path)}"
        final_output_df.write_database("all_cutting_plan_summary", connection=conn_str, if_table_exists="replace")
        if progress_callback:
            progress_callback("💾 บันทึกผลลัพธ์ลงฐานข้อมูลเรียบร้อย")


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
    selected_factory: str = "ALL"
):
    output_dir = "cache"
    os.makedirs(output_dir, exist_ok=True)
    log_message("info", "Starting calculation process")

    orders_df = _load_and_prepare_data(
        file_path, start_date, end_date, front, c_type, c, middle, b_type, b,
        back, processed_orders, max_records, output_dir, selected_factory
    )

    # rolls per specs (set of materials)
    rolls = [{"width": roll_width, "length": roll_length}]
    all_results = []

    order_num_col_idx = orders_df.columns.index("order_number")

    if material_substitutions is None:
        material_substitutions = {}

    rem_orders_df = orders_df.clone()
    for roll in rolls:
        roll_cuts, rem_orders_df = await _process_cuts_for_roll(
            roll, rem_orders_df, orders_df, c_type, b_type, c, b, order_num_col_idx,
            material_substitutions, progress_callback, out_of_stock_handler, roll_specs
        )
        all_results.extend(roll_cuts)

        _save_roll_results_to_db(roll_cuts, roll['width'], progress_callback, output_dir)


    _save_summary_results_to_db(all_results, progress_callback, output_dir)

    return all_results
