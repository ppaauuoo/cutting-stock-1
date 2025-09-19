import os
import pickle
from typing import Tuple, Optional, Callable

import polars as pl
import xgboost as xgb
from cuttingstock.utils import log_message

import sys

STATUS_OPTIMAL = "Optimal"
INCH_TO_M = 25.4 / 1000  # Conversion factor from inches
MIN_TRIM_WASTE = 1
MAX_TRIM_WASTE = 5
CORRUGATE_MULTIPLIERS = {
    "C": 1.45,
    "B": 1.35,
    "E": 1.25,
}

def resource_path(relative_path):
    """Get the absolute path to a resource, works for development and PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller creates a temp folder and stores files there
        base_path = sys._MEIPASS
    else:
        # Use the current working directory during development
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def process(features: pl.DataFrame) -> pl.DataFrame:
    # Use polars selectors to identify column types.
    # Note: pandas 'object' dtype is usually a string in polars.
    numerical_cols = features.select(pl.selectors.numeric()).columns
    categorical_cols = features.select(
        pl.selectors.string() | pl.selectors.categorical()
    ).columns
    boolean_cols = features.select(pl.selectors.boolean()).columns

    expressions = []
    # Process categorical features
    for col in categorical_cols:
        # Fill nulls with a placeholder value (e.g., 'missing')
        # Convert to categorical and encode to integer representation
        expressions.append(
            pl.col(col).fill_null("None").cast(pl.Categorical).to_physical().alias(col)
        )

    # Process numerical features
    for col in numerical_cols:
        expressions.append(pl.col(col).fill_null(0).alias(col))

    # Process boolean features
    for col in boolean_cols:
        # Convert bool to int8 for CuPy compatibility
        expressions.append(
            pl.col(col).cast(pl.Int8).fill_null(0).alias(col)
        )  # Fill nulls with False (or True, depending on your data)

    if not expressions:
        return features

    return features.with_columns(expressions)


_models_cache = {}


def load_models() -> dict:
    """Loads XGBoost models and label mappings from disk and caches them."""
    if _models_cache:
        return _models_cache

    label_out_path = resource_path("model/label_mapping_out.pkl")
    label_roll_width_path = resource_path("model/label_mapping_roll_width.pkl")
    out_model_path = resource_path("model/out.ubj")
    roll_width_model_path = resource_path("model/roll_width.ubj")

    # Debug information
    log_message("info", "Loading models from paths", {
        "label_out_path": label_out_path,
        "label_roll_width_path": label_roll_width_path,
        "out_model_path": out_model_path,
        "roll_width_model_path": roll_width_model_path
    })

    # Check if files exist before loading
    for path, name in [(label_out_path, "label_mapping_out.pkl"),
                       (label_roll_width_path, "label_mapping_roll_width.pkl"),
                       (out_model_path, "out.ubj"),
                       (roll_width_model_path, "roll_width.ubj")]:
        if not os.path.exists(path):
            log_message("error", f"Model file not found", {"file": name, "path": path})
            raise FileNotFoundError(f"Model file not found: {path}")

    try:
        with open(label_out_path, "rb") as f:
            label_mapping_out = pickle.load(f)
        _models_cache["reverse_label_mapping_out"] = {
            idx: val for val, idx in label_mapping_out.items()
        }

        with open(label_roll_width_path, "rb") as f:
            label_mapping_roll_width = pickle.load(f)
        _models_cache["reverse_label_mapping_roll_width"] = {
            idx: val for val, idx in label_mapping_roll_width.items()
        }

        out_model = xgb.XGBClassifier()
        out_model.load_model(out_model_path)
        _models_cache["out_model"] = out_model

        roll_width_model = xgb.XGBClassifier()
        roll_width_model.load_model(roll_width_model_path)
        _models_cache["roll_width_model"] = roll_width_model

        log_message("info", "All models loaded successfully")
        return _models_cache
    except Exception as e:
        log_message("error", "Failed to load models", {"error": str(e)})
        # Clear cache on failure to prevent partial loading
        _models_cache.clear()
        raise


async def try_xgboost_solution(
    orders_to_process: pl.DataFrame, roll: dict, c_type: Optional[str], b_type: Optional[str],
    progress_callback: Optional[Callable[[str], None]]
) -> Optional[dict]:
    """Tries to find a quick solution using the pre-trained XGBoost model."""
    try:
        if progress_callback:
            progress_callback("    🤖 Trying XGBoost for a quick solution...")

        xgb_cuts_preds, _ = _predict_with_xgboost(orders_to_process)
        candidate_orders = orders_to_process.with_columns(
            pl.Series("xgb_cuts", xgb_cuts_preds, dtype=pl.Int64),
            # pl.Series("xgb_roll_w", xgb_roll_w_preds, dtype=pl.Int64),
        )
        # candidate_orders = orders_with_preds.filter(pl.col("xgb_roll_w") == roll['width'])

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
    except (KeyError, FileNotFoundError, Exception) as e:
        log_message("error", "XGBoost prediction failed.", {"error": str(e), "type": type(e).__name__})
        if progress_callback:
            progress_callback(f"    ⚠️ XGBoost prediction failed: {e}. Falling back to linear solver.")
    return None

def _predict_with_xgboost(orders_df: pl.DataFrame) -> Tuple[list, list]:
    """
    Takes an order DataFrame, preprocesses it, and returns predictions from cached models.
    """
    try:
        models = load_models()
    except Exception as e:
        log_message("error", "Failed to load models for prediction", {"error": str(e)})
        raise

    # Check if required models are loaded
    required_models = ["out_model", "roll_width_model", "reverse_label_mapping_out", "reverse_label_mapping_roll_width"]
    missing_models = [model for model in required_models if model not in models]
    if missing_models:
        error_msg = f"Missing required models: {missing_models}"
        log_message("error", error_msg)
        raise KeyError(error_msg)

    out_model = models["out_model"]
    roll_width_model = models["roll_width_model"]
    reverse_label_mapping_out = models["reverse_label_mapping_out"]
    reverse_label_mapping_roll_width = models["reverse_label_mapping_roll_width"]

    feature_cols = [
        "component_type",
        "width",
        "length",
        "quantity",
        "front",
        "c",
        "middle",
        "b",
        "back",
        "type",
    ]
    # Ensure all required columns exist, filling with null if not.
    for col in feature_cols:
        if col not in orders_df.columns:
            orders_df = orders_df.with_columns(pl.lit(None).alias(col))

    try:
        X = process(orders_df.select(feature_cols))
        out_predictions = out_model.predict(X)
        roll_width_predictions = roll_width_model.predict(X)
    except Exception as e:
        log_message("error", "Failed to make predictions", {"error": str(e)})
        raise

    out_predictions_original = [
        reverse_label_mapping_out.get(pred) for pred in out_predictions
    ]
    roll_width_predictions_original = [
        reverse_label_mapping_roll_width.get(pred)
        for pred in roll_width_predictions
    ]

    return out_predictions_original, roll_width_predictions_original
