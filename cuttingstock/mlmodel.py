import os
import pickle
from pathlib import Path
from typing import Tuple

import polars as pl
import xgboost as xgb

import sys

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

    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent
    model_dir = project_root / "model"
    label_out_path = model_dir / "label_mapping_out.pkl"
    label_roll_width_path = model_dir / "label_mapping_roll_width.pkl"
    out_model_path = model_dir / "out.ubj"
    roll_width_model_path = model_dir / "roll_width.ubj"

    with open(label_out_path, "rb") as f:
        label_mapping_out = pickle.load(resource_path(f))
    _models_cache["reverse_label_mapping_out"] = {
        idx: val for val, idx in label_mapping_out.items()
    }

    with open(label_roll_width_path, "rb") as f:
        label_mapping_roll_width = pickle.load(resource_path(f))
    _models_cache["reverse_label_mapping_roll_width"] = {
        idx: val for val, idx in label_mapping_roll_width.items()
    }

    with open(out_model_path, "rb") as f:
        out_model_bytes = bytearray(f.read())
    out_model = xgb.XGBClassifier()
    out_model.load_model(resource_path(out_model_bytes))
    _models_cache["out_model"] = out_model

    with open(roll_width_model_path, "rb") as f:
        roll_width_model_bytes = bytearray(f.read())
    roll_width_model = xgb.XGBClassifier()
    roll_width_model.load_model(resource_path(roll_width_model_bytes))
    _models_cache["roll_width_model"] = roll_width_model

    return _models_cache


def predict_with_xgboost(orders_df: pl.DataFrame) -> Tuple[list, list]:
    """
    Takes an order DataFrame, preprocesses it, and returns predictions from cached models.
    """
    models = load_models()
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

    X = process(orders_df.select(feature_cols))

    out_predictions = out_model.predict(X)
    roll_width_predictions = roll_width_model.predict(X)

    out_predictions_original = [
        reverse_label_mapping_out.get(pred) for pred in out_predictions
    ]
    roll_width_predictions_original = [
        reverse_label_mapping_roll_width.get(pred)
        for pred in roll_width_predictions
    ]

    return out_predictions_original, roll_width_predictions_original


def main():
    from cuttingstock.cleaning import clean_data, load_data

    # Example usage:
    # These would be your inputs
    try:
        # Using a raw string for the path is safer on Windows
        raw_orders_df = load_data(r"D:\order.csv")
    except FileNotFoundError:
        print("Error: The file D:\\order.csv was not found.")
        return
    except Exception as e:
        print(f"An error occurred while loading the data: {e}")
        return

    start_date = None
    end_date = None
    front = 'KS231'
    c = 'CM127'
    middle = 'CM127'
    b = 'CM127'
    back = 'KB160'
    c_type = "C"  # example value
    b_type = "B"  # example value

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

    (
        out_predictions_original,
        roll_width_predictions_original,
    ) = predict_with_xgboost(orders_df)

    # Print or use predictions
    print("Order Width:", orders_df["width"].to_list()[0])
    print("Out Model Predictions (Original Labels):", out_predictions_original[0])
    print(
        "Roll Width Model Predictions (Original Labels):",
        roll_width_predictions_original[0],
    )


if __name__ == "__main__":
    main()
