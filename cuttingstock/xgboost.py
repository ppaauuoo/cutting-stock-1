import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from xgboost import XGBClassifier
import polars as pl
import pickle


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
            pl.col(col)
            .cast(pl.Int8)
            .fill_null(0)
            .alias(col)
        )  # Fill nulls with False (or True, depending on your data)

    if not expressions:
        return features

    return features.with_columns(expressions)


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
    front = None
    c = None
    middle = None
    b = None
    back = None
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
    X = process(
        orders_df.select(
            [
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
        )
    )

    with open("./model/label_mapping_out.pkl", "rb") as f:
        label_mapping = pickle.load(f)
    reverse_label_mapping_out = {idx: val for val, idx in label_mapping.items()}

    with open("./model/label_mapping_roll_width.pkl", "rb") as f:
        label_mapping = pickle.load(f)
    reverse_label_mapping_roll_width = {idx: val for val, idx in label_mapping.items()}

    # Load models
    out_model = XGBClassifier()
    out_model.n_classes_ = 1
    out_model.load_model("./model/out.ubj")

    roll_width_model = XGBClassifier()
    roll_width_model.n_classes_ = 1
    roll_width_model.load_model("./model/roll_width.ubj")

    # Get predictions
    out_predictions = out_model.predict(X)
    roll_width_predictions = roll_width_model.predict(X)

    # Convert predictions to original labels
    out_predictions_original = [
        reverse_label_mapping_out[pred] for pred in out_predictions
    ]
    roll_width_predictions_original = [
        reverse_label_mapping_roll_width[pred] for pred in roll_width_predictions
    ]

    # Print or use predictions
    print("Out Model Predictions (Original Labels):", out_predictions_original)
    print(
        "Roll Width Model Predictions (Original Labels):",
        roll_width_predictions_original,
    )


if __name__ == "__main__":
    main()
