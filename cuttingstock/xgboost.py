import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from xgboost import XGBClassifier
import polars as pl
import pickle


def process(features: pl.DataFrame) -> pl.DataFrame:
    numerical_features = features.select_dtypes(
        include=["int32", "int64", "float32", "float64"]
    ).columns.tolist()
    categorical_features = features.select_dtypes(
        include=["object", "category"]
    ).columns.tolist()
    boolean_features = features.select_dtypes(
        include=["bool"]
    ).columns.tolist()  # CryoSleep, VIP

    for col in categorical_features:
        # Fill nulls with a placeholder value (e.g., 'missing')
        features[col] = features[col].fillna("None")
        # Convert to categorical and encode
        features[col] = features[col].astype("category").cat.codes

    for col in numerical_features:
        features[col] = features[col].fillna(0)

    for col in boolean_features:
        features[col] = features[col].astype(
            "int8"
        )  # Convert bool to int8 for CuPy compatibility
        features[col] = features[col].fillna(
            0
        )  # Fill nulls with False (or True, depending on your data)

    return features


def main():
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
