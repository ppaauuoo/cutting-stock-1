import sys

import polars as pl

# Configure stdout to use UTF-8 encoding
sys.stdout.reconfigure(encoding='utf-8')


def load_data(file_path: str) -> pl.DataFrame:
    """
    Loads data from a CSV file into a Polars DataFrame.

    Args:
        file_path (str): The path to the CSV file.

    Returns:
        pl.DataFrame: The loaded DataFrame.
    """
    try:
        df = pl.read_csv(file_path)
        print(f"Successfully loaded data from {file_path}")
        return df
    except Exception as e:
        print(f"Error loading data from {file_path}: {e}")
        raise

def clean_data(df: pl.DataFrame) -> pl.DataFrame:
    """
    Placeholder function for data cleaning.
    You can add your specific cleaning logic here.

    Args:
        df (pl.DataFrame): The input DataFrame to clean.

    Returns:
        pl.DataFrame: The cleaned DataFrame.
    """
    print("Starting data cleaning...")
    # --- Add your data cleaning steps here ---
    # Example: df = df.drop_nulls()
    # Example: df = df.with_columns(pl.col('column_name').fill_null(0))
    # Example: df = df.filter(pl.col('value') > 10)
    # ---------------------------------------
    df = df.rename({
        "กำหนดส่ง": "due_date",
        "เลขที่ใบสั่งขาย": "order_number",
        "กว้างผลิต": "width",
        "ยาวผลิต": "length",
        "จำนวนสั่งขาย": "demand",
        "จำนวนสั่งผลิต": "quantity",
        "ประเภททับเส้น": "type"
    })
    
    df = df.with_columns(
        pl.col("due_date").str.strptime(pl.Date, "%Y/%m/%d", strict=False),
        pl.col("order_number").cast(pl.Int64),
        pl.col("width").cast(pl.Float64),
        pl.col("length").cast(pl.Float64),
        pl.col("demand").cast(pl.Int64),
        pl.col("quantity").cast(pl.Int64),
        pl.col("type").str.strip_chars()
    )

    df = df.drop_nulls(subset=["due_date", "order_number", "width", "length", "demand", "quantity", "type"])
    df = df.filter(pl.col("demand") > 0)
    df = df.with_columns([
        (pl.col("width") * 24.5).alias("width"),
        (pl.col("length") * 24.5).alias("length")
    ]).select([
        'due_date', 'order_number', 'width', 'length', 'demand', 'quantity', 'type'
    ])

    print("Data cleaning complete.")
    return df

if __name__ == "__main__":
    input_file = "order2024.csv"  # Assuming this file is in the same directory
    
    # Load the data
    raw_df = load_data(input_file)

    if raw_df.height > 0: # Check if DataFrame is not empty
        print("\nRaw Data Head:")
        print(raw_df.head())

        # Clean the data
        cleaned_df = clean_data(raw_df)

        print("\nCleaned Data Head:")
        print(cleaned_df.head())
        
        # Write cleaned data to CSV
        cleaned_df.write_csv("clean_order2024.csv", include_header=True)


