import sys

import pandas as pd

# Configure stdout to use UTF-8 encoding
sys.stdout.reconfigure(encoding='utf-8')


def load_data(file_path: str) -> pd.DataFrame:
    """
    Loads data from an Excel file into a Pandas DataFrame.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        pd.DataFrame: The loaded DataFrame.
    """
    try:
        df = pd.read_csv(file_path)
        print(f"Successfully loaded data from {file_path}")
        return df
    except Exception as e:
        print(f"Error loading data from {file_path}: {e}")
        raise

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Placeholder function for data cleaning.
    You can add your specific cleaning logic here.

    Args:
        df (pd.DataFrame): The input DataFrame to clean.

    Returns:
        pd.DataFrame: The cleaned DataFrame.
    """
    print("Starting data cleaning...")
    # --- Add your data cleaning steps here ---
    # Example: df = df.dropna()
    # Example: df['column_name'] = df['column_name'].fillna(0)
    # Example: df = df[df['value'] > 10]
    # ---------------------------------------
    df = df.rename(columns={"กำหนดส่ง": "due_date"})
    df = df.rename(columns={"เลขที่ใบสั่งขาย": "order_number"})
    df = df.rename(columns={"กว้างผลิต": "width"})
    df = df.rename(columns={"ยาวผลิต": "length"})
    df = df.rename(columns={"จำนวนสั่งขาย": "demand"})
    df = df.rename(columns={"จำนวนสั่งผลิต": "quantity"})
    df = df.rename(columns={"ประเภททับเส้น": "type"})


    print("Data cleaning complete.")
    return df

if __name__ == "__main__":
    input_file = "order2024.csv"  # Assuming this file is in the same directory
    
    # Load the data
    raw_df = load_data(input_file)

    if not raw_df.empty:
        print("\nRaw Data Head:")
        print(raw_df.head())  # ใช้ print() โดยตรง ซึ่งจะแสดง DataFrame หัวตารางได้

        # Clean the data
        cleaned_df = clean_data(raw_df)

        print("\nCleaned Data Head:")
        print(cleaned_df.head())
        
        # Write cleaned data to CSV
        cleaned_df.to_csv("clean_order2024.csv", index=False)


