import sys
import unicodedata  # เพิ่มส่วน import
from datetime import date, datetime
from typing import Optional

import polars as pl


def load_data(file_path: str) -> pl.DataFrame:
    """
    Loads data from a CSV file into a Polars DataFrame.

    Args:
        file_path (str): The path to the CSV file.

    Returns:
        pl.DataFrame: The loaded DataFrame.
    """
    try:
        # เพิ่ม encoding='utf8' และพารามิเตอร์เกี่ยวกับ null values
        df = pl.read_csv(
            file_path,
            encoding='utf8',  # ระบุ encoding ตรงนี้
            null_values=["", " ", "NULL", "N/A"]
        )
        print(f"Successfully loaded data from {file_path}")
        return df
    except Exception as e:
        print(f"Error loading data from {file_path}: {e}")
        raise

def clean_data(df: pl.DataFrame, start_date: Optional[str] = None, end_date: Optional[str] = None) -> pl.DataFrame:
    """
    Placeholder function for data cleaning.
    You can add your specific cleaning logic here.

    Args:
        df (pl.DataFrame): The input DataFrame to clean.
        start_date (Optional[str]): Start date for filtering (YYYY-MM-DD).
        end_date (Optional[str]): End date for filtering (YYYY-MM-DD).

    Returns:
        pl.DataFrame: The cleaned DataFrame.
    """
    print("Starting data cleaning...")
    
    # แปลงชื่อคอลัมน์ให้เป็นรูปแบบ Unicode มาตรฐานและตัดช่องว่าง
    def normalize_col_name(col: str) -> str:
        return unicodedata.normalize('NFC', col.strip())
    
    # สร้าง mapping สำหรับคอลัมน์ไทยพร้อมชื่อที่แปลงแล้ว
    thai_col_mapping = {
        normalize_col_name("กำหนดส่ง"): "due_date",
        normalize_col_name("เลขที่ใบสั่งขาย"): "order_number",
        normalize_col_name("กว้างผลิต"): "width",
        normalize_col_name("ยาวผลิต"): "length",
        normalize_col_name("จำนวนสั่งขาย"): "demand",
        normalize_col_name("จำนวนสั่งผลิต"): "quantity",
        normalize_col_name("ประเภททับเส้น"): "type",
        normalize_col_name("แผ่นหน้า"): "front",
        normalize_col_name("ลอน C"): "C",
        normalize_col_name("แผ่นกลาง"): "middle",
        normalize_col_name("ลอน B"): "B",
        normalize_col_name("แผ่นหลัง"): "back"
    }
    
    # สร้าง dictionary สำหรับเปลี่ยนชื่อคอลัมน์
    rename_dict = {}
    for orig_col in df.columns:
        normalized = normalize_col_name(orig_col)
        if normalized in thai_col_mapping:
            rename_dict[orig_col] = thai_col_mapping[normalized]
    
    # เปลี่ยนชื่อคอลัมน์
    df = df.rename(rename_dict)
    
    # ตรวจสอบว่ามีคอลัมน์จำเป็นครบ
    required_cols = ["due_date", "order_number", "width", "length", "demand", "quantity", "type", "front", "C", "middle", "B", "back"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"⚠️ คอลัมน์หาย: {missing} โปรดตรวจสอบชื่อคอลัมน์ในไฟล์ CSV")
    
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
    df = df.filter(pl.col("width") > 0)
    df = df.filter(pl.col("length") > 0)
    df = df.with_columns([
        (pl.col("width") / 24.5).alias("width"),
        (pl.col("length") / 24.5).alias("length")
    ]).select([
        'due_date', 'order_number', 'width', 'length', 'demand', 'quantity', 'type'
    ])

    # เพิ่มการกรองตามช่วงวันที่หากกำหนดมา
    if start_date or end_date:
        # สร้างฟังก์ชันแปลงสตริงเป็น datetime.date
        def parse_date(date_str: str) -> date:
            return datetime.strptime(date_str, "%Y-%m-%d").date()
        
        conditions = []
        
        if start_date:
            start = parse_date(start_date)
            conditions.append(pl.col("due_date") >= start)
        if end_date:
            end = parse_date(end_date)
            conditions.append(pl.col("due_date") <= end)
        
        df = df.filter(pl.all_horizontal(conditions))

    print("Data cleaning complete.")
    return df

if __name__ == "__main__":
    input_file = "order2024.csv"  # Assuming this file is in the same directory
    
    # Load the data
    raw_df = load_data(input_file)

    if raw_df.height > 0: # Check if DataFrame is not empty
        print("\nRaw Data Head:")
        print(raw_df.head())

        # Clean the data (example with no date filtering for CLI)
        cleaned_df = clean_data(raw_df)

        print("\nCleaned Data Head:")
        print(cleaned_df.head())
        
        # Write cleaned data to CSV
        cleaned_df.write_csv(
            "clean_order2024.csv", 
            include_header=True,
            encoding='utf8'
        )
