import sys
import unicodedata
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
            separator=';',  # ระบุตัวคั่นเป็นเซมิโคลอน
            encoding='TIS-620',  # เปลี่ยน encoding เป็น TIS-620 สำหรับภาษาไทย
            null_values=["", " ", "NULL", "N/A"],
            skip_rows=1,            # ข้ามบรรทัดแรกที่ระบุ 'sep=;'
            has_header=True,        # ระบุว่าบรรทัดที่ 2 เป็นส่วนหัวของคอลัมน์
            truncate_ragged_lines=True # จัดการกับบรรทัดที่มีจำนวนคอลัมน์ไม่เท่ากัน
        )
        print(f"Successfully loaded data from {file_path}")
        print(f"Data shape: {df.shape[0]} rows, {df.shape[1]} columns")
        print("Columns:", df.columns)
        # แสดงตัวอย่างข้อมูล 5 แถวแรก
        print("Sample data:")
        print(df.head(5))
        return df
    except Exception as e:
        print(f"Error loading data from {file_path}: {e}")
        raise

def clean_data(df: pl.DataFrame, 
               start_date: Optional[str] = None, 
               end_date: Optional[str] = None,
               front: Optional[str] = None, # เพิ่มพารามิเตอร์สำหรับกรองวัสดุ
               c: Optional[str] = None,
               middle: Optional[str] = None,
               b: Optional[str] = None,
               back: Optional[str] = None
               ) -> pl.DataFrame:
    """
    Placeholder function for data cleaning.
    You can add your specific cleaning logic here.

    Args:
        df (pl.DataFrame): The input DataFrame to clean.
        start_date (Optional[str]): Start date for filtering (YYYY-MM-DD).
        end_date (Optional[str]): End date for filtering (YYYY-MM-DD).
        front (Optional[str]): Filter for 'front' material.
        C (Optional[str]): Filter for 'C' material.
        middle (Optional[str]): Filter for 'middle' material.
        B (Optional[str]): Filter for 'B' material.
        back (Optional[str]): Filter for 'back' material.

    Returns:
        pl.DataFrame: The cleaned DataFrame.
    """
    print("Starting data cleaning...")
    
    # แปลงชื่อคอลัมน์ให้เป็นรูปแบบ Unicode มาตรฐานและตัดช่องว่าง
    def normalize_col_name(col: str) -> str:
        return unicodedata.normalize('NFC', col.strip())
    
    # สร้าง mapping สำหรับคอลัมน์ไทยพร้อมชื่อที่แปลงแล้ว
    thai_col_mapping = {
        normalize_col_name("กำหนดส่ง       "): "due_date",
        normalize_col_name(" เลขที่ใบสั่งขาย"): "order_number",
        normalize_col_name("กว้าง"): "width",
        normalize_col_name("ยาว"): "length",
        normalize_col_name("จำนวนสั่งส่ง   "): "demand",
        normalize_col_name("จำนวนสั่งผลิต"): "quantity",
        normalize_col_name("กระดาษหน้า"): "front",
        normalize_col_name("ลอนC"): "c",
        normalize_col_name("กระดาษกลาง"): "middle",
        normalize_col_name("ลอนB"): "b",
        normalize_col_name("กระดาษหลัง"): "back",
        normalize_col_name("ทับเส้น"): "type",
        normalize_col_name("ประเภทกล่อง"): "component_type",
        normalize_col_name("ผลิตได้"): "die_cut"
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
    required_cols = ["due_date", "order_number", "width", "length", "demand", "quantity",  "front", "c", "middle", "b", "back", "type", "component_type"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"⚠️ คอลัมน์หาย: {missing} โปรดตรวจสอบชื่อคอลัมน์ในไฟล์ CSV")
    
    df = df.with_columns(
        pl.col("due_date").str.strip_chars().str.strptime(pl.Date, "%d/%m/%y", strict=True), # Changed %Y to %y for 2-digit year, kept strict=True for debugging
        pl.col("order_number").str.strip_chars().cast(pl.Int64),
        pl.col("width").str.strip_chars().cast(pl.Float64),
        pl.col("length").str.strip_chars().cast(pl.Float64),
        # ทำความสะอาดข้อมูล 'demand' และ 'quantity' โดยการลบคอมม่าและแปลงเป็น Int64
        pl.col("demand").str.strip_chars().str.replace_all(",", "").cast(pl.Float64).cast(pl.Int64),
        pl.col("quantity").str.strip_chars().str.replace_all(",", "").cast(pl.Float64).cast(pl.Int64),
        pl.col("type").str.strip_chars(),
        pl.col("component_type").str.strip_chars(),
        pl.col("die_cut").str.strip_chars().cast(pl.Int64),
   )

    df = df.drop_nulls(subset=["due_date", "order_number", "width", "length", "demand", "quantity", "type", "component_type"])
    df = df.filter(pl.col("demand") > 0)
    df = df.filter(pl.col("width") > 0)
    df = df.filter(pl.col("length") > 0)
    df = df.with_columns([
        (pl.col("width")).alias("width"),
        (pl.col("length")).alias("length")
    ]).select([
        'due_date', 'order_number', 'width', 'length', 'demand', 'quantity', 'type', 'component_type', 'front', 'c', 'middle', 'b', 'back', 'die_cut' # เพิ่มคอลัมน์วัสดุ
    ])

    if df.height > 0:
        print(f"วันที่กำหนดส่งขั้นต่ำใน DataFrame (หลังการแยกวิเคราะห์และลบค่าว่าง): {df.select(pl.col('due_date').min()).item()}")
        print(f"วันที่กำหนดส่งสูงสุดใน DataFrame (หลังการแยกวิเคราะห์และลบค่าว่าง): {df.select(pl.col('due_date').max()).item()}")
    else:
        print("DataFrame ว่างเปล่าก่อนการกรองวันที่ ไม่มีวันที่กำหนดส่งให้ตรวจสอบ")

    # เพิ่มการกรองตามช่วงวันที่หากกำหนดมา
    if start_date or end_date:
        # สร้างฟังก์ชันแปลงสตริงเป็น datetime.date
        def parse_date(date_str: str) -> date:
            try:
                return datetime.strptime(date_str, "%Y-%m-%d").date()
            except ValueError:
                raise ValueError(f"Invalid date format for filter: '{date_str}'. Expected YYYY-MM-DD.")
        
        conditions = []
        
        if start_date:
            start = parse_date(start_date)
            conditions.append(pl.col("due_date") >= start)
            print(f"พารามิเตอร์ตัวกรอง 'start_date' แยกวิเคราะห์ได้เป็น: {start}")
        if end_date:
            end = parse_date(end_date)
            conditions.append(pl.col("due_date") <= end)
            print(f"พารามิเตอร์ตัวกรอง 'end_date' แยกวิเคราะห์ได้เป็น: {end}")
        
        df = df.filter(pl.all_horizontal(conditions))

    print("Data after date filtering:")
    print(df.head(5))  

    # เพิ่มการกรองตามวัสดุหากกำหนดมา (รองรับกรณีเป็น None/Null ด้วย)
    print("Filtering data based on material specifications...")
    print("Front:", front, "C:", c, "Middle:", middle, "B:", b, "Back:", back)

    material_conditions = []
    if front is not None:
        material_conditions.append(pl.col("front").str.contains(front, literal=False))
    if c is not None:
        material_conditions.append(pl.col("c").str.contains(c, literal=False))
    if middle is not None:
        material_conditions.append(pl.col("middle").str.contains(middle, literal=False))
    if b is not None:
        material_conditions.append(pl.col("b").str.contains(b, literal=False))
    if back is not None:
        material_conditions.append(pl.col("back").str.contains(back, literal=False))

    # ใช้ pl.all_horizontal เพื่อรวมเงื่อนไขทั้งหมดด้วย AND logic หากมีเงื่อนไข
    if material_conditions:
        df = df.filter(pl.all_horizontal(material_conditions))

    print("Data cleaning complete.")
    print(f"Cleaned data shape: {df.shape[0]} rows, {df.shape[1]} columns")
    print("Columns after cleaning:", df.columns)    
    print("Sample cleaned data:")
    print(df.head(5))
    return df

#depracted
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
        )
