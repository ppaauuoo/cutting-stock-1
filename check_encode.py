import polars as pl
import sys

# กำหนดให้ sys.stdout ใช้ UTF-8 สำหรับแสดงผลบนคอนโซล เพื่อรองรับตัวอักษรภาษาไทย
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

file_path='d:/order.csv'

df = pl.read_csv(
    file_path,
    separator=';',  # ระบุตัวคั่นเป็นเซมิโคลอน
    encoding='TIS-620',  # เปลี่ยน encoding เป็น TIS-620 สำหรับภาษาไทย
    null_values=["", " ", "NULL", "N/A"],
    skip_rows=1,            # ข้ามบรรทัดแรกที่ระบุ 'sep=;'
    has_header=True,        # ระบุว่าบรรทัดที่ 2 เป็นส่วนหัวของคอลัมน์
    truncate_ragged_lines=True # จัดการกับบรรทัดที่มีจำนวนคอลัมน์ไม่เท่ากัน
)


print(df.head(5))
