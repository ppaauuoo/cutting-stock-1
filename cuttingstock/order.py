import os
import shutil
import time

import polars as pl
from PyQt5.QtCore import QMutex, QMutexLocker, QObject, pyqtSignal

# สมมติว่า cleaning.py อยู่ในไดเรกทอรีเดียวกันและมีฟังก์ชันเหล่านี้
from cuttingstock.cleaning import clean_data, load_data


class OrderManager(QObject):
    """
    จัดการการโหลดและรีเฟรชข้อมูลออเดอร์เป็นระยะในเธรดแยก
    ส่งสัญญาณเพื่อสื่อสารกับเธรด UI หลัก
    """
    order_updated = pyqtSignal(object)  # ใช้ object สำหรับ DataFrame
    error_signal = pyqtSignal(str)
    file_not_found_signal = pyqtSignal(str)

    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self._file_path = file_path
        self._is_running = False
        self._mutex = QMutex()
        self._file_exists = True  # สมมติว่าไฟล์มีอยู่ตอนเริ่มต้น
        self._last_mod_time = 0

    def set_file_path(self, file_path):
        """เมธอดที่ปลอดภัยต่อเธรดเพื่ออัปเดตเส้นทางไฟล์"""
        with QMutexLocker(self._mutex):
            self._file_path = file_path
            # รีเซ็ตสถานะเพื่อบังคับให้ตรวจสอบใหม่
            self._file_exists = True
            self._last_mod_time = 0

    def run(self):
        """
        ลูปหลักสำหรับเธรดจัดการออเดอร์
        ตรวจสอบไฟล์ออเดอร์ ประมวลผล แล้วรอ 60 วินาที
        """
        self._is_running = True
        while self._is_running:
            try:
                with QMutexLocker(self._mutex):
                    current_path = self._file_path

                if not os.path.exists(current_path):
                    # ส่งสัญญาณเฉพาะเมื่อตรวจพบว่าไฟล์หายไปครั้งแรก
                    if self._file_exists:
                        self.file_not_found_signal.emit(current_path)
                        self._file_exists = False
                    # รอก่อนที่จะพยายามอีกครั้งเพื่อหลีกเลี่ยง busy-waiting
                    time.sleep(5)
                    continue

                # หากพบไฟล์ ให้รีเซ็ตแฟล็ก
                self._file_exists = True

                try:
                    mod_time = os.path.getmtime(current_path)
                    if mod_time != self._last_mod_time:
                        # Use a temporary copy to avoid issues with file locks
                        temp_csv_path = current_path + ".tmp"
                        shutil.copy2(current_path, temp_csv_path)

                        # โหลดข้อมูลจากไฟล์ CSV ที่คัดลอกมา
                        raw_order_df = load_data(temp_csv_path)
                        os.remove(temp_csv_path)  # ลบไฟล์ชั่วคราว

                        if raw_order_df is not None and not raw_order_df.is_empty():
                            # บันทึกข้อมูลลงในฐานข้อมูล SQLite เพื่อใช้เป็นแคช
                            cache_dir = "cache"
                            os.makedirs(cache_dir, exist_ok=True)
                            base_filename = os.path.splitext(os.path.basename(current_path))[0]
                            cache_db_path = os.path.join(cache_dir, f"{base_filename}.db")
                            table_name = base_filename
                            conn_str = f"sqlite:///{os.path.abspath(cache_db_path)}"
                            raw_order_df.write_database(table_name, connection=conn_str, if_table_exists="replace")

                            # ประมวลผลข้อมูลที่โหลดมา
                            cleaned_order_df = clean_data(raw_order_df, suggestion_mode=True)
                            self.order_updated.emit(cleaned_order_df)
                        else:
                            # หากไฟล์ต้นฉบับว่างเปล่า ให้ส่ง DataFrame ที่ว่างเปล่า
                            self.order_updated.emit(pl.DataFrame())

                        self._last_mod_time = mod_time
                except (IOError, PermissionError):
                    # File might be locked. Silently ignore and retry in the next cycle.
                    pass

            except Exception as e:
                self.error_signal.emit(
                    f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ออเดอร์ '{self._file_path}':\n{e}"
                )
                # หลีกเลี่ยงข้อความแสดงข้อผิดพลาดที่รวดเร็วสำหรับปัญหาเดียวกัน
                time.sleep(10)

            # รอ 60 วินาทีก่อนรอบถัดไป
            # ลูปนี้ช่วยให้ออกจากโปรแกรมได้เร็วขึ้นหากเรียกใช้ stop()
            for _ in range(60):
                if not self._is_running:
                    break
                time.sleep(1)

    def stop(self):
        """หยุดการทำงานของลูป"""
        self._is_running = False
