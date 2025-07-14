import os
import time

import polars as pl
from PyQt5.QtCore import QMutex, QMutexLocker, QObject, pyqtSignal

# สมมติว่า cleaning.py อยู่ในไดเรกทอรีเดียวกันและมีฟังก์ชันเหล่านี้
from cleaning import clean_stock, load_data


class StockManager(QObject):
    """
    จัดการการโหลดและรีเฟรชข้อมูลสต็อกเป็นระยะในเธรดแยก
    ส่งสัญญาณเพื่อสื่อสารกับเธรด UI หลัก
    """
    stock_updated = pyqtSignal(object)  # ใช้ object สำหรับ DataFrame
    error_signal = pyqtSignal(str)
    file_not_found_signal = pyqtSignal(str)

    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self._file_path = file_path
        self._is_running = False
        self._mutex = QMutex()
        self._file_exists = True  # สมมติว่าไฟล์มีอยู่ตอนเริ่มต้น

    def set_file_path(self, file_path):
        """เมธอดที่ปลอดภัยต่อเธรดเพื่ออัปเดตเส้นทางไฟล์"""
        with QMutexLocker(self._mutex):
            self._file_path = file_path
            # รีเซ็ตสถานะเพื่อบังคับให้ตรวจสอบใหม่
            self._file_exists = True

    def run(self):
        """
        ลูปหลักสำหรับเธรดจัดการสต็อก
        ตรวจสอบไฟล์สต็อก ประมวลผล แล้วรอ 60 วินาที
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

                # โหลดและทำความสะอาดข้อมูล
                raw_stock_df = load_data(current_path)
                if raw_stock_df is not None and not raw_stock_df.is_empty():
                    cleaned_stock_df = clean_stock(raw_stock_df)
                    self.stock_updated.emit(cleaned_stock_df)
                else:
                    # หากไฟล์ว่างหรือโหลดไม่สำเร็จ ให้ส่ง DataFrame ที่ว่างเปล่า
                    self.stock_updated.emit(pl.DataFrame())

            except Exception as e:
                self.error_signal.emit(
                    f"เกิดข้อผิดพลาดในการประมวลผลไฟล์สต็อก '{self._file_path}':\n{e}"
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
