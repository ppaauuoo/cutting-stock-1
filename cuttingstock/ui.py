import collections
import csv
import os
import re
import sys
import time
from math import floor

import polars as pl
from PyQt5.QtCore import (
    QDateTime,
    QLocale,
    Qt,
    QTextCodec,
    QThread,
    QTimer,
)
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QDialog,
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from cuttingstock.core import generate_suggestions
from cuttingstock.material import handle_unprocessed_orders
from cuttingstock.order import OrderManager, filter_orders_by_factory
from cuttingstock.stock import StockManager
from cuttingstock.widget import CustomTableWidget
from cuttingstock.dialog import MaterialSubstitutionDialog
from cuttingstock.worker import WorkerThread
from cuttingstock.export import ExportManager

class CuttingOptimizerUI(QMainWindow):

    def __init__(self):
        super().__init__()
        self.qtapp = QApplication.instance()
        if self.qtapp:
            self.qtapp.setQuitOnLastWindowClosed(False)
        self.setWindowTitle("กระดาษม้วนตัด Optimizer")
        self.setGeometry(100, 100, 800, 700)

        self.ROLL_SPECS = {}
        self.cleaned_orders_df = None
        self.calculated_length = 0
        self.suggestions_list = []
        self.current_suggestion_index = 0
        self.processed_order_numbers = set()
        self.material_substitutions = {}

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)

        # Default file paths
        self.order_file_path = "order.csv"
        self.stock_file_path = "stock.csv"

        layout.addWidget(QLabel(f"Order File: {self.order_file_path}"))
        layout.addWidget(QLabel(f"Stock File: {self.stock_file_path}"))

        # Factory selection
        factory_layout = QHBoxLayout()
        factory_layout.addWidget(QLabel("โรงงาน:"))
        self.factory_combo = QComboBox()
        self.factory_combo.addItems(["รวม", "1", "2", "3", "4", "5"])
        factory_layout.addWidget(self.factory_combo)
        layout.addLayout(factory_layout)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)  # ตั้งค่าระยะ 0-100%
        self.progress_bar.setTextVisible(True)  # แสดงข้อความ
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setFormat("กำลังรอการเริ่มต้น...") # ตั้งค่าข้อความเริ่มต้น
        layout.addWidget(self.progress_bar)

        # Buttons layout
        buttons_layout = QHBoxLayout()

        self.run_button = QPushButton("เริ่มการคำนวณอัตโนมัติ")
        self.run_button.clicked.connect(self.start_main_loop)
        buttons_layout.addWidget(self.run_button)

        self.export_button = QPushButton("ส่งออก")
        self.export_button.clicked.connect(self.export_results)
        buttons_layout.addWidget(self.export_button)

        self.clear_button = QPushButton("ล้างผลลัพธ์")
        self.clear_button.clicked.connect(self.clear_results)
        # LEGACY: clear button
        # buttons_layout.addWidget(self.clear_button)

        self.show_unprocessed_checkbox = QCheckBox("แสดงออเดอร์ที่ไม่สามารถออกได้")
        self.show_unprocessed_checkbox.setChecked(True)
        self.show_unprocessed_checkbox.toggled.connect(self._refresh_results_display)
        buttons_layout.addWidget(self.show_unprocessed_checkbox)

        layout.addLayout(buttons_layout)

        # Log display (Collapsible)
        self.log_group_box = QGroupBox("Show Logs:")
        self.log_group_box.setCheckable(True) # ทำให้ GroupBox ยุบ/ขยายได้
        self.log_group_box.setChecked(False)  # เริ่มต้นให้ยุบอยู่

        log_layout = QVBoxLayout()
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        log_layout.addWidget(self.log_display)
        self.log_group_box.setLayout(log_layout)

        layout.addWidget(self.log_group_box)

        # เชื่อมต่อ signal เพื่อซ่อน/แสดง log_display
        self.log_group_box.toggled.connect(self.log_display.setVisible)
        self.log_display.setVisible(False) # ตรวจสอบให้แน่ใจว่าสถานะเริ่มต้นตรงกัน

        # เพิ่มตารางแสดงผล
        layout.addWidget(QLabel("ผลลัพธ์การตัด:"))
        self.result_table = CustomTableWidget() # ใช้ CustomTableWidget
        self.result_table.setColumnCount(12)
        self.result_table.setHorizontalHeaderLabels([
            "ความกว้างม้วน", "หมายเลขออเดอร์", "กำหนดส่ง", "ชนิดส่วนประกอบ", "ความกว้างออเดอร์", "จำนวนออก", "เศษเหลือ",
            "ความยาวออเดอร์", "จำนวนสั่งส่ง", "ผลิตได้", "จำนวนสั่งผลิต", "ปริมาณตัด"
            #, "กระดาษที่ใช้", "กระดาษคงเหลือ"
        ])
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        # ตั้งค่าตารางให้สามารถเลือกและคัดลอกได้
        self.result_table.setSelectionMode(QTableWidget.SingleSelection)
        self.result_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.result_table)

        self.setCentralWidget(central_widget)

        # เพิ่มตัวแปรสำหรับเก็บข้อมูลผลลัพธ์ทั้งหมด
        self.results_data = []
        self.display_data = []
        # เชื่อมต่อสัญญาณ enterPressed ของตารางไปยังเมธอดที่แสดงป๊อปอัป
        self.result_table.enterPressed.connect(self.show_row_details_popup)
        # เชื่อมต่อสัญญาณ doubleClicked ของตารางไปยังเมธอดที่แสดงป๊อปอัป
        self.result_table.doubleClicked.connect(self.show_row_details_popup)

        self.setup_order_manager()
        self.setup_stock_manager()

    def setup_order_manager(self):
        """เริ่มต้นและเริ่มการทำงานของเธรดจัดการออเดอร์"""
        order_file_path = self.order_file_path
        self.order_thread = QThread()
        self.order_manager = OrderManager(order_file_path)
        self.order_manager.moveToThread(self.order_thread)

        # เชื่อมต่อสัญญาณจาก manager ไปยัง slots ของ UI
        self.order_manager.order_updated.connect(self.update_order_data)
        self.order_manager.error_signal.connect(self.handle_order_error)
        self.order_manager.file_not_found_signal.connect(self.handle_order_file_not_found)

        # เชื่อมต่อสัญญาณของเธรด
        self.order_thread.started.connect(self.order_manager.run)

        # เริ่มการทำงานของเธรด
        self.order_thread.start()

    def handle_order_file_not_found(self, file_path):
        """แสดงกล่องคำเตือนแบบ modal เมื่อไม่พบไฟล์ออเดอร์"""
        self.log_message(f"⚠️ ไม่พบไฟล์ออเดอร์: {file_path}. กรุณาตรวจสอบตำแหน่งไฟล์.")
        QMessageBox.warning(
            self,
            "ไม่พบไฟล์",
            f"ไม่พบไฟล์ออเดอร์ที่ระบุ:\n{file_path}\n\nโปรแกรมจะพยายามโหลดไฟล์อีกครั้งในภายหลัง"
        )

    def handle_order_error(self, error_message):
        """บันทึกข้อผิดพลาดจาก order manager"""
        self.log_message(f"❌ เกิดข้อผิดพลาดกับ Order Manager: {error_message}")

    def update_order_data(self, order_df):
        """อัปเดต DataFrame ออเดอร์ที่ทำความสะอาดแล้ว"""
        timestamp = convert_thai_digits_to_arabic(QDateTime.currentDateTime().toString("hh:mm:ss"))
        if order_df is not None:
            self.cleaned_orders_df = order_df
            self.log_message(f"[{timestamp}] 🔄 อัปเดตข้อมูลออเดอร์เรียบร้อยแล้ว")
        else:
            self.cleaned_orders_df = None # หรือ pl.DataFrame()
            self.log_message(f"[{timestamp}] ℹ️ ข้อมูลออเดอร์ว่างเปล่าหรือไม่สามารถโหลดได้")

    def closeEvent(self, event):
        """หยุดการทำงานของ worker threads อย่างถูกต้องเมื่อปิดโปรแกรม"""
        self.log_message("กำลังปิดโปรแกรม...")
        self.run_button.setEnabled(False) # ป้องกันการคลิกซ้ำ

        # --- Phase 1: Request all threads to stop ---
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.log_message("กำลังส่งคำขอหยุดการคำนวณ...")
            self.worker.requestInterruption()
        if hasattr(self, 'order_manager'):
            self.order_manager.stop()
        if hasattr(self, 'stock_manager'):
            self.stock_manager.stop()

        QApplication.processEvents()

        # --- Phase 2: Wait for all threads to finish gracefully ---
        threads_to_wait = []
        if hasattr(self, 'worker') and self.worker.isRunning():
            threads_to_wait.append(("Worker", self.worker))
        if hasattr(self, 'order_thread') and self.order_thread.isRunning():
            threads_to_wait.append(("Order Manager", self.order_thread))
        if hasattr(self, 'stock_thread') and self.stock_thread.isRunning():
            threads_to_wait.append(("Stock Manager", self.stock_thread))

        for name, thread in threads_to_wait:
            self.log_message(f"กำลังรอให้เธรด {name} หยุดทำงาน...")
            if not thread.wait(5000):  # 5-second timeout
                self.log_message(f"⚠️ เธรด {name} ไม่หยุดทำงานในเวลาที่กำหนด, กำลังบังคับปิด.")
                thread.terminate()
                thread.wait()

        self.log_message("ปิดโปรแกรมเรียบร้อยแล้ว")
        event.accept()

    def _pause_background_threads(self):
        """Stops and cleans up the background file monitoring threads gracefully."""
        self.log_message("ℹ️ Pausing background file monitoring...")
        if hasattr(self, 'order_thread') and self.order_thread and self.order_thread.isRunning():
            self.order_manager.stop()
            self.order_thread.quit()  # Request normal exit
            if not self.order_thread.wait(5000):  # Extended timeout
                self.order_manager.stop()  # Ensure worker stops
                self.order_thread.terminate()  # Force exit if needed
                self.order_thread.wait()  # Block until thread finishes
            # Add state verification
            if self.order_thread.isRunning():
                self.log_message("❌ Thread still running after termination")
                return  # Block deletion until thread is fully stopped
            self.order_manager.deleteLater()
            self.order_thread.deleteLater()
            self.order_manager = None
            self.order_thread = None

        if hasattr(self, 'stock_thread') and self.stock_thread and self.stock_thread.isRunning():
            self.stock_manager.stop()
            self.stock_thread.quit()
            if not self.stock_thread.wait(3000):
                self.log_message("⚠️ Stock manager thread did not stop gracefully. Terminating.")
                self.stock_thread.terminate()
                self.stock_thread.wait()
            self.stock_manager.deleteLater()
            self.stock_thread.deleteLater()
            self.stock_manager = None
            self.stock_thread = None

    def _resume_background_threads(self):
        """Resumes the background file monitoring by creating new threads."""
        self.log_message("ℹ️ Resuming background file monitoring...")
        if not (hasattr(self, 'order_thread') and self.order_thread and self.order_thread.isRunning()):
             self.setup_order_manager()
        if not (hasattr(self, 'stock_thread') and self.stock_thread and self.stock_thread.isRunning()):
             self.setup_stock_manager()

    def setup_stock_manager(self):
        """เริ่มต้นและเริ่มการทำงานของเธรดจัดการสต็อก"""
        stock_file_path = self.stock_file_path
        self.stock_thread = QThread()
        self.stock_manager = StockManager(stock_file_path)
        self.stock_manager.moveToThread(self.stock_thread)

        # เชื่อมต่อสัญญาณจาก manager ไปยัง slots ของ UI
        self.stock_manager.stock_updated.connect(self.update_stock_data)
        self.stock_manager.error_signal.connect(self.handle_stock_error)
        self.stock_manager.file_not_found_signal.connect(self.handle_stock_file_not_found)

        # เชื่อมต่อสัญญาณของเธรด
        self.stock_thread.started.connect(self.stock_manager.run)

        # เริ่มการทำงานของเธรด
        self.stock_thread.start()

    def handle_stock_file_not_found(self, file_path):
        """แสดงกล่องคำเตือนแบบ modal เมื่อไม่พบไฟล์สต็อก"""
        self.log_message(f"⚠️ ไม่พบไฟล์สต็อก: {file_path}. กรุณาตรวจสอบตำแหน่งไฟล์.")
        QMessageBox.warning(
            self,
            "ไม่พบไฟล์",
            f"ไม่พบไฟล์สต็อกที่ระบุ:\n{file_path}\n\nโปรแกรมจะพยายามโหลดไฟล์อีกครั้งในภายหลัง"
        )

    def handle_stock_error(self, error_message):
        """บันทึกข้อผิดพลาดจาก stock manager"""
        self.log_message(f"❌ เกิดข้อผิดพลาดกับ Stock Manager: {error_message}")

    def update_stock_data(self, stock_df):
        """อัปเดต ROLL_SPECS จาก DataFrame สต็อกที่ทำความสะอาดแล้ว ให้มีโครงสร้างตาม mock-up"""
        new_roll_specs = {}
        if stock_df is not None and not stock_df.is_empty():
            try:
                required_cols = ["roll_size", "roll_type", "length", "roll_number"]
                # ตรวจสอบว่ามีคอลัมน์ที่จำเป็นอยู่
                if all(col in stock_df.columns for col in required_cols):
                    # วนลูปตามข้อมูลสต็อกแต่ละม้วนโดยไม่มีการรวมกลุ่ม
                    for row in stock_df.iter_rows(named=True):
                        roll_number = str(row['roll_number']).strip()
                        if not roll_number or roll_number.isspace():
                            self.log_message(f"⚠️ คำเตือน: พบม้วนที่ไม่มี 'roll_number' ในไฟล์สต็อก, ข้อมูลแถว: {row}")
                            continue # ข้ามม้วนที่ไม่มี ID ที่ถูกต้อง

                        width = str(row['roll_size']).strip()
                        material = str(row['roll_type']).strip()
                        length = row['length']

                        #TEST
                        length = 10000000

                        if width not in new_roll_specs:
                            new_roll_specs[width] = {}
                        if material not in new_roll_specs[width]:
                            new_roll_specs[width][material] = {}

                        # ใช้ key ที่เพิ่มขึ้นเรื่อยๆ สำหรับแต่ละม้วนภายใต้ width/material เดียวกัน
                        roll_key = len(new_roll_specs[width][material]) + 1

                        new_roll_specs[width][material][roll_key] = {
                            'id': roll_number,
                            'length': length
                        }
                else:
                    missing_cols = [col for col in required_cols if col not in stock_df.columns]
                    self.log_message(f"⚠️ ไฟล์สต็อกไม่มีคอลัมน์ที่ต้องการ: {', '.join(missing_cols)}")

            except Exception as e:
                self.handle_stock_error(f"ไม่สามารถประมวลผลข้อมูลสต็อกได้: {e}")
                return

        if self.ROLL_SPECS != new_roll_specs:
            timestamp = convert_thai_digits_to_arabic(QDateTime.currentDateTime().toString("hh:mm:ss"))
            self.ROLL_SPECS = new_roll_specs
            self.log_message(f"[{timestamp}] 🔄 อัปเดตข้อมูลสต็อกเรียบร้อยแล้ว")

    # LEGACY : don't use this logic anymore.
    def calculate_length_for_suggestion(self, width, spec):
        """Calculate effective roll length for a given suggestion."""
        selected_materials = [
            spec.get('front', ''),
            spec.get('c', ''),
            spec.get('middle', ''),
            spec.get('b', ''),
            spec.get('back', '')
        ]

        material_counts = collections.Counter(m for m in selected_materials if m)
        unique_materials = list(material_counts.keys())

        effective_lengths = []
        if not unique_materials or not width:
            return 0

        if width not in self.ROLL_SPECS:
            return 0

        stock_data_for_width = self.ROLL_SPECS[width]
        for material in unique_materials:
            material_rolls = stock_data_for_width.get(material)
            if material_rolls:
                min_length = min(roll['length'] for roll in material_rolls.values())
                usage_count = material_counts[material]
                if usage_count == 0: continue

                roll_used = floor(len(material_rolls) / usage_count)
                effective_length = min_length * roll_used
                effective_lengths.append(effective_length)
            else:
                # If any required material is not in stock for this width, this suggestion is invalid for calculation.
                return 0

        if effective_lengths:
            min_effective_length = min(effective_lengths)
            # return int(min_effective_length)
            return 1000000000
        else:
            return 0

    def log_message(self, message: str):
        self.log_display.append(message)
        self.log_display.verticalScrollBar().setValue(
            self.log_display.verticalScrollBar().maximum()
        )

    def get_all_suggestions(self):
        """
        Delegates suggestion generation to the core module and handles UI feedback.
        """
        if self.cleaned_orders_df is None or self.cleaned_orders_df.is_empty():
            self.log_message("⚠️ Cannot generate suggestions: No order data available.")
            return []

        self.log_message("🤔 Analyzing orders to generate all possible settings...")
        try:
            selected_factory = self.factory_combo.currentText()
            filtered_orders = filter_orders_by_factory(self.cleaned_orders_df, selected_factory)
            self.log_message(f"🏭 Using factory filter: '{selected_factory}'")

            suggestions = generate_suggestions(
                filtered_orders, self.ROLL_SPECS, selected_factory
            )

            self.log_message(f"✅ Generated {len(suggestions)} potential settings to test.")
            return suggestions

        except Exception as e:
            QMessageBox.critical(self, "Error Generating Suggestions", f"An error occurred: {e}")
            self.log_message(f"❌ Error during suggestion generation: {e}")
            return []

    def start_main_loop(self):
        self.start_time = time.time()
        timestamp = convert_thai_digits_to_arabic(QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss"))
        self.log_message(f"[{timestamp}] 🚀 Starting automated calculation process...")
        self.run_button.setEnabled(False)

        self.results_data.clear()
        self.processed_order_numbers.clear()
        self.material_substitutions.clear()
        self.result_table.setRowCount(0)

        self.log_display.clear()

        self._pause_background_threads()

        self.suggestions_list = self.get_all_suggestions()
        if not self.suggestions_list:
            self.log_message("⏹️ No suggestions found. Process finished.")
            self.run_button.setEnabled(True)
            self._resume_background_threads()
            QMessageBox.information(self, "No Suggestions", "No valid settings could be suggested from the order data.")
            return

        self.log_message(f"Found {len(self.suggestions_list)} settings to test.")
        self.current_suggestion_index = 0
        self.run_next_calculation()

    def run_next_calculation(self):
        if self.current_suggestion_index >= len(self.suggestions_list):
            timestamp = convert_thai_digits_to_arabic(QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss"))
            self.log_message(f"[{timestamp}] ✅ All suggestions processed. Automated calculation finished.")

            if hasattr(self, 'start_time'):
                end_time = time.time()
                elapsed_time = end_time - self.start_time
                self.log_message(f"⏱️ Total elapsed time: {elapsed_time:.2f} seconds.")

            if self.cleaned_orders_df is not None and not self.cleaned_orders_df.is_empty():
                selected_factory = self.factory_combo.currentText()
                filtered_orders_df = filter_orders_by_factory(self.cleaned_orders_df, selected_factory)
                rem_orders_df = filtered_orders_df.filter(
                    ~pl.col("order_number").is_in(list(self.processed_order_numbers))
                )
                if not rem_orders_df.is_empty():
                    self.log_message(f"ℹ️ Handling {rem_orders_df.shape[0]} unprocessed orders.")
                    unprocessed_results = handle_unprocessed_orders(rem_orders_df, self.log_message)
                    self.append_results_to_table(unprocessed_results)

            if self.results_data:
                self.log_message("Sorting final results by roll width...")
                try:
                    # Sort the results data in place by roll_w, treating it as an integer.
                    self.results_data.sort(key=lambda r: int(r.get('roll_w', 0)))
                    # Repopulate the table with the sorted data by calling append_results_to_table
                    # with an empty list. This re-uses the existing repopulation logic.
                    self.append_results_to_table([])
                except (ValueError, TypeError) as e:
                    self.log_message(f"⚠️ Could not sort results by roll width: {e}")

            self.run_button.setEnabled(True)
            self.progress_bar.setFormat("✅ Finished all tasks!")
            self._resume_background_threads()
            QMessageBox.information(self, "Finished", "All suggested settings have been processed.")
            return

        suggestion = self.suggestions_list[self.current_suggestion_index]
        width_str = suggestion['width']
        spec = suggestion['spec']

        try:
            width = int(width_str)
        except ValueError:
            self.log_message(f"⚠️ Skipping suggestion with invalid width: {width_str}")
            self.on_calculation_error(f"Invalid width '{width_str}'")
            return

        length = self.calculate_length_for_suggestion(width_str, spec)
        if length <= 0:
            self.log_message(f"ℹ️ Skipping suggestion {self.current_suggestion_index + 1} due to zero calculated length (insufficient stock).")
            # Manually advance to the next suggestion.
            # We can't call on_calculation_error() directly as there is no sender thread.
            self.current_suggestion_index += 1
            # Since no worker was created, the 'destroyed' signal won't fire,
            # so we trigger the next calculation manually.
            QTimer.singleShot(0, self.run_next_calculation)
            return

        front_material = spec.get('front') or None
        c_material = spec.get('c') or None
        middle_material = spec.get('middle') or None
        b_material = spec.get('b') or None
        back_material = spec.get('back') or None

        c_type = 'C' if c_material else None
        b_type = 'B' if b_material else None

        spec_str = ", ".join(f"{k}: {v}" for k, v in spec.items() if v)
        self.log_message(f"--- Running Suggestion {self.current_suggestion_index + 1}/{len(self.suggestions_list)} ---")
        self.log_message(f"Width: {width}, Length: {length}, Spec: {spec_str}")

        self.progress_bar.setValue(0)
        self.progress_bar.setFormat(f"Processing suggestion {self.current_suggestion_index + 1}...")


        selected_factory = self.factory_combo.currentText()
        self.worker = WorkerThread(
            width, length, None, None, self.order_file_path,
            front_material,
            c_type, c_material,
            middle_material,
            b_type, b_material,
            back_material,
            self.ROLL_SPECS,
            self.processed_order_numbers.copy(),
            self.material_substitutions,
            selected_factory
        )
        self.worker.update_signal.connect(self.log_message)
        self.worker.progress_updated.connect(self.update_progress_bar)
        self.worker.calculation_succeeded.connect(self.on_calculation_finished)
        self.worker.error_signal.connect(self.on_calculation_error)
        self.worker.out_of_stock_signal.connect(self.handle_out_of_stock)
        # เชื่อมต่อสัญญาณ destroyed เพื่อให้แน่ใจว่าเธรดเก่าถูกลบอย่างสมบูรณ์
        # ก่อนที่จะเริ่มการคำนวณครั้งถัดไปโดยอัตโนมัติ
        self.worker.destroyed.connect(self.run_next_calculation)
        self.worker.start()

    def handle_out_of_stock(self, details: dict):
        """Shows a custom dialog to the user to select a new material."""
        width = details.get("width")
        material = details.get("material")
        material_specs = details.get("material_specs", {})
        known_out_of_stock = details.get("known_out_of_stock", [])

        available_materials = []
        if width and str(width) in self.ROLL_SPECS:
            available_materials = sorted(list(self.ROLL_SPECS[str(width)].keys()))

        if not available_materials:
            QMessageBox.warning(self, "ไม่มีสต็อก", f"ไม่มีวัสดุอื่นสำหรับความกว้าง {width} นิ้ว ในสต็อก")
            self.worker.set_user_choice(None)
            return

        dialog = MaterialSubstitutionDialog(self, str(width), available_materials, material, material_specs, known_out_of_stock)

        result = dialog.exec_()
        if result == QDialog.Accepted:
            new_specs = dialog.get_selected_specs()
            changes = {k: v for k, v in new_specs.items() if material_specs.get(k) != v}
            changes_str = ", ".join([f"{k.title()}: {v}" for k, v in changes.items()])
            self.log_message(f"ผู้ใช้เลือกวัสดุทดแทน: {changes_str}")
            self.worker.set_user_choice(new_specs)
        elif result == 2: # Abort button
            self.log_message("ผู้ใช้สั่งหยุดการทำงานทั้งหมด")
            self.worker.requestInterruption()
            self.worker.set_user_choice(None)
            self.current_suggestion_index = len(self.suggestions_list)
        else: # Cancel button
            self.log_message("ผู้ใช้ยกเลิกการเลือกวัสดุทดแทน")
            self.worker.set_user_choice(None)

    def update_progress_bar(self, value: int, message: str):
        """Updates the progress bar."""
        self.progress_bar.setValue(value)
        if self.suggestions_list:
            self.progress_bar.setFormat(f"({self.current_suggestion_index + 1}/{len(self.suggestions_list)}) {message} ({value}%)")
        else:
            self.progress_bar.setFormat(f"{message} ({value}%)")

        if value == 100:
            pass

    def on_calculation_finished(self, results):
        if not results:
            self.log_message(f"ℹ️ Suggestion {self.current_suggestion_index + 1} was feasible but yielded no cutting patterns.")
        else:
            self.log_message(f"✅ Suggestion {self.current_suggestion_index + 1} finished with {len(results)} results.")
            self.append_results_to_table(results)
            for result in results:
                if order_num := result.get('order_number'):
                    self.processed_order_numbers.add(order_num)

        self.current_suggestion_index += 1
        sender_thread = self.sender()
        if sender_thread:
            # The thread has finished its work. We just need to wait for it to
            # fully exit and then schedule it for deletion. The 'destroyed'
            # signal will then trigger the next calculation.
            if not sender_thread.wait(5000):
                self.log_message("⚠️ Worker thread did not exit cleanly. Terminating.")
                sender_thread.terminate()
                sender_thread.wait()
            sender_thread.deleteLater()

    def _refresh_results_display(self):
        """Refreshes the results table display based on current filters."""
        self.append_results_to_table([])

    def append_results_to_table(self, results):
        """
        Adds new results, identifies the best entry (lowest trim) for each order,
        highlights duplicates in yellow, and sorts the entire list before repopulating the table.
        """
        # Add new results to the master list
        self.results_data.extend(results)

        # Find the best result (lowest trim) for each order number
        best_results_map = {}
        for res in self.results_data:
            order_num = res.get('order_number')
            try:
                # Ensure trim is a comparable number, default to infinity if missing/invalid
                trim = float(res.get('trim', float('inf')))
            except (ValueError, TypeError):
                trim = float('inf')

            if not order_num:
                continue

            # Check if this result is better (lower trim) than the one stored for the order
            current_best_trim = best_results_map.get(order_num, {}).get('trim', float('inf'))
            try:
                current_best_trim = float(current_best_trim)
            except (ValueError, TypeError):
                current_best_trim = float('inf')

            if trim < current_best_trim:
                best_results_map[order_num] = res

        # Create a set of IDs for the best results for quick lookup
        best_results_ids = {id(res) for res in best_results_map.values()}

        # Filter data based on UI controls
        if self.show_unprocessed_checkbox.isChecked():
            self.display_data = self.results_data
        else:
            self.display_data = [r for r in self.results_data if not isinstance(r.get('roll_w'), str)]

        # Repopulate the entire table
        self.result_table.setRowCount(0)
        self.result_table.setRowCount(len(self.display_data))

        for row_idx, result in enumerate(self.display_data):
            is_duplicate = id(result) not in best_results_ids
            is_unprocessed = isinstance(result.get('roll_w'), str)

            has_no_suitable_roll = False
            roll_info_keys = ['front_roll_info', 'c_roll_info', 'middle_roll_info', 'b_roll_info', 'back_roll_info']
            for key in roll_info_keys:
                if "ไม่มี" in result.get(key, ''):
                    has_no_suitable_roll = True
                    break

            cuts = result.get('cuts')
            order_qty = result.get('order_qty')
            demand_per_cut_val = ""
            if cuts is not None and cuts > 0 and order_qty is not None:
                demand_per_cut_val = f"{order_qty / cuts:.2f}"
            else:
                demand_per_cut_val = "N/A"

            for col_idx, value in enumerate([
                str(result.get('roll_w', '')),
                str(result.get('order_number', '')),
                str(result.get('due_date', '')),
                str(result.get('component_type', '')),
                f"{result.get('order_w', ''):.4f}",
                str(result.get('cuts') or '').split('.')[0],
                f"{result.get('trim', ''):.2f}",
                f"{result.get('order_l', ''):.4f}",
                f"{result.get('order_dmd', '')}",
                str(result.get('die_cut', '')),
                f"{result.get('order_qty', '')}",
                demand_per_cut_val,
            ]):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() | Qt.ItemIsSelectable | Qt.ItemIsEnabled)

                # Apply colors based on status (red takes precedence)
                if has_no_suitable_roll or is_unprocessed:
                    item.setBackground(QColor(255, 224, 224))  # Red for invalid rolls
                elif is_duplicate:
                    item.setBackground(QColor(255, 255, 224))  # Yellow for duplicates

                self.result_table.setItem(row_idx, col_idx, item)

        # Apply spans for grouped results
        processed_groups = set()
        for row_idx, result in enumerate(self.display_data):
            group_id = result.get("group_id")
            if group_id and group_id not in processed_groups:
                group_rows = [
                    i for i, r in enumerate(self.display_data)
                    if i >= row_idx and r.get("group_id") == group_id
                ]

                if len(group_rows) > 1:
                    start_row = group_rows[0]
                    row_span = len(group_rows)

                    # Span Roll Width (col 0) and Trim (col 6)
                    self.result_table.setSpan(start_row, 0, row_span, 1)
                    self.result_table.setSpan(start_row, 6, row_span, 1)

                    # Vertically center the text in the spanned cells
                    for col_to_center in [0, 6]:
                        item = self.result_table.item(start_row, col_to_center)
                        if item:
                            item.setTextAlignment(item.textAlignment() | Qt.AlignVCenter)

                processed_groups.add(group_id)

        self.result_table.resizeColumnsToContents()

    def on_calculation_error(self, error_message: str):
        self.log_message(f"❌ Error or infeasible on suggestion {self.current_suggestion_index + 1}: {error_message}")
        self.current_suggestion_index += 1
        sender_thread = self.sender()
        if sender_thread:
            # The thread has finished its work. We just need to wait for it to
            # fully exit and then schedule it for deletion. The 'destroyed'
            # signal will then trigger the next calculation.
            if not sender_thread.wait(5000):
                self.log_message("⚠️ Worker thread did not exit cleanly after error. Terminating.")
                sender_thread.terminate()
                sender_thread.wait()
            sender_thread.deleteLater()

    def clear_results(self):
        """Clears the results table and resets related data."""
        reply = QMessageBox.question(self, 'ยืนยันการล้างผลลัพธ์',
                                     "คุณแน่ใจหรือไม่ว่าต้องการล้างผลลัพธ์ทั้งหมด?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.results_data.clear()
            self.display_data.clear()
            self.processed_order_numbers.clear()
            self.material_substitutions.clear()
            self.result_table.setRowCount(0)
            self.log_message("🧹 ผลลัพธ์ทั้งหมดถูกล้างแล้ว")

    def export_results(self):
        """ส่งออกข้อมูลในตารางผลลัพธ์ไปยังไฟล์ CSV หรือ XLSX"""
        if not self.results_data:
            QMessageBox.information(self, "ไม่มีข้อมูล", "ไม่มีข้อมูลสำหรับส่งออก")
            return

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "บันทึกผลลัพธ์",
            "cutting_results",
            "Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)",
            options=options,
        )

        # Check if user canceled the dialog
        if not file_path:
            return

        # Add appropriate extension based on selected filter
        if selected_filter == "Excel Files (*.xlsx)" and not file_path.endswith('.xlsx'):
            file_path += '.xlsx'
        elif selected_filter == "CSV Files (*.csv)" and not file_path.endswith('.csv'):
            file_path += '.csv'

        try:
            headers = [self.result_table.horizontalHeaderItem(i).text() for i in range(self.result_table.columnCount())]
            # Filter out failed/unprocessed orders (those with string roll_w values)
            filtered_results = [result for result in self.results_data if not isinstance(result.get('roll_w'), str)]
            export_manager = ExportManager(filtered_results, headers)
            
            if file_path.endswith('.xlsx'):
                success = export_manager.export_to_xlsx(file_path)
                if success:
                    self.log_message(f"✅ ส่งออกผลลัพธ์ไปยัง {file_path} เรียบร้อยแล้ว")
                    QMessageBox.information(self, "ส่งออกสำเร็จ", f"บันทึกผลลัพธ์ไปยัง:\n{file_path} เรียบร้อยแล้ว")
                else:
                    self.log_message(f"❌ เกิดข้อผิดพลาดในการส่งออกเป็น XLSX")
                    QMessageBox.critical(self, "เกิดข้อผิดพลาดในการส่งออก", "ไม่สามารถส่งออกเป็น XLSX ได้ กรุณาตรวจสอบว่าติดตั้ง xlsxwriter แล้ว")
            else:
                success = export_manager.export_to_csv(file_path)
                if success:
                    self.log_message(f"✅ ส่งออกผลลัพธ์ไปยัง {file_path} เรียบร้อยแล้ว")
                    QMessageBox.information(self, "ส่งออกสำเร็จ", f"บันทึกผลลัพธ์ไปยัง:\n{file_path} เรียบร้อยแล้ว")
                else:
                    self.log_message(f"❌ เกิดข้อผิดพลาดในการส่งออกเป็น CSV")
                    QMessageBox.critical(self, "เกิดข้อผิดพลาดในการส่งออก", f"เกิดข้อผิดพลาดขณะส่งออกไฟล์")
        
        except Exception as e:
            self.log_message(f"❌ เกิดข้อผิดพลาดในการส่งออก: {e}")
            QMessageBox.critical(self, "เกิดข้อผิดพลาดในการส่งออก", f"เกิดข้อผิดพลาด:\n{e}")

    def _format_roll_usage_to_html(self, roll_info_str: str) -> str:
        """Parses roll usage string and formats it as an HTML table."""
        if not roll_info_str or "->" not in roll_info_str:
            return roll_info_str  # Return as is if empty or not in expected format

        if "(ไม่มี" in roll_info_str:
            # e.g., "-> (ไม่มีข้อมูลสต็อก)"
            return f"<i>{roll_info_str.replace('-> ', '')}</i>"

        parts = roll_info_str.split(': ', 1)
        if len(parts) < 2:
            return roll_info_str  # Fallback for unexpected format

        status_text = parts[0].replace('-> ', '').strip()
        roll_details_str = parts[1]

        roll_strings = roll_details_str.split(' + ')

        # Using a more robust regex to handle various whitespace and characters in roll ID
        roll_pattern = re.compile(r'(.+?)\s*\(ยาว\s*(\d+)\s*ม\.,\s*(?:เหลือ\s*(\d+)\s*ม\.|(ใช้หมด))\)')

        table_rows = []
        for roll_str in roll_strings:
            match = roll_pattern.match(roll_str.strip())
            if match:
                roll_id = match.group(1).strip()
                original_len = int(match.group(2))

                if match.group(4) and match.group(4) == "ใช้หมด":
                    remaining_len = 0
                else:
                    remaining_len = int(match.group(3)) if match.group(3) else 0

                used_len = original_len - remaining_len

                table_rows.append(f'<tr><td style="padding-right:10px;">{roll_id}</td><td align="right" style="padding-right:10px;">{original_len:,}</td><td align="right" style="padding-right:10px;">{used_len:,}</td><td align="right">{remaining_len:,}</td></tr>')
            else:
                # Fallback for unexpected format
                table_rows.append(f'<tr><td colspan="4" style="color: gray;"><i>(ข้อมูลไม่สมบูรณ์: {roll_str.strip()})</i></td></tr>')

        if not table_rows:
            return f"<i>{status_text}</i>"

        html = f'<table border="0" cellpadding="2" cellspacing="0" style="margin-top: 4px; margin-left: 15px; border-collapse: collapse;">'
        html += '<tr><th align="left" style="padding-right:10px; border-bottom: 1px solid black;">ID ม้วน</th><th align="right" style="padding-right:10px; border-bottom: 1px solid black;">ยาวเดิม (ม.)</th><th align="right" style="padding-right:10px; border-bottom: 1px solid black;">ใช้ไป (ม.)</th><th align="right" style="border-bottom: 1px solid black;">คงเหลือ (ม.)</th></tr>'
        html += "".join(table_rows)
        html += '</table>'

        return f"<i>{status_text}:</i>{html}"

    def _format_roll_usage_for_csv(self, roll_info_str: str) -> str:
        """Parses roll usage string and formats it for readable CSV export."""
        if not roll_info_str or "->" not in roll_info_str:
            return roll_info_str.replace('-> ', '').strip()

        if "(ไม่มี" in roll_info_str:
            return roll_info_str.replace('-> ', '').strip()

        parts = roll_info_str.split(': ', 1)
        if len(parts) < 2:
            return roll_info_str

        status_text = parts[0].replace('-> ', '').strip()
        roll_details_str = parts[1]

        roll_strings = roll_details_str.split(' + ')

        roll_pattern = re.compile(r'(.+?)\s*\(ยาว\s*(\d+)\s*ม\.,\s*(?:เหลือ\s*(\d+)\s*ม\.|(ใช้หมด))\)')

        csv_parts = [f"{status_text}:"]
        for roll_str in roll_strings:
            match = roll_pattern.match(roll_str.strip())
            if match:
                roll_id = match.group(1).strip()
                original_len = int(match.group(2))

                if match.group(4) and match.group(4) == "ใช้หมด":
                    remaining_len = 0
                else:
                    remaining_len = int(match.group(3)) if match.group(3) else 0

                used_len = original_len - remaining_len

                csv_parts.append(f"  ID: {roll_id}, ยาวเดิม: {original_len}, ใช้ไป: {used_len}, คงเหลือ: {remaining_len}")
            else:
                csv_parts.append(f"  (ข้อมูลไม่สมบูรณ์: {roll_str.strip()})")

        return "\n".join(csv_parts)

    def show_row_details_popup(self):
        """
        แสดงป๊อปอัปพร้อมรายละเอียดของแถวที่เลือกในตาราง (ใช้ HTML สำหรับการจัดรูปแบบ)
        """
        selected_rows = self.result_table.selectedIndexes()
        if not selected_rows:
            return

        row_index = selected_rows[0].row()
        try:
            # Use the filtered display_data list which matches the table view
            result = self.display_data[row_index]
        except (IndexError, TypeError):
            result = {}

        details = []
        for col_idx in range(self.result_table.columnCount()):
            item = self.result_table.item(row_index, col_idx)
            if item:
                header = self.result_table.horizontalHeaderItem(col_idx).text()
                details.append(f"<b>{header}:</b> {item.text()}")

        type_details = []
        if result.get('type'):
            type_details.append(f"<b>ประเภททับเส้น:</b> {result['type']}")
        if result.get('component_type'):
            type_details.append(f"<b>ชนิดส่วนประกอบ:</b> {result['component_type']}")

        if type_details:
            details.append("<br/><b>📌 ข้อมูลประเภท:</b>")
            details.extend(type_details)

        material_details_parts = []
        c_type = result.get('c_type', '')
        b_type = result.get('b_type', '')
        type_demand = 1.0
        if c_type == 'C': type_demand = 1.45
        elif b_type == 'B': type_demand = 1.35
        elif c_type == 'E' or b_type == 'E': type_demand = 1.25

        def create_material_html(label: str, material: str, value: float, roll_info: str) -> str:
            roll_html = self._format_roll_usage_to_html(roll_info)
            return f"<b>{label}:</b> {material} = {value:.2f}<br/>{roll_html}"

        if result.get('front'):
            value = result.get('demand_per_cut', 0) / type_demand
            material_details_parts.append(create_material_html("แผ่นหน้า", result.get('front'), value, result.get('front_roll_info', '')))

        if result.get('c'):
            c_material = result.get('c')
            if c_type == 'C':
                value = result.get('demand_per_cut', 0)
                material_details_parts.append(create_material_html("ลอน C", c_material, value, result.get('c_roll_info', '')))
            elif c_type == 'E':
                demand = result.get('demand_per_cut', 0)
                value = (demand / 1.35 * 1.25) if b_type == 'B' else demand
                material_details_parts.append(create_material_html("ลอน E", c_material, value, result.get('c_roll_info', '')))

        if result.get('middle'):
            value = result.get('demand_per_cut', 0) / type_demand
            material_details_parts.append(create_material_html("แผ่นกลาง", result.get('middle'), value, result.get('middle_roll_info', '')))

        if result.get('b'):
            b_material = result.get('b')
            demand = result.get('demand_per_cut', 0)
            if b_type == 'B':
                value = (demand / 1.45 * 1.35) if c_type == 'C' else demand
                material_details_parts.append(create_material_html("ลอน B", b_material, value, result.get('b_roll_info', '')))
            elif b_type == 'E':
                value = (demand / 1.45 * 1.25) if c_type == 'C' else demand
                material_details_parts.append(create_material_html("ลอน E", b_material, value, result.get('b_roll_info', '')))

        if result.get('back'):
            value = result.get('demand_per_cut', 0) / type_demand
            material_details_parts.append(create_material_html("แผ่นหลัง", result.get('back'), value, result.get('back_roll_info', '')))

        if material_details_parts:
            details.append("<br/><b>⚙️ ข้อมูลแผ่นและลอน:</b>")
            details.append("<br/><br/>".join(material_details_parts))

        detail_message = "<br/>".join(details)
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(detail_message)
        msg_box.setWindowTitle("รายละเอียดผลลัพธ์การตัด")
        msg_box.setTextFormat(Qt.RichText) # Ensure HTML is rendered
        msg_box.setTextInteractionFlags(Qt.TextSelectableByMouse)
        msg_box.exec_()

def convert_thai_digits_to_arabic(text: str) -> str:
    """Convert Thai digits to Arabic digits"""
    thai_digits = "๐๑๒๓๔๕๖๗๘๙"
    arabic_digits = "0123456789"
    translation_table = str.maketrans(thai_digits, arabic_digits)
    return text.translate(translation_table)

if __name__ == "__main__":
    if sys.platform == "win32":
        os.environ["QT_QPA_PLATFORM"] = "windows:fontengine=freetype"
        os.environ["PYTHONIOENCODING"] = "utf-8"

    QTextCodec.setCodecForLocale(QTextCodec.codecForName("UTF-8"))

    thai_locale = QLocale(QLocale.Thai, QLocale.Thailand)
    QLocale.setDefault(thai_locale)

    app = QApplication(sys.argv)
    app.setFont(QFont('Tahoma', 9))  # ตั้งค่าฟอนต์ภาษาไทย

    window = CuttingOptimizerUI()
    window.show()
    sys.exit(app.exec_())
