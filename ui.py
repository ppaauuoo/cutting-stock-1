import asyncio
import collections
import copy
import os
import re
import sys
from math import floor

import polars as pl
from PyQt5.QtCore import (
    QDate,
    QLocale,
    Qt,
    QTextCodec,
    QThread,
    pyqtSignal,
)
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
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

import main  # Import our modified main module
from stock import StockManager


class WorkerThread(QThread):
    update_signal = pyqtSignal(str)
    progress_updated = pyqtSignal(int, str)  # เพิ่มสัญญาณใหม่สำหรับอัปเดตโปรเกรสบาร์
    finished = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def __init__(self, width, length, start_date, end_date, file_path,
                 front_material,
                 corrugate_c_type, corrugate_c_material_name,
                 middle_material,
                 corrugate_b_type, corrugate_b_material_name,
                 back_material,
                 roll_specs,
                 parent=None):
        super().__init__(parent)
        self.width = width
        self.length = length
        self.start_date = start_date
        self.end_date = end_date
        self.file_path = file_path
        self.front_material = front_material
        self.corrugate_c_type = corrugate_c_type
        self.corrugate_c_material_name = corrugate_c_material_name
        self.middle_material = middle_material
        self.corrugate_b_type = corrugate_b_type
        self.corrugate_b_material_name = corrugate_b_material_name
        self.back_material = back_material
        self.roll_specs = roll_specs

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        def progress_callback(message: str):
            self.update_signal.emit(message)
            # ส่งสัญญาณพร้อมเปอร์เซ็นต์ความคืบหน้าประมาณการ
            if "กำลังเริ่มการคำนวณ" in message:
                self.progress_updated.emit(5, message)
            elif "โหลดและจัดเรียงข้อมูลเรียบร้อย" in message:
                self.progress_updated.emit(20, message)
            elif "กำลังประมวลผลม้วน" in message:
                # สามารถปรับเปอร์เซ็นต์ได้ละเอียดขึ้นหากมีข้อมูลจำนวนม้วนทั้งหมด
                self.progress_updated.emit(50, message)
            elif "บันทึกผลลัพธ์ลงไฟล์ CSV เรียบร้อย" in message:
                self.progress_updated.emit(95, message)
            
        try:
            results = loop.run_until_complete(
                main.main_algorithm(
                    roll_width=self.width,
                    roll_length=self.length,
                    progress_callback=progress_callback,
                    start_date=self.start_date,
                    end_date=self.end_date,
                    file_path=self.file_path,
                    front=self.front_material,
                    c_type=self.corrugate_c_type,
                    c=self.corrugate_c_material_name,
                    middle=self.middle_material,
                    b_type=self.corrugate_b_type,
                    b=self.corrugate_b_material_name,
                    back=self.back_material,
                    roll_specs=self.roll_specs,
                )
            )
            self.progress_updated.emit(100, "✅ เสร็จสิ้น")  # สัญญาณเสร็จสมบูรณ์
            self.finished.emit(results)
        except Exception as e:
            self.error_signal.emit(f"Error: {str(e)}")
            self.progress_updated.emit(0, "❌ เกิดข้อผิดพลาด!") # รีเซ็ตโปรเกรสบาร์เมื่อเกิดข้อผิดพลาด
        finally:
            loop.close()

class CustomTableWidget(QTableWidget):
    """
    QTableWidget ที่กำหนดเองเพื่อส่งสัญญาณเมื่อกดปุ่ม Enter
    """
    enterPressed = pyqtSignal() # สัญญาณที่กำหนดเอง

    def __init__(self, parent=None): # เพิ่ม parent และเรียก super().__init__(parent)
        super().__init__(parent)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            if self.selectedItems(): # ตรวจสอบว่ามีการเลือกรายการอยู่หรือไม่
                self.enterPressed.emit()
                return
        super().keyPressEvent(event) # เรียกเมธอดของคลาสพื้นฐานสำหรับปุ่มอื่นๆ

class CuttingOptimizerUI(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("กระดาษม้วนตัด Optimizer")
        self.setGeometry(100, 100, 800, 700)

        self.ROLL_SPECS = {}
        self.calculated_length = 0

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        
        # เพิ่มส่วนเลือกไฟล์
        layout.addWidget(QLabel("เลือกไฟล์ออเดอร์:"))
        
        file_layout = QHBoxLayout()
        self.file_path_input = QLineEdit("order.csv")
        self.file_path_input.setPlaceholderText("เลือกไฟล์ CSV...")
        file_layout.addWidget(self.file_path_input)
        
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.select_file)
        file_layout.addWidget(browse_button)
        
        layout.addLayout(file_layout)


        # เพิ่มส่วนเลือกไฟล์สต็อก
        layout.addWidget(QLabel("เลือกไฟล์สต็อก:"))
        stock_file_layout = QHBoxLayout()
        self.stock_file_path_input = QLineEdit("stock.csv")
        self.stock_file_path_input.setPlaceholderText("เลือกไฟล์ CSV สต็อก...")
        stock_file_layout.addWidget(self.stock_file_path_input)

        browse_stock_button = QPushButton("Browse...")
        browse_stock_button.clicked.connect(self.select_stock_file)
        stock_file_layout.addWidget(browse_stock_button)
        layout.addLayout(stock_file_layout)


        layout.addWidget(QLabel("ความกว้างม้วนกระดาษ (inch):")) 
        self.width_combo = QComboBox()
        # self.width_combo.addItems(self.ROLL_SPECS.keys()) # จะถูกเติมค่าจากไฟล์สต็อก
        self.width_combo.setPlaceholderText("รอข้อมูลสต็อก...")
        layout.addWidget(self.width_combo)
                # เพิ่มช่องกรอกข้อมูลแผ่นและลอนในแถวเดียวกัน
        material_layout = QHBoxLayout()

        material_layout.addWidget(QLabel("แผ่นหน้า:"))
        self.sheet_front_input = QComboBox() # Changed to QComboBox
        material_layout.addWidget(self.sheet_front_input)

        material_layout.addWidget(QLabel("ลอน:"))
        self.corrugate_c_type_combo = QComboBox()
        self.corrugate_c_type_combo.addItems(["C", "E"]) # Added "None" option
        self.corrugate_c_type_combo.setCurrentText("C")
        material_layout.addWidget(self.corrugate_c_type_combo)
        self.corrugate_c_material_input = QComboBox() # Changed to QComboBox
        self.corrugate_c_material_input.setPlaceholderText("เลือกลอน")
        material_layout.addWidget(self.corrugate_c_material_input)

        
        material_layout.addWidget(QLabel("แผ่นกลาง:"))
        self.sheet_middle_input = QComboBox() # Changed to QComboBox
        material_layout.addWidget(self.sheet_middle_input)

        material_layout.addWidget(QLabel("ลอน:"))
        self.corrugate_b_type_combo = QComboBox()
        self.corrugate_b_type_combo.addItems(["B", "E"]) # Added "None" option
        self.corrugate_b_type_combo.setCurrentText("B")
        material_layout.addWidget(self.corrugate_b_type_combo)
        self.corrugate_b_material_input = QComboBox() # Changed to QComboBox
        material_layout.addWidget(self.corrugate_b_material_input) 

        material_layout.addWidget(QLabel("แผ่นหลัง:"))
        self.sheet_back_input = QComboBox() # Changed to QComboBox
        material_layout.addWidget(self.sheet_back_input)
    
        # Connect material/width changes to stock update
        self.sheet_front_input.currentTextChanged.connect(self.update_length_based_on_stock) # Changed from textChanged to currentTextChanged
        self.sheet_middle_input.currentTextChanged.connect(self.update_length_based_on_stock)
        self.sheet_back_input.currentTextChanged.connect(self.update_length_based_on_stock)
        self.width_combo.currentTextChanged.connect(self.update_length_based_on_stock)

        self.corrugate_c_material_input.currentTextChanged.connect(self.update_length_based_on_stock)
        self.corrugate_b_material_input.currentTextChanged.connect(self.update_length_based_on_stock)

        layout.addLayout(material_layout)

        # เพิ่มช่องกรอกปริมาณม้วนกระดาษ พร้อมไอคอน info
        # info_layout = QHBoxLayout()
        # info_label = QLabel("ปริมาณม้วนกระดาษที่ใช้ได้สูงสุด (ม้วน):")
        # info_icon = QLabel()
        # info_icon.setPixmap(self.style().standardIcon(QApplication.style().SP_MessageBoxInformation).pixmap(16, 16))
        # info_icon.setToolTip(
        #     "ปริมาณม้วนกระดาษสูงสุดในสเปคนี้ที่ใช้ได้โดยไม่เกินม้วนอื่น (หน่วย: ม้วน)\n"
        #     "ระบบจะใช้ค่านี้เป็นขีดจำกัดในการคำนวณการตัดม้วนกระดาษ\n"
        # )
        # info_layout.addWidget(info_label)
        # info_layout.addWidget(info_icon)
        # info_layout.addStretch()
        # layout.addLayout(info_layout)

        # self.roll_qty = QLineEdit("")
        # self.update_length_based_on_stock() 
        # self.roll_qty.setPlaceholderText("ปริมาณม้วนกระดาษ (เมตร)")
        # self.roll_qty.setEnabled(False)
        # layout.addWidget(self.roll_qty)


        # เพิ่มช่องกรอกวันที่ด้วย QDateEdit (บังคับให้แสดงเลขอารบิก)
        layout.addWidget(QLabel("วันที่เริ่มต้น (YYYY-MM-DD):"))
        self.start_date_input = QDateEdit()
        self.start_date_input.setDisplayFormat("yyyy-MM-dd")
        self.start_date_input.setCalendarPopup(True)
        self.start_date_input.setDate(QDate(2023, 1, 1)) # ตั้งค่าเริ่มต้นเป็นวันที่ 1 มกราคม 2023
        self.start_date_input.setLocale(QLocale(QLocale.English, QLocale.Thailand))
        layout.addWidget(self.start_date_input)

        layout.addWidget(QLabel("วันที่สิ้นสุด (YYYY-MM-DD):"))
        self.end_date_input = QDateEdit()
        self.end_date_input.setDisplayFormat("yyyy-MM-dd")
        self.end_date_input.setCalendarPopup(True)
        self.end_date_input.setDate(QDate.currentDate().addYears(1)) # ตั้งค่าเริ่มต้นเป็น 1 ปีข้างหน้า
        self.end_date_input.setLocale(QLocale(QLocale.English, QLocale.Thailand))
        layout.addWidget(self.end_date_input)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)  # ตั้งค่าระยะ 0-100%
        self.progress_bar.setTextVisible(True)  # แสดงข้อความ
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setFormat("กำลังรอการเริ่มต้น...") # ตั้งค่าข้อความเริ่มต้น
        layout.addWidget(self.progress_bar)
        
        # Run button
        self.run_button = QPushButton("เริ่มการคำนวณ")
        self.run_button.clicked.connect(self.run_calculation)
        layout.addWidget(self.run_button)
        
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
        self.result_table.setColumnCount(10)
        self.result_table.setHorizontalHeaderLabels([
            "ความกว้างม้วน", "หมายเลขออเดอร์", "ความกว้างออเดอร์", "จำนวนออก", "เศษเหลือ",
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
        # เชื่อมต่อสัญญาณ enterPressed ของตารางไปยังเมธอดที่แสดงป๊อปอัป
        self.result_table.enterPressed.connect(self.show_row_details_popup)
        # เชื่อมต่อสัญญาณ doubleClicked ของตารางไปยังเมธอดที่แสดงป๊อปอัป
        self.result_table.doubleClicked.connect(self.show_row_details_popup)

        self.setup_stock_manager()

    def closeEvent(self, event):
        """หยุดการทำงานของ worker threads อย่างถูกต้องเมื่อปิดโปรแกรม"""
        self.log_message("กำลังปิดโปรแกรม...")
        if hasattr(self, 'stock_manager'):
            self.stock_manager.stop()
        if hasattr(self, 'stock_thread'):
            self.stock_thread.quit()
            self.stock_thread.wait(5000) # รอสูงสุด 5 วินาที

        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.quit()
            self.worker.wait(5000)
        
        event.accept()

    def setup_stock_manager(self):
        """เริ่มต้นและเริ่มการทำงานของเธรดจัดการสต็อก"""
        stock_file_path = self.stock_file_path_input.text()
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
            self.ROLL_SPECS = new_roll_specs
            self.log_message("🔄 อัปเดตข้อมูลสต็อกเรียบร้อยแล้ว")

            # อัปเดต QComboBox ของความกว้างม้วนด้วยข้อมูลใหม่จากสต็อก
            current_width = self.width_combo.currentText()
            self.width_combo.blockSignals(True)
            self.width_combo.clear()
            if self.ROLL_SPECS:
                # เรียงลำดับความกว้างตามตัวเลข
                sorted_widths = sorted(self.ROLL_SPECS.keys(), key=lambda x: int(re.sub(r'\D', '', x) or 0))
                self.width_combo.addItems(sorted_widths)
                if current_width in sorted_widths:
                    self.width_combo.setCurrentText(current_width)
                elif sorted_widths:
                    self.width_combo.setCurrentIndex(0) # เลือกอันแรกถ้าอันเดิมไม่มี
            self.width_combo.blockSignals(False)
          
            # รีเฟรชองค์ประกอบ UI ที่ขึ้นอยู่กับข้อมูลสต็อก
            self.update_length_based_on_stock()

    def update_length_based_on_stock(self):
        """Update material combobox and roll quantity based on stock, avoiding recursion."""
        sender = self.sender()

        material_combos = [
            self.sheet_front_input,
            self.sheet_middle_input,
            self.sheet_back_input,
            self.corrugate_c_material_input,
            self.corrugate_b_material_input
        ]

        # Only update material lists if width changed or it's the initial call (sender is None)
        if sender is self.width_combo or sender is None:
            # Block signals to prevent recursive calls during programmatic updates
            for combo in material_combos:
                combo.blockSignals(True)

            # Preserve current selections to restore them after repopulating
            selections = [combo.currentText() for combo in material_combos]
            
            current_width = self.width_combo.currentText()
            materials = []
            if current_width and current_width in self.ROLL_SPECS:
                materials = list(self.ROLL_SPECS[current_width].keys())

            # Add a "None" option (empty string) at the beginning
            materials.insert(0, "")
            # materials.append("")
          

            for i, combo in enumerate(material_combos):
                combo.clear()
                if materials:
                    combo.addItems(materials)
                    # Restore previous selection if it's valid, otherwise default to first item
                    if selections[i] in materials:
                        combo.setCurrentText(selections[i])
                    elif combo.count() > 0:
                        combo.setCurrentIndex(0)
            
            # Re-enable signals
            for combo in material_combos:
                combo.blockSignals(False)

        # --- This part always runs to update the length based on current selections ---
        current_width = self.width_combo.currentText()
        selected_materials = [combo.currentText().strip() for combo in material_combos]
        
        # นับจำนวนครั้งที่ใช้วัสดุแต่ละชนิด
        material_counts = collections.Counter(m for m in selected_materials if m)
        unique_materials = list(material_counts.keys())

        effective_lengths = []
        rolls_used = 0 
        if unique_materials and current_width:
            if current_width not in self.ROLL_SPECS:
                self.ROLL_SPECS[current_width] = {}
            stock_data_for_width = self.ROLL_SPECS[current_width]
            for material in unique_materials:
                material_rolls = stock_data_for_width.get(material)
                if material_rolls:
                    # 1. หาความยาวม้วนที่น้อยที่สุดของวัสดุชนิดนั้น แล้วคูณด้วยจำนวนม้วน
                    min_length = min(roll['length'] for roll in material_rolls.values())

                    # 2. หารด้วยจำนวนครั้งที่ใช้วัสดุนั้น
                    usage_count = material_counts[material]
                    # เอาม้วนที่มีมาหารกับจำนวนครั้งที่ใช้วัสดุว่าใช้ได้กี่รอบ
                    roll_used = floor(len(material_rolls)/usage_count)
                    rolls_used += roll_used
  
                    # คำนวณความยาวที่ใช้ได้จริง
                    effective_length = min_length * roll_used
                    effective_lengths.append(effective_length)
        
        if effective_lengths:
            # 3. ใช้ค่าที่น้อยที่สุดเป็นขีดจำกัด
            min_effective_length = min(effective_lengths)
            self.calculated_length = int(min_effective_length)
            # self.roll_qty.setText(str(int(rolls_used))) # แสดงเป็นจำนวนเต็ม
        else:
            self.calculated_length = 0
            # self.roll_qty.clear()
            
    def select_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "เลือกไฟล์ออเดอร์",
            "",
            "CSV Files (*.csv);;All Files (*)",
            options=options
        )
        if file_path:
            self.file_path_input.setText(file_path)

    def select_stock_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "เลือกไฟล์สต็อก",
            "",
            "CSV Files (*.csv);;All Files (*)",
            options=options
        )
        if file_path:
            self.stock_file_path_input.setText(file_path)
            # แจ้ง Stock Manager เกี่ยวกับเส้นทางไฟล์ใหม่
            if hasattr(self, 'stock_manager'):
                self.stock_manager.set_file_path(file_path)

    def log_message(self, message: str):
        self.log_display.append(message)
        self.log_display.verticalScrollBar().setValue(
            self.log_display.verticalScrollBar().maximum()
        )
        
    def run_calculation(self):
        try:
            width = int(self.width_combo.currentText())
            length = self.calculated_length
            
            # อ่านค่าเส้นทางไฟล์จาก UI
            file_path = self.file_path_input.text().strip() or "order2024.csv"

            # อ่านค่าจาก QDateEdit และแปลงเป็นสตริง
            start_date = self.start_date_input.date().toString("yyyy-MM-dd") 
            end_date = self.end_date_input.date().toString("yyyy-MM-dd")
            
            # แปลงตัวเลขไทยเป็นอารบิก
            start_date = convert_thai_digits_to_arabic(start_date).strip() or None
            end_date = convert_thai_digits_to_arabic(end_date).strip() or None

            # อ่านค่าจากช่องกรอกวัสดุ
            front_material = self.sheet_front_input.currentText().strip() or None # Changed from .text() to .currentText()
            middle_material = self.sheet_middle_input.currentText().strip() or None
            back_material = self.sheet_back_input.currentText().strip() or None
            corrugate_c_type = self.corrugate_c_type_combo.currentText().strip() or None
            corrugate_c_material_name = self.corrugate_c_material_input.currentText().strip() or None
            corrugate_b_type = self.corrugate_b_type_combo.currentText().strip() or None
            corrugate_b_material_name = self.corrugate_b_material_input.currentText().strip() or None
            
            roll_specs_copy = copy.deepcopy(self.ROLL_SPECS)
            
            self.log_display.clear()
            self.result_table.setRowCount(0)
            self.log_message("⚙️ กำลังเริ่มการคำนวณ...")
            self.run_button.setEnabled(False)
            self.progress_bar.setValue(0) # รีเซ็ตโปรเกรสบาร์
            self.progress_bar.setFormat("กำลังประมวลผล...")
            
            # Create worker thread
            self.worker = WorkerThread(
                width, length, start_date, end_date, file_path,
                front_material, 
                corrugate_c_type, corrugate_c_material_name,
                middle_material, 
                corrugate_b_type, corrugate_b_material_name,
                back_material,
                roll_specs_copy
            )
            self.worker.update_signal.connect(self.log_message)
            self.worker.progress_updated.connect(self.update_progress_bar)
            self.worker.finished.connect(self.complete_calculation)
            self.worker.error_signal.connect(self.handle_error)
            self.worker.start()
            
        except ValueError:
            QMessageBox.warning(self, "ข้อผิดพลาด", "⚠️ ไม่สามารถแปลงค่าความกว้างเป็นตัวเลขได้ โปรดตรวจสอบข้อมูลสต็อก")
            self.log_message("⚠️ ไม่สามารถแปลงค่าความกว้างเป็นตัวเลขได้ โปรดตรวจสอบข้อมูลสต็อก")
            self.run_button.setEnabled(True) # เปิดปุ่มกลับมา
            self.progress_bar.setFormat("เกิดข้อผิดพลาด!")

    def update_progress_bar(self, value: int, message: str):
        """อัปเดตแถบความคืบหน้าและข้อความ"""
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(f"{message} ({value}%)")
        if value == 100:
            self.run_button.setEnabled(True)

    def complete_calculation(self, results):
        self.run_button.setEnabled(True)
        if not results:
            self.progress_bar.setFormat("ไม่พบม้วนที่เหมาะสม")
            self.log_message("❌ ไม่พบม้วนที่เหมาะสม กรุณาเปลี่ยนหน้าม้วน")
            QMessageBox.warning(self, "ไม่พบม้วนที่เหมาะสม", "❌ ไม่พบม้วนที่เหมาะสม กรุณาเปลี่ยนหน้าม้วน")
            self.results_data = []
            self.result_table.setRowCount(0)
            return

        self.progress_bar.setFormat("เสร็จสิ้น!")
        self.log_message(f"✅ เสร็จสิ้นการคำนวณสำหรับม้วน {len(results)} ครั้ง")
        QMessageBox.information(self, "เสร็จสิ้น", f"✅ เสร็จสิ้นการคำนวณสำหรับม้วน {len(results)} ครั้ง")

        # เก็บผลลัพธ์ทั้งหมดไว้ในตัวแปรของคลาส
        self.results_data = results

        # แสดงผลลัพธ์ในตารางด้วยการตั้งค่า Flag เพื่อให้สามารถเลือกคัดลอกได้
        self.result_table.setRowCount(len(results))
        for row_idx, result in enumerate(results):
            # ตรวจสอบว่ามีม้วนที่ไม่เหมาะสมหรือไม่ เพื่อเปลี่ยนสีแถว
            has_no_suitable_roll = False
            roll_info_keys = ['front_roll_info', 'c_roll_info', 'middle_roll_info', 'b_roll_info', 'back_roll_info']
            for key in roll_info_keys:
                if "ไม่มี" in result.get(key, ''):
                    has_no_suitable_roll = True
                    break

            for col_idx, value in enumerate([
                str(result.get('roll_w', '')),
                str(result.get('order_number', '')),
                f"{result.get('order_w', ''):.4f}",
                str(result.get('cuts', '')),
                f"{result.get('trim', ''):.2f}",
                f"{result.get('order_l', ''):.4f}",
                f"{result.get('order_dmd', '')}",
                str(result.get('die_cut', '')),
                f"{result.get('order_qty', '')}",
                f"{result.get('order_qty', '')/result.get('cuts', ''):.2f}",
                # f"{result.get('demand_per_cut', ''):.4f}",
                # f"{result.get('rem_roll_l', ''):.4f}"
            ]):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                if has_no_suitable_roll:
                    item.setBackground(QColor(255, 224, 224)) # สีแดงอ่อน

                self.result_table.setItem(row_idx, col_idx, item)
        self.result_table.resizeColumnsToContents()

    def handle_error(self, error_message: str):
        self.run_button.setEnabled(True)
        self.progress_bar.setFormat("เกิดข้อผิดพลาด!")
        self.log_message(f"❌ {error_message}")
        QMessageBox.critical(self, "ข้อผิดพลาด", f"❌ เกิดข้อผิดพลาดในการคำนวณ:\n{error_message}")

    def show_row_details_popup(self):
        """
        แสดงป๊อปอัปพร้อมรายละเอียดของแถวที่เลือกในตาราง
        """
        selected_rows = self.result_table.selectedIndexes()
        if not selected_rows:
            return

        row_index = selected_rows[0].row() # รับดัชนีของแถวที่เลือก
        
        # พยายามดึงข้อมูลผลลัพธ์เต็มจาก self.results_data
        try:
            result = self.results_data[row_index]
        except (IndexError, TypeError):
            # Fallback to table data only if full results are not available
            result = {}
        
        # แสดงข้อมูลเดิมจากตาราง
        details = []
        for col_idx in range(self.result_table.columnCount()):
            item = self.result_table.item(row_index, col_idx)
            if item:
                header = self.result_table.horizontalHeaderItem(col_idx).text()
                details.append(f"{header}: {item.text()}")

        # เพิ่มข้อมูลประเภททับเส้นและชนิดส่วนประกอบ
        type_details = []
        if result.get('type'):
            type_details.append(f"ประเภททับเส้น: {result['type']}")
        if result.get('component_type'):
            type_details.append(f"ชนิดส่วนประกอบ: {result['component_type']}")
        
        if type_details:
            details.append("\n📌 ข้อมูลประเภท:")
            details.extend(type_details)

        # เพิ่มข้อมูลวัสดุแบบมีเงื่อนไขเฉพาะที่มีค่าเท่านั้น
        material_details = []

        # Determine a common divisor based on corrugate types in result for front/middle/back materials
        c_type = result.get('c_type', '')
        b_type = result.get('b_type', '')

        type_demand = 1.0 # Default divisor if no specific C or B corrugate
        if c_type == 'C':
            type_demand = 1.45
        elif b_type == 'B':
            type_demand = 1.35
        elif c_type == 'E' or b_type == 'E':
            type_demand = 1.25

        if result.get('front'):
            front_material = result.get('front')
            front_value = result.get('demand_per_cut', 0) / type_demand
            roll_info_str = result.get('front_roll_info', '')
            material_details.append(f"แผ่นหน้า: {front_material} = {front_value:.2f} {roll_info_str}")
            
        if result.get('c') and c_type == 'C':
            c_material = result.get('c')
            c_value = result.get('demand_per_cut', 0)
            roll_info_str = result.get('c_roll_info', '')
            material_details.append(f"ลอน C: {c_material} = {c_value:.2f} {roll_info_str}")
        elif result.get('c') and c_type == 'E':
            c_material = result.get('c')
            # Removed redundant 'front_value' calculation
            if b_type == 'B':
                c_e_value = result.get('demand_per_cut', 0) / 1.35 * 1.25    # This might need a specific E-type B factor if it exists
            else:
                c_e_value = result.get('demand_per_cut', 0)
            roll_info_str = result.get('c_roll_info', '')
            material_details.append(f"ลอน E: {c_material} = {c_e_value:.2f} {roll_info_str}")

        if result.get('middle'):
            middle_material = result.get('middle')
            middle_value = result.get('demand_per_cut', 0) / type_demand
            roll_info_str = result.get('middle_roll_info', '')
            material_details.append(f"แผ่นกลาง: {middle_material} = {middle_value:.2f} {roll_info_str}")
           
        #B is value, if B exist and corrugate_b_type is 'B' or 'E', calculate accordingly
        if result.get('b') and b_type == 'B':
            b_material = result.get('b')
            if c_type == 'C':
                b_value = (result.get('demand_per_cut', 0) / 1.45) * 1.35
            else:
                b_value = result.get('demand_per_cut', 0)
            roll_info_str = result.get('b_roll_info', '')
            material_details.append(f"ลอน B: {b_material} = {b_value:.2f} {roll_info_str}")
        elif result.get('b') and b_type == 'E':
            b_material = result.get('b')
            if c_type == 'C':
                b_e_value = (result.get('demand_per_cut', 0) / 1.45) * 1.25
            else:
                b_e_value = result.get('demand_per_cut', 0)
            roll_info_str = result.get('b_roll_info', '')
            material_details.append(f"ลอน E: {b_material} = {b_e_value:.2f} {roll_info_str}")

        if result.get('back'):
            back_material = result.get('back')
            back_value = result.get('demand_per_cut', 0) / type_demand
            roll_info_str = result.get('back_roll_info', '')
            material_details.append(f"แผ่นหลัง: {back_material} = {back_value:.2f} {roll_info_str}")
        
        if material_details:
            details.append("\n⚙️ ข้อมูลแผ่นและลอน:")
            details.extend(material_details)

        detail_message = "\n".join(details)
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(detail_message)
        msg_box.setWindowTitle("รายละเอียดผลลัพธ์การตัด")
        msg_box.setTextInteractionFlags(Qt.TextSelectableByMouse)
        msg_box.exec_()

def convert_thai_digits_to_arabic(text: str) -> str:
    """Convert Thai digits to Arabic digits"""
    thai_digits = "๐๑๒๓๔๕๖๗๘๙"
    arabic_digits = "0123456789"
    translation_table = str.maketrans(thai_digits, arabic_digits)
    return text.translate(translation_table)

if __name__ == "__main__":
    # ตั้งค่า environment สำหรับภาษาไทยบน Windows
    if sys.platform == "win32":
        os.environ["QT_QPA_PLATFORM"] = "windows:fontengine=freetype"
        os.environ["PYTHONIOENCODING"] = "utf-8"
    
    # ตั้งค่า encoding สำหรับแอปพลิเคชัน
    QTextCodec.setCodecForLocale(QTextCodec.codecForName("UTF-8"))
    
    # แก้ไขตรงนี้: ใช้ QLocale.setDefault() แทน app.setLocale()
    thai_locale = QLocale(QLocale.Thai, QLocale.Thailand)
    QLocale.setDefault(thai_locale)
    
    app = QApplication(sys.argv)
    app.setFont(QFont('Tahoma', 9))  # ตั้งค่าฟอนต์ภาษาไทย
    
    window = CuttingOptimizerUI()
    window.show()
    sys.exit(app.exec_())

