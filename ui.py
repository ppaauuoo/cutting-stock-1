import asyncio
import os
import re
import sys

from PyQt5.QtCore import (
    QDate,
    QLocale,
    Qt,
    QTextCodec,
    QThread,
    pyqtSignal,
)
from PyQt5.QtGui import QFont
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
        
        # Input fields
        # กำหนดค่าความกว้างม้วนกระดาษที่ใช้ได้
        ROLL_PAPER = [66, 68, 70, 73, 74, 75, 79, 82, 85, 88, 91, 93, 95, 97]
    
        self.ROLL_SPECS = {
                            '66': 
                            {
                            'CM127': 20000, 
                            'KB120':9000, 
                            'CM100':8000
                            },
                            '75': 
                            {
                            'CM127': 1920
                            },
                        }  # Mockup stock specs

        # Constants for material calculations
        self.E_FACTOR = 1.25
        self.C_FACTOR = 1.45
        self.B_FACTOR = 1.35

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        
        # เพิ่มส่วนเลือกไฟล์
        layout.addWidget(QLabel("เลือกไฟล์ออเดอร์:"))
        
        file_layout = QHBoxLayout()
        self.file_path_input = QLineEdit("order2024.csv")
        self.file_path_input.setPlaceholderText("เลือกไฟล์ CSV...")
        file_layout.addWidget(self.file_path_input)
        
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.select_file)
        file_layout.addWidget(browse_button)
        
        layout.addLayout(file_layout)


        layout.addWidget(QLabel("ความกว้างม้วนกระดาษ (inch):")) 
        self.width_combo = QComboBox()
        self.width_combo.addItems(self.ROLL_SPECS.keys())
        self.width_combo.setCurrentText("ความกว้างม้วนกระดาษ (inch)")
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

        # เพิ่มช่องกรอกความยาวม้วนกระดาษ พร้อมไอคอน info
        info_layout = QHBoxLayout()
        info_label = QLabel("ม้วนกระดาษที่ใช้ได้สูงสุด (m):")
        info_icon = QLabel()
        info_icon.setPixmap(self.style().standardIcon(QApplication.style().SP_MessageBoxInformation).pixmap(16, 16))
        info_icon.setToolTip(
            "ความยาวม้วนกระดาษสูงสุดในสเปคนี้ที่ใช้ได้โดยไม่เกินม้วนอื่น (หน่วย: เมตร)\n"
            "ระบบจะใช้ค่านี้เป็นขีดจำกัดในการคำนวณการตัดม้วนกระดาษ\n"
        )
        info_layout.addWidget(info_label)
        info_layout.addWidget(info_icon)
        info_layout.addStretch()
        layout.addLayout(info_layout)

        self.length_input = QLineEdit("")
        self.update_length_based_on_stock() 
        self.length_input.setPlaceholderText("ความยาวม้วนกระดาษ (เมตร)")
        self.length_input.setEnabled(False)
        layout.addWidget(self.length_input)

        # เพิ่มช่องกรอกวันที่ด้วย QDateEdit (บังคับให้แสดงเลขอารบิก)
        layout.addWidget(QLabel("วันที่เริ่มต้น (YYYY-MM-DD):"))
        self.start_date_input = QDateEdit()
        self.start_date_input.setDisplayFormat("yyyy-MM-dd")
        self.start_date_input.setCalendarPopup(True)
        self.start_date_input.setDate(QDate.currentDate())
        self.start_date_input.setLocale(QLocale(QLocale.English, QLocale.Thailand))
        layout.addWidget(self.start_date_input)

        layout.addWidget(QLabel("วันที่สิ้นสุด (YYYY-MM-DD):"))
        self.end_date_input = QDateEdit()
        self.end_date_input.setDisplayFormat("yyyy-MM-dd")
        self.end_date_input.setCalendarPopup(True)
        self.end_date_input.setDate(QDate.currentDate().addDays(3))
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
        self.result_table.setColumnCount(9)
        self.result_table.setHorizontalHeaderLabels([
            "ความกว้างม้วน", "หมายเลขออเดอร์", "ความกว้างออเดอร์", "จำนวนตัด", "เศษเหลือ",
            "ความยาวออเดอร์", "ปริมาณออเดอร์",  "กระดาษที่ใช้", "กระดาษคงเหลือ"
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

    def update_length_based_on_stock(self):
        """Update length_input and material combobox based on stock, avoiding recursion."""
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
        unique_materials = list(set(m for m in selected_materials if m))

        stocks = []
        if unique_materials and current_width and current_width in self.ROLL_SPECS:
            stock_data = self.ROLL_SPECS[current_width]
            for material in unique_materials:
                stock = stock_data.get(material)
                if stock is not None:
                    stocks.append(stock)
        
        if stocks:
            min_stock = min(stocks)
            self.length_input.setText(str(min_stock))
        else:
            self.length_input.setText("0")
            
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

    def log_message(self, message: str):
        self.log_display.append(message)
        self.log_display.verticalScrollBar().setValue(
            self.log_display.verticalScrollBar().maximum()
        )
        
    def run_calculation(self):
        try:
            width = int(self.width_combo.currentText())
            length = int(self.length_input.text())
            
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
                back_material
            )
            self.worker.update_signal.connect(self.log_message)
            self.worker.progress_updated.connect(self.update_progress_bar)
            self.worker.finished.connect(self.complete_calculation)
            self.worker.error_signal.connect(self.handle_error)
            self.worker.start()
            
        except ValueError:
            QMessageBox.warning(self, "ข้อผิดพลาด", "⚠️ โปรดป้อนค่าความยาวเป็นตัวเลขเท่านั้น!")
            self.log_message("⚠️ โปรดป้อนค่าความยาวเป็นตัวเลขเท่านั้น!")
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
            for col_idx, value in enumerate([
                str(result.get('roll width', '')),
                str(result.get('order_number', '')),
                f"{result.get('selected_order_width', ''):.2f}",
                str(result.get('num_cuts_z', '')),
                f"{result.get('calculated_trim', ''):.2f}",
                f"{result.get('selected_order_length', ''):.2f}",
                f"{result.get('selected_order_quantity', ''):.2f}",
                f"{result.get('demand', ''):.2f}",
                f"{result.get('roll length', ''):.2f}"
            ]):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
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
            type_demand = self.C_FACTOR
        elif b_type == 'B':
            type_demand = self.B_FACTOR
        elif c_type == 'E' or b_type == 'E':
            type_demand = self.E_FACTOR

        if result.get('front'):
            front_value = result.get('demand', 0) / type_demand
            material_details.append(f"แผ่นหน้า: {result.get('front')} = {front_value:.2f}") # Use .get() for consistency
            
        print(result.get('c'))
        # Corrected bitwise '&' to logical 'and'
        if result.get('c') and c_type == 'C':
            c_value = result.get('demand', 0)
            material_details.append(f"ลอน C: {result.get('c')} = {c_value:.2f}") # Use .get() for consistency
        elif result.get('c') and c_type == 'E':
            # Removed redundant 'front_value' calculation
            if b_type == 'B':
                c_e_value = result.get('demand', 0) / self.B_FACTOR * self.E_FACTOR    # This might need a specific E-type B factor if it exists
            else:
                c_e_value = result.get('demand', 0)
            material_details.append(f"ลอน E: {result.get('c')} = {c_e_value:.2f}") # Use .get() for consistency

        if result.get('middle'):
            middle_value = result.get('demand', 0) / type_demand
            material_details.append(f"แผ่นกลาง: {result.get('middle')} = {middle_value:.2f}") # Use .get() for consistency
           
        #B is value, if B exist and corrugate_b_type is 'B' or 'E', calculate accordingly
        if result.get('b') and b_type == 'B':
            if c_type == 'C':
                b_value = (result.get('demand', 0) / self.C_FACTOR) * self.B_FACTOR
            else:
                b_value = result.get('demand', 0)
            material_details.append(f"ลอน B: {result.get('b')} = {b_value:.2f}")
        elif result.get('b') and b_type == 'E':
            if c_type == 'C':
                b_e_value = (result.get('demand', 0) / self.C_FACTOR) * self.E_FACTOR
            else:
                b_e_value = result.get('demand', 0)
            material_details.append(f"ลอน E: {result.get('b')} = {b_e_value:.2f}")

        if result.get('back'):
            back_value = result.get('demand', 0) / type_demand
            material_details.append(f"แผ่นหลัง: {result.get('back')} = {back_value:.2f}") # Use .get() for consistency
        
        if material_details:
            details.append("\n⚙️ ข้อมูลแผ่นและลอน:")
            details.extend(material_details)

        detail_message = "\n".join(details)
        QMessageBox.information(self, "รายละเอียดผลลัพธ์การตัด", detail_message)

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

