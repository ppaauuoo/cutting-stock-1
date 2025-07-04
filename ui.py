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
    finished = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def __init__(self, width, length, start_date, end_date, file_path,
                 front_material, c_material, middle_material, b_material, back_material, # เพิ่มพารามิเตอร์วัสดุ
                 parent=None):
        super().__init__(parent)
        self.width = width
        self.length = length
        self.start_date = start_date
        self.end_date = end_date
        self.file_path = file_path
        self.front_material = front_material # เก็บค่าพารามิเตอร์วัสดุ
        self.c_material = c_material
        self.middle_material = middle_material
        self.b_material = b_material
        self.back_material = back_material

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        def progress_callback(message: str):
            self.update_signal.emit(message)
        
        try:
            results = loop.run_until_complete(
                main.main_algorithm(
                    roll_width=self.width,
                    roll_length=self.length,
                    progress_callback=progress_callback,
                    start_date=self.start_date,
                    end_date=self.end_date,
                    file_path=self.file_path,
                    front=self.front_material, # ส่งพารามิเตอร์วัสดุ
                    C=self.c_material,
                    middle=self.middle_material,
                    B=self.b_material,
                    back=self.back_material,
                )
            )
            self.finished.emit(results)
        except Exception as e:
            self.error_signal.emit(f"Error: {str(e)}")
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

        # Input fields
        # กำหนดค่าความกว้างม้วนกระดาษที่ใช้ได้
        ROLL_PAPER = [66, 68, 70, 73, 74, 75, 79, 82, 85, 88, 91, 93, 95, 97]

        layout.addWidget(QLabel("ความกว้างม้วนกระดาษ (inch):")) 
        self.width_combo = QComboBox()
        self.width_combo.addItems([str(w) for w in ROLL_PAPER])
        self.width_combo.setCurrentText("75")
        layout.addWidget(self.width_combo)
                # เพิ่มช่องกรอกข้อมูลแผ่นและลอนในแถวเดียวกัน
        material_layout = QHBoxLayout()

        material_layout.addWidget(QLabel("แผ่นหน้า:"))
        self.sheet_front_input = QLineEdit('CM127')
        self.sheet_front_input.setPlaceholderText("แผ่นหน้า")
        material_layout.addWidget(self.sheet_front_input)

        material_layout.addWidget(QLabel("ลอน C:"))
        self.corrugate_c_input = QLineEdit('CM127')
        self.corrugate_c_input.setPlaceholderText("ลอน C")
        material_layout.addWidget(self.corrugate_c_input)

        material_layout.addWidget(QLabel("แผ่นกลาง:"))
        self.sheet_middle_input = QLineEdit('CM127')
        self.sheet_middle_input.setPlaceholderText("แผ่นกลาง")
        material_layout.addWidget(self.sheet_middle_input)

        material_layout.addWidget(QLabel("ลอน B:"))
        self.corrugate_b_input = QLineEdit('CM127')
        self.corrugate_b_input.setPlaceholderText("ลอน B")
        material_layout.addWidget(self.corrugate_b_input)

        material_layout.addWidget(QLabel("แผ่นหลัง:"))
        self.sheet_back_input = QLineEdit('CM127')
        self.sheet_back_input.setPlaceholderText("แผ่นหลัง")
        material_layout.addWidget(self.sheet_back_input)
        
        layout.addLayout(material_layout)

        layout.addWidget(QLabel("ความยาวม้วนกระดาษ (m):")) 
        self.length_input = QLineEdit("111175")
        self.length_input.setPlaceholderText("เช่น 111175")
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
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setFormat("กำลังรอการเริ่มต้น...")
        layout.addWidget(self.progress_bar)
        
        # Run button
        self.run_button = QPushButton("เริ่มการคำนวณ")
        self.run_button.clicked.connect(self.run_calculation)
        layout.addWidget(self.run_button)
        
        # Log display (Collapsible)
        self.log_group_box = QGroupBox("บันทึกการทำงาน:")
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
            "ความกว้างม้วน", "หมายเลขออเดอร์", "ความกว้างออเดอร์", 
            "ความยาวออเดอร์", "ปริมาณออเดอร์", "จำนวนตัด", "Trim Waste", "ความยาวใช้ไป", "ความยาวคงเหลือ"
        ])
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        # ตั้งค่าตารางให้สามารถเลือกและคัดลอกได้
        self.result_table.setSelectionMode(QTableWidget.SingleSelection)
        self.result_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.result_table)
        
        self.setCentralWidget(central_widget)
        
        # เชื่อมต่อสัญญาณ enterPressed ของตารางไปยังเมธอดที่แสดงป๊อปอัป
        self.result_table.enterPressed.connect(self.show_row_details_popup)

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
            front_material = self.sheet_front_input.text().strip() or None
            c_material = self.corrugate_c_input.text().strip() or None
            middle_material = self.sheet_middle_input.text().strip() or None
            b_material = self.corrugate_b_input.text().strip() or None
            back_material = self.sheet_back_input.text().strip() or None
            
            self.log_display.clear()
            self.result_table.setRowCount(0)
            self.log_message("⚙️ กำลังเริ่มการคำนวณ...")
            self.run_button.setEnabled(False)
            self.progress_bar.setFormat("กำลังประมวลผล...")
            
            # Create worker thread
            self.worker = WorkerThread(
                width, length, start_date, end_date, file_path,
                front_material, c_material, middle_material, b_material, back_material # ส่งค่าวัสดุ
            )
            self.worker.update_signal.connect(self.log_message)
            self.worker.finished.connect(self.complete_calculation)
            self.worker.error_signal.connect(self.handle_error)
            self.worker.start()
            
        except ValueError:
            QMessageBox.warning(self, "ข้อผิดพลาด", "⚠️ โปรดป้อนค่าความยาวเป็นตัวเลขเท่านั้น!")
            self.log_message("⚠️ โปรดป้อนค่าความยาวเป็นตัวเลขเท่านั้น!")

    def complete_calculation(self, results):
        self.run_button.setEnabled(True)
        self.progress_bar.setFormat("เสร็จสิ้น!")
        self.log_message(f"✅ เสร็จสิ้นการคำนวณสำหรับม้วน {len(results)} ครั้ง")
        QMessageBox.information(self, "เสร็จสิ้น", f"✅ เสร็จสิ้นการคำนวณสำหรับม้วน {len(results)} ครั้ง")

        # แสดงผลลัพธ์ในตารางด้วยการตั้งค่า Flag เพื่อให้สามารถเลือกคัดลอกได้

        # แสดงผลลัพธ์ในตารางด้วยการตั้งค่า Flag เพื่อให้สามารถเลือกคัดลอกได้
        self.result_table.setRowCount(len(results))
        for row_idx, result in enumerate(results):
            for col_idx, value in enumerate([
                str(result.get('roll width', '')),
                str(result.get('order_number', '')),
                f"{result.get('selected_order_width', ''):.2f}",
                f"{result.get('selected_order_length', ''):.2f}",
                f"{result.get('selected_order_quantity', ''):.2f}",
               str(result.get('num_cuts_z', '')),
                f"{result.get('calculated_trim', ''):.2f}",
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

        # รับชื่อหัวข้อคอลัมน์เพื่อนำไปแสดง
        headers = [self.result_table.horizontalHeaderItem(col).text() for col in range(self.result_table.columnCount())]
        
        details = []
        for col_idx in range(self.result_table.columnCount()):
            item = self.result_table.item(row_index, col_idx)
            if item:
                details.append(f"{headers[col_idx]}: {item.text()}")
        
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
