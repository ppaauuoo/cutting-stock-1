import sys
import asyncio
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, 
    QProgressBar, QMessageBox, QComboBox, QTableWidget, QTableWidgetItem # เพิ่ม QComboBox, QTableWidget, QTableWidgetItem
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import main  # Import our modified main module

class WorkerThread(QThread):
    update_signal = pyqtSignal(str)
    finished = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def __init__(self, width, length, start_date, end_date): # เพิ่มพารามิเตอร์
        super().__init__()
        self.width = width
        self.length = length
        self.start_date = start_date # เก็บค่าวันที่
        self.end_date = end_date     # เก็บค่าวันที่

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
                    start_date=self.start_date, # ส่งค่าวันที่ไปยัง main_algorithm
                    end_date=self.end_date      # ส่งค่าวันที่ไปยัง main_algorithm
                )
            )
            self.finished.emit(results)
        except Exception as e:
            self.error_signal.emit(f"Error: {str(e)}")
        finally:
            loop.close()

class CuttingOptimizerUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("กระดาษม้วนตัด Optimizer")
        self.setGeometry(100, 100, 800, 700) # ปรับขนาดหน้าต่างให้ใหญ่ขึ้น
        
        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        
        # Input fields
        # กำหนดค่าความกว้างม้วนกระดาษที่ใช้ได้
        ROLL_PAPER = [66, 68, 70, 73, 74, 75, 79, 82, 85, 88, 91, 93, 95, 97]

        layout.addWidget(QLabel("ความกว้างม้วนกระดาษ (inch):")) 
        self.width_combo = QComboBox()
        self.width_combo.addItems([str(w) for w in ROLL_PAPER]) # เพิ่มตัวเลือกจาก ROLL_PAPER
        self.width_combo.setCurrentText("75") # กำหนดค่าเริ่มต้น
        layout.addWidget(self.width_combo)
        
        layout.addWidget(QLabel("ความยาวม้วนกระดาษ (m):")) 
        self.length_input = QLineEdit("111175")
        self.length_input.setPlaceholderText("เช่น 111175")
        layout.addWidget(self.length_input)

        # เพิ่มช่องกรอกวันที่
        layout.addWidget(QLabel("วันที่เริ่มต้น (YYYY-MM-DD):"))
        self.start_date_input = QLineEdit()
        self.start_date_input.setPlaceholderText("เช่น 2024-01-01")
        layout.addWidget(self.start_date_input)
        
        layout.addWidget(QLabel("วันที่สิ้นสุด (YYYY-MM-DD):"))
        self.end_date_input = QLineEdit()
        self.end_date_input.setPlaceholderText("เช่น 2024-12-31")
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
        
        # Log display
        layout.addWidget(QLabel("บันทึกการทำงาน:"))
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        layout.addWidget(self.log_display)

        # เพิ่มตารางแสดงผล
        layout.addWidget(QLabel("ผลลัพธ์การตัด:"))
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(8)
        self.result_table.setHorizontalHeaderLabels([
            "ความกว้างม้วน", "หมายเลขออเดอร์", "ความกว้างออเดอร์", 
            "ความยาวออเดอร์", "จำนวนตัด", "Trim Waste", "ความยาวใช้ไป", "ความยาวคงเหลือ"
        ])
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers) # ทำให้แก้ไขไม่ได้
        layout.addWidget(self.result_table)
        
        self.setCentralWidget(central_widget)
        
    def log_message(self, message: str):
        self.log_display.append(message)
        self.log_display.verticalScrollBar().setValue(
            self.log_display.verticalScrollBar().maximum()
        )
        
    def run_calculation(self):
        try:
            width = int(self.width_combo.currentText()) # ดึงค่าจาก QComboBox
            length = int(self.length_input.text())
            
            # อ่านค่าวันที่จาก UI
            start_date = self.start_date_input.text().strip()
            end_date = self.end_date_input.text().strip()

            # กำหนดให้เป็น None ถ้าเป็นสตริงว่าง
            start_date = start_date if start_date else None
            end_date = end_date if end_date else None
            
            self.log_display.clear()
            self.result_table.setRowCount(0) # ล้างตารางก่อนเริ่มคำนวณใหม่
            self.log_message("⚙️ กำลังเริ่มการคำนวณ...")
            self.run_button.setEnabled(False)
            self.progress_bar.setFormat("กำลังประมวลผล...")
            
            # Create worker thread
            self.worker = WorkerThread(width, length, start_date, end_date) # ส่งค่าวันที่ไปยัง worker thread
            self.worker.update_signal.connect(self.log_message)
            self.worker.finished.connect(self.complete_calculation)
            self.worker.error_signal.connect(self.handle_error)
            self.worker.start()
            
        except ValueError:
            QMessageBox.warning(self, "ข้อผิดพลาด", "⚠️ โปรดป้อนค่าความยาวเป็นตัวเลขเท่านั้น!") # แก้ไขข้อความเตือน
            self.log_message("⚠️ โปรดป้อนค่าความยาวเป็นตัวเลขเท่านั้น!") # แก้ไขข้อความเตือน

    def complete_calculation(self, results):
        self.run_button.setEnabled(True)
        self.progress_bar.setFormat("เสร็จสิ้น!")
        self.log_message(f"✅ เสร็จสิ้นการคำนวณสำหรับม้วน {len(results)} ครั้ง")
        QMessageBox.information(self, "เสร็จสิ้น", f"✅ เสร็จสิ้นการคำนวณสำหรับม้วน {len(results)} ครั้ง")

        # แสดงผลลัพธ์ในตาราง
        self.result_table.setRowCount(len(results))
        for row_idx, result in enumerate(results):
            self.result_table.setItem(row_idx, 0, QTableWidgetItem(str(result.get('roll width', ''))))
            self.result_table.setItem(row_idx, 1, QTableWidgetItem(str(result.get('order_number', ''))))
            self.result_table.setItem(row_idx, 2, QTableWidgetItem(f"{result.get('selected_order_width', ''):.2f} inch"))
            self.result_table.setItem(row_idx, 3, QTableWidgetItem(f"{result.get('selected_order_length', ''):.2f} m"))
            self.result_table.setItem(row_idx, 4, QTableWidgetItem(str(result.get('num_cuts_z', ''))))
            self.result_table.setItem(row_idx, 5, QTableWidgetItem(f"{result.get('calculated_trim', ''):.2f} inch"))
            self.result_table.setItem(row_idx, 6, QTableWidgetItem(f"{result.get('demand', ''):.2f} m"))
            self.result_table.setItem(row_idx, 7, QTableWidgetItem(f"{result.get('roll length', ''):.2f} m")) # ความยาวคงเหลือของม้วน
        self.result_table.resizeColumnsToContents() # ปรับขนาดคอลัมน์ให้พอดีเนื้อหา

    def handle_error(self, error_message: str):
        self.run_button.setEnabled(True)
        self.progress_bar.setFormat("เกิดข้อผิดพลาด!")
        self.log_message(f"❌ {error_message}")
        QMessageBox.critical(self, "ข้อผิดพลาด", f"❌ เกิดข้อผิดพลาดในการคำนวณ:\n{error_message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CuttingOptimizerUI()
    window.show()
    sys.exit(app.exec_())
