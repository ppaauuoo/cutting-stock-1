
import threading
from PyQt5.QtCore import QThread, pyqtSignal

import asyncio
import re
from cuttingstock.core import main_algorithm
from cuttingstock.material import OutOfStockError


class WorkerThread(QThread):
    update_signal = pyqtSignal(str)
    progress_updated = pyqtSignal(int, str)
    calculation_succeeded = pyqtSignal(list)
    error_signal = pyqtSignal(str)
    out_of_stock_signal = pyqtSignal(dict)

    def __init__(self, width, length, start_date, end_date, file_path,
                 front_material,
                 corrugate_c_type, corrugate_c_material_name,
                 middle_material,
                 corrugate_b_type, corrugate_b_material_name,
                 back_material,
                 roll_specs,
                 processed_orders,
                 material_substitutions,
                 selected_factory,
                 parent=None):
        super().__init__(parent)
        self._wait_for_input_event = threading.Event()
        self._user_choice = None
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
        self.processed_orders = processed_orders
        self.material_substitutions = material_substitutions
        self.current_iteration_step = 0
        self.selected_factory = selected_factory

    def set_user_choice(self, choice):
        """Called from the UI thread to provide the user's choice."""
        self._user_choice = choice
        self._wait_for_input_event.set()

    def out_of_stock_handler(self, e: OutOfStockError):
        """
        This handler is called from within main_algorithm in the worker thread.
        It signals the UI and blocks until the user makes a choice.
        """
        self._wait_for_input_event.clear()
        self.out_of_stock_signal.emit({
            "width": e.width,
            "material": e.material,
            "required_length": e.required_length,
            "material_specs": e.material_specs,
            "known_out_of_stock": e.known_out_of_stock,
        })
        self._wait_for_input_event.wait()  # Block until set_user_choice is called
        return self._user_choice

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

        def progress_callback(message: str):
            if self.isInterruptionRequested():
                # Raise an exception to break out of the blocking call
                raise InterruptedError("Calculation was interrupted.")

            self.update_signal.emit(message)
            # ส่งสัญญาณพร้อมเปอร์เซ็นต์ความคืบหน้าประมาณการ
            if "กำลังเริ่มการคำนวณ" in message:
                self.progress_updated.emit(5, message)
            elif "โหลดและจัดเรียงข้อมูลเรียบร้อย" in message:
                self.progress_updated.emit(20, message)
            elif "Iteration" in message:
                # พยายามดึงตัวเลขการวนซ้ำทั้งหมด (X/Y)
                match = re.search(r'Iteration (\d+)(?:/| of )(\d+)', message)
                if match:
                    current_iter = int(match.group(1))
                    total_iters = int(match.group(2))
                    if total_iters > 0:
                        # คำนวณเปอร์เซ็นต์ความคืบหน้าในช่วง 50-95%
                        progress_percentage = 50 + (current_iter / total_iters) * 45
                        self.progress_updated.emit(int(progress_percentage), message)
                    else:
                        # หากไม่มีตัวเลขรวมหรือเป็น 0 ให้ใช้การเพิ่มค่าทีละน้อย
                        self.current_iteration_step += 1
                        estimated_progress = min(95, 50 + self.current_iteration_step) # เพิ่มทีละ 1%
                        self.progress_updated.emit(estimated_progress, message)
                else:
                    # หากไม่พบรูปแบบตัวเลข ให้เพิ่มค่าทีละน้อย
                    self.current_iteration_step += 1
                    estimated_progress = min(95, 50 + self.current_iteration_step) # เพิ่มทีละ 1%
                    self.progress_updated.emit(estimated_progress, message)
            elif "บันทึกผลลัพธ์ลงฐานข้อมูลเรียบร้อย" in message:
                self.progress_updated.emit(95, message)

        try:
            results = loop.run_until_complete(
                main_algorithm(
                    roll_width=self.width,
                    roll_length=self.length,
                    progress_callback=progress_callback,
                    out_of_stock_handler=self.out_of_stock_handler,
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
                    processed_orders=self.processed_orders,
                    material_substitutions=self.material_substitutions,
                    selected_factory=self.selected_factory
                )
            )
            if not self.isInterruptionRequested():
                self.progress_updated.emit(100, "✅ เสร็จสิ้น")  # สัญญาณเสร็จสมบูรณ์
                self.calculation_succeeded.emit(results)
        except InterruptedError:
            self.update_signal.emit("⏹️ การคำนวณถูกหยุดโดยผู้ใช้")
        except Exception as e:
            if not self.isInterruptionRequested():
                self.error_signal.emit(f"Error: {str(e)}")
                self.progress_updated.emit(0, "❌ เกิดข้อผิดพลาด!") # รีเซ็ตโปรเกรสบาร์เมื่อเกิดข้อผิดพลาด
        finally:
            loop.close()
