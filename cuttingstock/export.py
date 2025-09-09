import csv
import re
from typing import List, Dict, Any
import polars as pl

try:
    import xlsxwriter
    XLSX_SUPPORT = True
except ImportError:
    XLSX_SUPPORT = False


class ExportManager:
    """Handles exporting cutting results to various formats"""
    
    def __init__(self, results_data: List[Dict[str, Any]], headers: List[str]):
        self.results_data = results_data
        self.headers = headers
        self.detail_headers = [
            "แผ่นหน้า (วัสดุ)", "แผ่นหน้า (ใช้)", "แผ่นหน้า (ID ม้วน)",
            "ลอน C (วัสดุ)", "ลอน C (ใช้)", "ลอน C (ID ม้วน)",
            "แผ่นกลาง (วัสดุ)", "แผ่นกลาง (ใช้)", "แผ่นกลาง (ID ม้วน)",
            "ลอน B (วัสดุ)", "ลอน B (ใช้)", "ลอน B (ID ม้วน)",
            "แผ่นหลัง (วัสดุ)", "แผ่นหลัง (ใช้)", "แผ่นหลัง (ID ม้วน)",
            "ประเภททับเส้น", "ชนิดส่วนประกอบ"
        ]
    
    def export_to_csv(self, file_path: str) -> bool:
        """Export results to CSV format"""
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csv_file:
                writer = csv.writer(csv_file)
                
                # Write headers
                writer.writerow(self.headers + self.detail_headers)
                
                # Write data rows
                group_id = '000'
                for result in self.results_data:
                    row_data = self._build_main_row_data(result, group_id)
                    group_id = result.get('group_id', '')
                    
                    detail_data = self._build_detail_row_data(result)
                    writer.writerow(row_data + detail_data)
                    
            return True
            
        except Exception as e:
            print(f"Error exporting to CSV: {e}")
            return False
    
    def export_to_xlsx(self, file_path: str) -> bool:
        """Export results to XLSX format with row coloring"""
        if not XLSX_SUPPORT:
            print("xlsxwriter not available. Please install: pip install xlsxwriter")
            return False
            
        try:
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet('Cutting Results')
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D7E4BC',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            # Define alternating row colors
            color_formats = [
                workbook.add_format({'bg_color': '#FFFFFF', 'border': 1}),
                workbook.add_format({'bg_color': '#F2F2F2', 'border': 1})
            ]
            
            # Group change highlight format
            group_change_format = workbook.add_format({
                'bg_color': '#FFE6CC',
                'border': 1
            })
            
            # Write headers
            for col, header in enumerate(self.headers + self.detail_headers):
                worksheet.write(0, col, header, header_format)
            
            # Write data rows with coloring
            row_idx = 1
            group_id = '000'
            previous_front_value = None
            
            for result in self.results_data:
                # Check if group changed
                current_group_id = result.get('group_id', '')
                group_changed = current_group_id != group_id
                group_id = current_group_id
                
                # Check if front value changed for coloring
                current_front = result.get('front', '')
                front_changed = current_front != previous_front_value
                previous_front_value = current_front
                
                # Build row data
                row_data = self._build_main_row_data(result, group_id if group_changed else '')
                detail_data = self._build_detail_row_data(result)
                full_row_data = row_data + detail_data
                
                # Determine row format based on changes
                if group_changed:
                    row_format = group_change_format
                else:
                    # Alternate colors based on front value changes
                    color_index = 0 if not front_changed else 1
                    row_format = color_formats[color_index]
                
                # Write row with formatting
                for col, value in enumerate(full_row_data):
                    worksheet.write(row_idx, col, value, row_format)
                
                row_idx += 1
            
            # Auto-adjust column widths
            for col in range(len(self.headers + self.detail_headers)):
                worksheet.set_column(col, col, 15)
            
            workbook.close()
            return True
            
        except Exception as e:
            print(f"Error exporting to XLSX: {e}")
            return False
    
    def _build_main_row_data(self, result: Dict[str, Any], group_id: str) -> List[str]:
        """Build the main row data for export"""
        cuts = result.get('cuts')
        order_qty = result.get('order_qty')
        demand_per_cut_val = ""
        if cuts is not None and cuts > 0 and order_qty is not None:
            demand_per_cut_val = f"{order_qty / cuts:.2f}"
        else:
            demand_per_cut_val = "N/A"

        row_data = [
            str(result.get('roll_w', '')) if group_id != result.get('group_id', '') else '',
        ]

        detail = [
            str(result.get('order_number', '')),
            str(result.get('due_date', '')),
            str(result.get('component_type', '')),
            f"{result.get('order_w', ''):.4f}",
            str(result.get('cuts', '')),
            f"{result.get('trim', ''):.2f}",
            f"{result.get('order_l', ''):.4f}",
            f"{result.get('order_dmd', '')}",
            str(result.get('die_cut', '')),
            f"{result.get('order_qty', '')}",
            demand_per_cut_val,
        ]
        row_data.extend(detail)
        
        return row_data
    
    def _build_detail_row_data(self, result: Dict[str, Any]) -> List[str]:
        """Build the detail row data for export"""
        c_type = result.get('c_type', '')
        b_type = result.get('b_type', '')

        type_demand = 1.0
        if c_type == 'C':
            type_demand = 1.45
        elif b_type == 'B':
            type_demand = 1.35
        elif c_type == 'E' or b_type == 'E':
            type_demand = 1.25

        # Front sheet
        front_str, front_value, front_roll_info = "", "", ""
        if result.get('front'):
            front_material = result.get('front')
            demand_per_cut = result.get('demand_per_cut', 0)
            if type_demand > 0:
                front_value = f"{demand_per_cut / type_demand:.2f}"
            front_str = front_material
            front_roll_info = self._format_roll_usage_for_csv(result.get('front_roll_info', ''))

        # C sheet
        c_str, c_value, c_roll_info = "", "", ""
        if result.get('c'):
            c_material = result.get('c')
            demand_per_cut = result.get('demand_per_cut', 0)
            if c_type == 'C':
                c_value = f"{demand_per_cut:.2f}"
            elif c_type == 'E':
                if b_type == 'B':
                    c_value = f"{(demand_per_cut / 1.35 * 1.25):.2f}"
                else:
                    c_value = f"{demand_per_cut:.2f}"
            c_str = c_material
            c_roll_info = self._format_roll_usage_for_csv(result.get('c_roll_info', ''))

        # Middle sheet
        middle_str, middle_value, middle_roll_info = "", "", ""
        if result.get('middle'):
            middle_material = result.get('middle')
            demand_per_cut = result.get('demand_per_cut', 0)
            if type_demand > 0:
                middle_value = f"{demand_per_cut / type_demand:.2f}"
            middle_str = middle_material
            middle_roll_info = self._format_roll_usage_for_csv(result.get('middle_roll_info', ''))

        # B sheet
        b_str, b_value, b_roll_info = "", "", ""
        if result.get('b'):
            b_material = result.get('b')
            demand_per_cut = result.get('demand_per_cut', 0)
            if b_type == 'B':
                if c_type == 'C':
                    b_value = f"{(demand_per_cut / 1.45 * 1.35):.2f}"
                else:
                    b_value = f"{demand_per_cut:.2f}"
            elif b_type == 'E':
                if c_type == 'C':
                    b_value = f"{(demand_per_cut / 1.45 * 1.25):.2f}"
                else:
                    b_value = f"{demand_per_cut:.2f}"
            b_str = b_material
            b_roll_info = self._format_roll_usage_for_csv(result.get('b_roll_info', ''))

        # Back sheet
        back_str, back_value, back_roll_info = "", "", ""
        if result.get('back'):
            back_material = result.get('back')
            demand_per_cut = result.get('demand_per_cut', 0)
            if type_demand > 0:
                back_value = f"{demand_per_cut / type_demand:.2f}"
            back_str = back_material
            back_roll_info = self._format_roll_usage_for_csv(result.get('back_roll_info', ''))

        detail_data = [
            front_str, front_value, front_roll_info,
            c_str, c_value, c_roll_info,
            middle_str, middle_value, middle_roll_info,
            b_str, b_value, b_roll_info,
            back_str, back_value, back_roll_info,
            result.get('type', ''),
            result.get('component_type', '')
        ]

        return detail_data
    
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