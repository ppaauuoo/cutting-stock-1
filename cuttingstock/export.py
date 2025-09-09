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

            # Width change highlight format
            width_change_format = workbook.add_format({
                'bg_color': '#FFE6CC',
                'border': 1
            })

            # Material change highlight format
            material_change_format = workbook.add_format({
                'bg_color': '#E2EFDA',
                'border': 1
            })

            # New roll highlight format
            new_roll_format = workbook.add_format({
                'bg_color': '#FFF2CC',
                'border': 1
            })

            # Write headers
            for col, header in enumerate(self.headers + self.detail_headers):
                worksheet.write(0, col, header, header_format)

            # Write data rows with cell-specific coloring
            row_idx = 1
            previous_width = None
            previous_group_id = 'XXX'
            previous_materials = {
                'front': None, 'c': None, 'middle': None, 'b': None, 'back': None
            }

            for result in self.results_data:
                # Check if roll width changed
                current_width = result.get('roll_w')
                width_changed = current_width != previous_width
                previous_width = current_width

                # Check if this is a singular order (no group_id key at all)
                is_singular_order = 'group_id' not in result
                
                # For group orders, track group changes to determine if we should show roll width
                if not is_singular_order:
                    current_group_id = result.get('group_id', '')
                    if previous_group_id == 'XXX':
                        previous_group_id = current_group_id
                    group_changed = current_group_id != previous_group_id
                    previous_group_id = current_group_id
                    
                    # Show roll width for first order in group or when width changes
                    show_roll_width = group_changed or width_changed
                else:
                    # For singular orders, always show roll width
                    show_roll_width = True

                # Build row data with roll width visibility logic
                row_data = self._build_main_row_data_with_visibility(result, show_roll_width)
                detail_data = self._build_detail_row_data(result)
                full_row_data = row_data + detail_data

                # Write row with default format first
                for col, value in enumerate(full_row_data):
                    worksheet.write(row_idx, col, value, color_formats[0])

                # Apply special coloring to roll width cell (column 0) based on width changes
                # Only color if the cell actually shows a value (not blank)
                if width_changed and current_width and show_roll_width:
                    worksheet.write(row_idx, 0, row_data[0], width_change_format)

                # Apply coloring to material columns when values change
                material_columns = {
                    'front': len(self.headers) + 0,      # แผ่นหน้า (วัสดุ)
                    'c': len(self.headers) + 3,           # ลอน C (วัสดุ)
                    'middle': len(self.headers) + 6,       # แผ่นกลาง (วัสดุ)
                    'b': len(self.headers) + 9,           # ลอน B (วัสดุ)
                    'back': len(self.headers) + 12        # แผ่นหลัง (วัสดุ)
                }

                for material_type, col_idx in material_columns.items():
                    current_material = result.get(material_type, '')
                    if current_material != previous_materials[material_type]:
                        worksheet.write(row_idx, col_idx, current_material, material_change_format)
                        previous_materials[material_type] = current_material

                # Apply coloring to roll ID columns when value is "เปิดม้วนใหม่"
                roll_id_columns = {
                    'front_roll_info': len(self.headers) + 2,    # แผ่นหน้า (ID ม้วน)
                    'c_roll_info': len(self.headers) + 5,         # ลอน C (ID ม้วน)
                    'middle_roll_info': len(self.headers) + 8,    # แผ่นกลาง (ID ม้วน)
                    'b_roll_info': len(self.headers) + 11,        # ลอน B (ID ม้วน)
                    'back_roll_info': len(self.headers) + 14      # แผ่นหลัง (ID ม้วน)
                }

                for roll_info_key, col_idx in roll_id_columns.items():
                    roll_info = result.get(roll_info_key, '')
                    if 'เปิดม้วนใหม่' in str(roll_info):
                        worksheet.write(row_idx, col_idx, roll_info, new_roll_format)

                row_idx += 1

            # Auto-adjust column widths
            for col in range(len(self.headers + self.detail_headers)):
                worksheet.set_column(col, col, 15)

            workbook.close()
            return True

        except Exception as e:
            print(f"Error exporting to XLSX: {e}")
            return False

    def _build_main_row_data_with_visibility(self, result: Dict[str, Any], show_roll_width: bool) -> List[str]:
        """Build the main row data for export with visibility control"""
        cuts = result.get('cuts')
        order_qty = result.get('order_qty')
        demand_per_cut_val = ""
        if cuts is not None and cuts > 0 and order_qty is not None:
            demand_per_cut_val = f"{order_qty / cuts:.2f}"
        else:
            demand_per_cut_val = "N/A"

        # Show roll width only if specified (for group orders) or always for singular orders
        roll_width_value = str(result.get('roll_w', '')) if show_roll_width else ''
        
        row_data = [roll_width_value]

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
