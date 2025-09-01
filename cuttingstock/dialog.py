
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QGroupBox, QLabel, QMessageBox, QVBoxLayout


class MaterialSubstitutionDialog(QDialog):
    def __init__(self, parent, width: str, available_materials: list, out_of_stock_material: str, material_specs: dict, known_out_of_stock: list = None):
        super().__init__(parent)
        self.setWindowTitle("แก้ไข/เปลี่ยนวัสดุ")
        self.original_specs = material_specs.copy()
        self.out_of_stock_material = out_of_stock_material
        self.combos = {}

        layout = QVBoxLayout(self)

        spec_str = ", ".join(f"{k.title().replace('_', ' ')}: {v}" for k, v in self.original_specs.items() if v)
        spec_label = QLabel(f"<b>Current Spec:</b><br>{spec_str}")
        spec_label.setTextFormat(Qt.RichText)
        layout.addWidget(spec_label)

        all_known_oos_materials = set()
        if known_out_of_stock:
            for item in known_out_of_stock:
                if isinstance(item, tuple) and len(item) == 2:  # It's a (width, material) tuple
                    all_known_oos_materials.add(item[1])
                elif isinstance(item, str):  # It's a material name
                    all_known_oos_materials.add(item)

        message = f"วัสดุ '{out_of_stock_material}' สำหรับความกว้าง {width} นิ้วไม่พอ"
        other_oos_to_display = sorted(list(all_known_oos_materials - {out_of_stock_material}))
        if other_oos_to_display:
            message += f"\nวัสดุต่อไปนี้ก็อาจไม่พอ: {', '.join(other_oos_to_display)}"
        message += "\n\nคุณสามารถเลือกวัสดุทดแทนสำหรับแต่ละรายการได้:"

        self.message_label = QLabel(message)
        layout.addWidget(self.message_label)

        spec_group = QGroupBox("เลือกวัสดุ:")
        spec_layout = QFormLayout()

        # Define a consistent order for materials
        material_types_ordered = ['front', 'c', 'middle', 'b', 'back']
        for key in material_types_ordered:
            value = self.original_specs.get(key)
            if value: # Only show rows for materials that are part of the spec
                combo = QComboBox()
                combo.addItems(available_materials)
                try:
                    # Find by exact match first, trimming any whitespace
                    index = combo.findText(str(value).strip(), Qt.MatchFixedString)
                    if index != -1:
                        combo.setCurrentIndex(index)
                except (ValueError, AttributeError):
                    pass # Keep default if not found

                self.combos[key] = combo
                label = QLabel(f"{key.replace('_', ' ').title()}:")
                if value == self.out_of_stock_material:
                    label.setStyleSheet("font-weight: bold; color: red;")
                spec_layout.addRow(label, combo)

        spec_group.setLayout(spec_layout)
        layout.addWidget(spec_group)

        materials_to_disable = all_known_oos_materials | {out_of_stock_material}

        for material_to_disable in materials_to_disable:
            for combo in self.combos.values():
                try:
                    index = combo.findText(material_to_disable)
                    if index != -1:
                        item = combo.model().item(index)
                        if item:
                            item.setEnabled(False)
                except (ValueError, AttributeError):
                    pass

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Abort, Qt.Horizontal, self)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        abort_button = self.buttons.button(QDialogButtonBox.Abort)
        if abort_button:
            abort_button.clicked.connect(lambda: self.done(2)) # Use custom code 2 for abort
        layout.addWidget(self.buttons)

    def get_selected_specs(self) -> dict:
        """Returns the full new material spec dictionary."""
        new_specs = self.original_specs.copy()
        for key, combo in self.combos.items():
            new_specs[key] = combo.currentText()
        return new_specs

    def accept(self):
        """Overrides accept to check if the out-of-stock material was changed."""
        new_specs = self.get_selected_specs()

        # Find which spec key corresponds to the out-of-stock material
        oos_spec_key = None
        for key, value in self.original_specs.items():
            if value == self.out_of_stock_material:
                oos_spec_key = key
                break

        if oos_spec_key and new_specs.get(oos_spec_key) == self.out_of_stock_material:
            QMessageBox.warning(self, "ยังคงเลือกวัสดุที่หมด",
                                f"วัสดุ '{self.out_of_stock_material}' ไม่พอใช้\nกรุณาเลือกวัสดุทดแทนสำหรับรายการนี้ หรือกดยกเลิก")
            return # Do not close the dialog

        super().accept()
