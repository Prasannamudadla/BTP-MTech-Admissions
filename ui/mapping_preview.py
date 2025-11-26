# # database/db_manager.py: (Assume this exists)
# # REQUIRED_DB_COLUMNS = [...] # You should add this list here or define it globally.

# from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout,QWidget, QLabel, QComboBox, QPushButton, QMessageBox
# from PySide6.QtCore import Qt
# # Assuming you define REQUIRED_DB_COLUMNS somewhere accessible, 
# # for now, let's hardcode a few for demonstration:
# REQUIRED_DB_FIELDS = [
#     "COAP", "App_no", "Email", "Full_Name", "Ews", "Gender", "Category","Pwd", 
#     "MaxGATEScore_3yrs", "HSSC_per", "SSC_per", "Degree_CGPA_8th", 
# ]
# # REQUIRED_DB_FIELDS = [
# #     "Si_NO", "App_no", "Email", "Full_Name", "Adm_cat", "Pwd", "Ews", "Gender", 
# #     "Category", "COAP", "GATE22RollNo", "GATE22Rank", "GATE22Score", "GATE22Disc", 
# #     "GATE21RollNo", "GATE21Rank", "GATE21Score", "GATE21Disc", "GATE20RollNo", 
# #     "GATE20Rank", "GATE20Score", "GATE20Disc", "MaxGATEScore_3yrs", "GATE_Roll_num", 
# #     "HSSC_board", "HSSC_date", "HSSC_per", "SSC_board", "SSC_date", "SSC_per", 
# #     "Degree_Qualification", "Degree_PassingDate", "Degree_Branch", "Degree_OtherBranch", 
# #     "Degree_Institute", "Degree_CGPA_7th", "Degree_CGPA_8th", "Degree_Per_7th", 
# #     "Degree_Per_8th", "ExtraColumn"
# # ]

# class MappingDialog(QDialog):
#     def __init__(self, excel_headers, parent=None):
#         super().__init__(parent)
#         self.setWindowTitle("Final Field Details - Column Mapping")
#         self.setMinimumWidth(600)
        
#         self.excel_headers = ['-- IGNORE --'] + excel_headers
#         self.mapping_widgets = {} # Stores {DB_Field: QComboBox}
#         self.final_mapping = {}

#         self.setup_ui()
        
#     def setup_ui(self):
#         main_layout = QVBoxLayout(self)
        
#         # Instruction Label
#         instruction_label = QLabel("Select the correct fields from Excel Sheet in dropdown against the database fields!")
#         main_layout.addWidget(instruction_label, alignment=Qt.AlignmentFlag.AlignCenter)
        
#         # --- Mapping Grid (Simulated 3-column layout) ---
#         mapping_area = QWidget()
#         mapping_layout = QHBoxLayout(mapping_area)
        
#         # Create 3 vertical columns for better screen usage (as shown in your image)
#         columns = [QVBoxLayout() for _ in range(3)]
        
#         for i, db_field in enumerate(REQUIRED_DB_FIELDS):
#             # Container for one field mapping
#             field_widget = QWidget()
#             field_layout = QVBoxLayout(field_widget)
            
#             # 1. Label (DB Field Name)
#             label = QLabel(f"<b>{db_field}</b>")
#             field_layout.addWidget(label)
            
#             # 2. Dropdown (Excel Column Names)
#             combo = QComboBox()
#             combo.addItems(self.excel_headers)
#             self.mapping_widgets[db_field] = combo
#             field_layout.addWidget(combo)
            
#             # Add the field widget to one of the 3 columns
#             columns[i % 3].addWidget(field_widget)

#         for col in columns:
#             col.addStretch() # Push fields to the top
#             mapping_layout.addLayout(col)

#         main_layout.addWidget(mapping_area)
        
#         # --- Action Buttons ---
#         button_layout = QHBoxLayout()
        
#         self.save_btn = QPushButton("SAVE TO DATABASE")
#         self.save_btn.clicked.connect(self.accept_mapping)
#         button_layout.addWidget(self.save_btn)
        
#         self.reset_btn = QPushButton("Reset Data")
#         self.reset_btn.clicked.connect(self.reset_form)
#         button_layout.addWidget(self.reset_btn)

#         main_layout.addLayout(button_layout)
        
#     def accept_mapping(self):
#         """Processes the user's selections and prepares the final mapping."""
#         # 1. Validation (Optional but Recommended)
#         used_excel_fields = []
        
#         for db_field, combo in self.mapping_widgets.items():
#             selected_excel_field = combo.currentText()
            
#             if selected_excel_field == '-- IGNORE --':
#                 # Allow user to ignore fields if they aren't in the Excel file
#                 continue
            
#             if selected_excel_field in used_excel_fields:
#                 QMessageBox.warning(self, "Mapping Error", 
#                                     f"The Excel column '{selected_excel_field}' is mapped to multiple DB fields. Please fix.")
#                 return

#             self.final_mapping[selected_excel_field] = db_field
#             used_excel_fields.append(selected_excel_field)
            
#         if not self.final_mapping:
#             QMessageBox.critical(self, "No Fields Selected", "Please map at least one field before saving.")
#             return

#         # Success - close dialog and return QDialog.Accepted
#         self.accept()

#     def get_mapping(self):
#         """Returns the {Excel_Header: DB_Field} dictionary."""
#         return self.final_mapping

#     def reset_form(self):
#         """Resets all dropdowns to the default ignore state."""
#         for combo in self.mapping_widgets.values():
#             combo.setCurrentIndex(0) # Assumes '-- IGNORE --' is index 0
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QHBoxLayout, QLabel, QComboBox
)

class MappingPreviewDialog(QDialog):
    def __init__(self, mapping, src_columns, required_targets=None, parent=None):
        """
        mapping: dict (db_column -> suggested excel source)
        src_columns: Excel columns list
        required_targets: list of DB columns to show in UI
        """
        super().__init__(parent)
        self.setWindowTitle("Confirm Column Mapping")
        self.resize(700, 400)

        self.all_mapping = mapping.copy()
        self.src_columns = list(src_columns)

        # If not provided, show all mapping (fallback)
        if required_targets is None:
            required_targets = list(mapping.keys())

        # Only show these required DB fields in UI
        self.display_targets = required_targets  

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Select correct Excel column for each required database field:"))

        # UI table (only required rows)
        self.table = QTableWidget(len(self.display_targets), 2)
        self.table.setHorizontalHeaderLabels(["DB Column", "Mapped Column"])

        for i, tgt in enumerate(self.display_targets):
            # Pre-filled suggestion
            suggested_src = self.all_mapping.get(tgt)

            # DB column name
            self.table.setItem(i, 0, QTableWidgetItem(tgt))

            # Dropdown of Excel columns
            combo = QComboBox()
            combo.addItem("")  # Allow empty
            for col in self.src_columns:
                combo.addItem(col)

            # Auto-select suggested column
            if suggested_src:
                combo.setCurrentText(suggested_src)

            self.table.setCellWidget(i, 1, combo)

        layout.addWidget(self.table)

        # OK / Cancel buttons
        btns = QHBoxLayout()
        ok = QPushButton("OK")
        cancel = QPushButton("Cancel")
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        btns.addWidget(ok)
        btns.addWidget(cancel)
        layout.addLayout(btns)

    def get_final_mapping(self):
        """
        Return mapping for ALL DB columns:
        - required targets: use user selection
        - non-required: keep autosuggested values
        """
        final = self.all_mapping.copy()

        for r in range(self.table.rowCount()):
            tgt = self.table.item(r, 0).text()   # DB field
            combo = self.table.cellWidget(r, 1)
            choice = combo.currentText().strip()
            final[tgt] = choice if choice else None

        return final