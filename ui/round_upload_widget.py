import os
import sqlite3
import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,QInputDialog,
    QFileDialog, QComboBox, QTableWidget, QTableWidgetItem, QMessageBox
)

DB_NAME = "mtech_offers.db"
from ui.rounds_manager import upload_round_decisions
def _sanitize_col_name(name: str) -> str:
    """Sanitize SQL column names (letters, numbers, underscore only)."""
    return "".join(c if c.isalnum() or c == "_" else "_" for c in str(name)).lower()
class SingleFileUpload(QWidget):
    """Handles single Excel file upload and column mapping."""
    def __init__(self, title, required_map):
        super().__init__()
        self.title = title
        self.required_map = required_map   # list of tuples: (db_col, human_label)

        self.file_path = None
        self.df = None
        self.col_map = {}    # {"mtech_app_no": "Excel Col Name", ...}

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.title_label = QLabel(f"{self.title}: <font color='red'>No file uploaded</font>")
        self.layout.addWidget(self.title_label)

        row = QHBoxLayout()
        self.upload_btn = QPushButton("Select File")
        self.upload_btn.clicked.connect(self.select_file)
        row.addWidget(self.upload_btn)

        self.get_cols_btn = QPushButton("Show Column Mapping")
        self.get_cols_btn.setEnabled(False)
        self.get_cols_btn.clicked.connect(self.show_column_match_table)
        row.addWidget(self.get_cols_btn)

        self.layout.addLayout(row)
        self.table_widget = None
        
    def reset_widget(self):
        """Reset file selection and column mapping."""
        self.file_path = None
        self.df = None
        self.col_map = {}

        self.title_label.setText(f"{self.title}: <font color='red'>No file uploaded</font>")
        self.get_cols_btn.setEnabled(False)

        if self.table_widget:
            self.layout.removeWidget(self.table_widget)
            self.table_widget.deleteLater()
            self.table_widget = None
        
    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return

        self.file_path = path

        try:
            self.df = pd.read_excel(self.file_path)
            self.title_label.setText(f"{self.title}: <font color='green'>{os.path.basename(path)}</font>")
            self.get_cols_btn.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "File Error", f"Could not read Excel file:\n{e}")
            self.file_path = None
            self.df = None
            self.get_cols_btn.setEnabled(False)

    def show_column_match_table(self):
        if self.table_widget:
            self.layout.removeWidget(self.table_widget)
            self.table_widget.deleteLater()
            self.table_widget = None

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(2)
        self.table_widget.setRowCount(len(self.required_map))
        self.table_widget.setHorizontalHeaderLabels(["DB Column (Required)", "Excel Column (Select)"])
        self.table_widget.verticalHeader().setVisible(False)

        for i, (db_col, human_label) in enumerate(self.required_map):
            self.table_widget.setItem(i, 0, QTableWidgetItem(human_label))

            combo = QComboBox()
            combo.addItems(self.df.columns.tolist())

            # --- AUTO-MATCH HERE ---
            auto = self.best_match(db_col, self.df.columns.tolist())
            if auto:
                combo.setCurrentText(auto)
                self.set_col_map(db_col, auto)
            else:
                # fallback: first column
                combo.setCurrentText(self.df.columns[0])
                self.set_col_map(db_col, self.df.columns[0])

            combo.currentTextChanged.connect(lambda val, db=db_col: self.set_col_map(db, val))
            self.table_widget.setCellWidget(i, 1, combo)


        self.layout.addWidget(self.table_widget)

    def set_col_map(self, db_col, excel_col):
        self.col_map[db_col] = excel_col
    def get_file_path(self):
        return self.file_path
    @staticmethod
    def best_match(required, columns):
        req_norm = required.lower().replace(" ", "_")
        cols_norm = [c.lower().replace(" ", "_") for c in columns]

        # 1. Exact match
        if req_norm in cols_norm:
            return columns[cols_norm.index(req_norm)]

        # 2. Fuzzy match
        import difflib
        close = difflib.get_close_matches(req_norm, cols_norm, n=1, cutoff=0.55)
        if close:
            return columns[cols_norm.index(close[0])]

        return None


    def get_mapped_dataframe(self):
        """Return df with DB column names based on selected mapping."""
        if self.df is None or len(self.col_map) == 0:
            return None
        df = self.df.copy()
        rename_dict = {excel: db for db, excel in self.col_map.items()}  
        df = df.rename(columns=rename_dict)

        # Filter only required DB columns
        df = df[list(self.col_map.keys())]

        return df

class RoundUploadWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        # Required column mappings
        required_map_goa = [
            ("mtech_app_no", "MTech Application Number"),
            ("applicant_decision", "Applicant Decision"),
        ]

        required_map_other = [
            ("mtech_app_no", "MTech Application Number"),
            ("other_institute_decision", "Other Institute Decision"),
        ]

        required_map_consolidated = [
            ("coap_reg_id", "COAP Registration ID"),
            ("applicant_decision", "Applicant Decision"),
        ]

        # Three uploaders
        self.goa_widget = SingleFileUpload("IIT Goa Candidate Decision Report", required_map_goa)
        self.other_widget = SingleFileUpload("Other IIT Freeze/Accept Report", required_map_other)
        self.cons_widget = SingleFileUpload("Consolidated All-Institutes Report", required_map_consolidated)

        self.layout.addWidget(self.goa_widget)
        self.layout.addWidget(self.other_widget)
        self.layout.addWidget(self.cons_widget)
