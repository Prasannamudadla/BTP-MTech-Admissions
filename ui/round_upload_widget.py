import os
import sqlite3
import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QComboBox, QTableWidget, QTableWidgetItem, QMessageBox
)

DB_NAME = "mtech_offers.db"

def _sanitize_col_name(name: str) -> str:
    """Sanitize SQL column names (letters, numbers, underscore only)."""
    return "".join(c if c.isalnum() or c == "_" else "_" for c in str(name)).lower()


class SingleFileUpload(QWidget):
    """Handles single Excel file upload, column mapping, and DB save."""
    def __init__(self, title, required_map, table_name_fn):
        super().__init__()
        self.title = title
        self.required_map = required_map  # list of tuples (db_col, human_label)
        self.table_name_fn = table_name_fn

        self.file_path = None
        self.df = None
        self.col_map = {}

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.title_label = QLabel(f"{self.title}: <font color='red'>No file uploaded</font>")
        self.layout.addWidget(self.title_label)

        row = QHBoxLayout()
        self.upload_btn = QPushButton("Select File")
        self.upload_btn.clicked.connect(self.select_file)
        row.addWidget(self.upload_btn)

        self.get_cols_btn = QPushButton("Get Column Names")
        self.get_cols_btn.setEnabled(False)
        self.get_cols_btn.clicked.connect(self.show_column_match_table)
        row.addWidget(self.get_cols_btn)

        self.save_btn = QPushButton("Save to DB")
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self.save_to_db)
        row.addWidget(self.save_btn)

        self.layout.addLayout(row)
        self.table_widget = None

    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        self.file_path = path
        self.df = pd.read_excel(self.file_path)
        self.title_label.setText(f"{self.title}: <font color='green'>{os.path.basename(path)}</font>")
        self.get_cols_btn.setEnabled(True)

    def show_column_match_table(self):
        if self.table_widget:
            self.layout.removeWidget(self.table_widget)
            self.table_widget.deleteLater()
            self.table_widget = None
            self.col_map = {}

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(2)
        self.table_widget.setRowCount(len(self.required_map))
        self.table_widget.setHorizontalHeaderLabels(["DB Column", "Excel Column"])
        self.table_widget.verticalHeader().setVisible(False)

        for i, (db_col, human_label) in enumerate(self.required_map):
            self.table_widget.setItem(i, 0, QTableWidgetItem(human_label))
            combo = QComboBox()
            combo.addItems(self.df.columns.tolist())
            combo.currentTextChanged.connect(lambda val, i=i, db=db_col: self.set_col_map(db, val))
            self.table_widget.setCellWidget(i, 1, combo)
            # preselect first column
            self.set_col_map(db_col, self.df.columns[0])

        self.layout.addWidget(self.table_widget)
        self.save_btn.setEnabled(True)

    def set_col_map(self, db_col, val):
        self.col_map[db_col] = val

    def save_to_db(self, round_no=None):
    # ✅ Default round number fallback
        if not round_no:
            try:
                # Try to fetch from parent if available
                if hasattr(self.parent(), "get_current_round"):
                    round_no = int(self.parent().get_current_round())
                else:
                    round_no = 1  # Default to Round 1
            except Exception:
                round_no = 1

        # ✅ Correct DataFrame check
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "Error", "No file uploaded or file is empty")
            return
        if not self.col_map:
            QMessageBox.warning(self, "Error", "No columns selected")
            return

        # ✅ Build table name with correct round number
        table_name = self.table_name_fn(round_no)
        print(f"[DEBUG] Saving data to table: {table_name}")

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        # Create table
        cols_sql = ", ".join([f"{_sanitize_col_name(c)} TEXT" for c in self.col_map.keys()])
        cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({cols_sql})")

        # Prepare insert
        insert_cols = [_sanitize_col_name(c) for c in self.col_map.keys()]
        placeholders = ", ".join(["?"] * len(insert_cols))
        rows = [[row[self.col_map[c]] for c in self.col_map.keys()] for _, row in self.df.iterrows()]

        cursor.executemany(
            f"INSERT INTO {table_name} ({', '.join(insert_cols)}) VALUES ({placeholders})",
            rows
        )
        conn.commit()
        conn.close()
        QMessageBox.information(self, "Saved", f"File saved to table {table_name}")

class RoundUploadWidget(QWidget):
    """Wrapper to hold a SingleFileUpload and provide save/reset."""
    def __init__(self, title=None, required_map=None, table_name_fn=None):
        super().__init__()
        self.title = title
        self.required_map = required_map
        self.table_name_fn = table_name_fn
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.upload_widget = SingleFileUpload(title, required_map, table_name_fn)
        self.layout.addWidget(self.upload_widget)
    def get_file_path(self):
        if self.upload_widget:
            return self.upload_widget.file_path
        return None
    
    def save_to_db(self, round_no=None):
        if self.upload_widget:
            self.upload_widget.save_to_db(round_no)

    def reset_widget(self):
        if not self.upload_widget:
            return
        self.upload_widget.file_path = None
        self.upload_widget.df = None
        self.upload_widget.col_map = {}
        self.upload_widget.title_label.setText(f"{self.upload_widget.title}: <font color='red'>No file uploaded</font>")
        self.upload_widget.get_cols_btn.setEnabled(False)
        self.upload_widget.save_btn.setEnabled(False)
        if self.upload_widget.table_widget:
            self.upload_widget.layout.removeWidget(self.upload_widget.table_widget)
            self.upload_widget.table_widget.deleteLater()
            self.upload_widget.table_widget = None
