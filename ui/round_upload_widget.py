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
        if round_no is None:
            # Attempt to fetch the selected round number from the parent widget (RoundsWidget)
            # We use an integer fallback (1) instead of allowing boolean/None
            current_round_text = None
            parent_widget = self.parent() 
            while parent_widget is not None and not hasattr(parent_widget, "get_current_round"):
                parent_widget = parent_widget.parent()

            if parent_widget and hasattr(parent_widget, "get_current_round"):
                try:
                    # NOTE: When running Round N, decisions are for Round N-1. 
                    # This internal save logic should save the round currently selected in the combo box.
                    round_no = int(parent_widget.get_current_round())
                except Exception:
                    # Fallback to 1 if conversion fails
                    round_no = 1
            else:
                # Default to Round 1 if the parent structure cannot be found
                round_no = 1
        
        # Ensure round_no is an integer (required by self.table_name_fn)
        if not isinstance(round_no, int) or round_no < 1:
            QMessageBox.critical(self, "Error", "Invalid round number determined for saving.")
            return

        # Build table name with correct round number
        table_name = self.table_name_fn(round_no)
        print(f"[DEBUG] Saving data to table: {table_name}")

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        # Determine the primary key column name based on the table name
        pk_col = "mtech_app_no" if "goa" in table_name or "other_institute" in table_name else "coap_reg_id"

        # Map selected Excel column names to the expected standard DB column names
        # We only keep the required columns and rename them to the standard snake_case
        required_db_cols = [db_col for db_col, _ in self.required_map]
        
        # This new mapping ensures that the data saved locally matches what rounds_manager.py expects, 
        # even if it's currently redundant due to the main round logic using file paths.
        renamed_df = pd.DataFrame()
        for db_col, excel_col in self.col_map.items():
            renamed_db_col_name = {
                "Mtech App No": "mtech_app_no",
                "Other Institute Decision": "other_institute_decision",
                "COAP Reg Id": "coap_reg_id",
                "Applicant Decision": "applicant_decision"
            }.get(db_col, db_col) # Use a standardized mapping

            if excel_col in self.df.columns:
                 renamed_df[renamed_db_col_name] = self.df[excel_col]

        # Save to DB using pandas, replacing the manual cursor execution
        # Ensure only the necessary columns exist
        if renamed_df.empty:
            QMessageBox.critical(self, "Error", "Selected columns could not be found or mapped correctly.")
            conn.close()
            return
            
        try:
            # Drop and recreate table with correct structure using pandas and Primary Key
            # NOTE: Pandas' to_sql doesn't directly support adding PRIMARY KEY easily without raw SQL.
            # However, since rounds_manager._create_decision_tables is called separately, we rely on that for schema.
            # We simply use replace here for simplicity.
            renamed_df.to_sql(table_name, conn, if_exists='replace', index=False)
            conn.commit()
            QMessageBox.information(self, "Saved", f"File saved to table {table_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save data to DB table {table_name}:\n{e}")
        finally:
            conn.close()
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
