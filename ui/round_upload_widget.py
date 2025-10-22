# from PySide6.QtWidgets import (
#     QApplication, QWidget, QVBoxLayout, QHBoxLayout,
#     QLabel, QPushButton, QFileDialog, QComboBox
# )
# import sys

# class RoundUploadWidget(QWidget):
#     def __init__(self, max_rounds=10):
#         super().__init__()
#         self.setWindowTitle("Rounds Upload")
#         self.max_rounds = max_rounds

#         self.uploaded_files = [None, None, None]  # temp storage for 3 files

#         # Main layout
#         self.layout = QVBoxLayout()

#         # Round selection
#         self.round_label = QLabel("Select Round:")
#         self.round_combo = QComboBox()
#         self.round_combo.addItems([f"Round {i+1}" for i in range(self.max_rounds)])
#         self.layout.addWidget(self.round_label)
#         self.layout.addWidget(self.round_combo)

#         # File upload buttons
#         self.file_buttons = []
#         self.file_labels = []
#         for i in range(3):
#             hbox = QHBoxLayout()
#             btn = QPushButton(f"Upload File {i+1}")
#             btn.clicked.connect(lambda checked, idx=i: self.upload_file(idx))
#             lbl = QLabel("No file selected")
#             hbox.addWidget(btn)
#             hbox.addWidget(lbl)
#             self.layout.addLayout(hbox)
#             self.file_buttons.append(btn)
#             self.file_labels.append(lbl)

#         # Reset button
#         self.reset_btn = QPushButton("Reset Files")
#         self.reset_btn.clicked.connect(self.reset_files)
#         self.layout.addWidget(self.reset_btn)

#         self.setLayout(self.layout)

#     def upload_file(self, idx):
#         file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Excel Files (*.xlsx *.xls);;All Files (*)")
#         if file_path:
#             self.uploaded_files[idx] = file_path
#             self.file_labels[idx].setText(file_path.split("/")[-1])

#     def reset_files(self):
#         self.uploaded_files = [None, None, None]
#         for lbl in self.file_labels:
#             lbl.setText("No file selected")

# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#     window = RoundUploadWidget(max_rounds=10)
#     window.show()
#     sys.exit(app.exec())

# gave red match columns first code

# ui/round_upload_widget.py
# round_upload_widget.py
# import os
# import sqlite3
# import pandas as pd
# from PySide6.QtCore import Qt
# from PySide6.QtWidgets import (
#     QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
#     QFileDialog, QComboBox, QTableWidget, QTableWidgetItem, QMessageBox, QSizePolicy
# )

# DB_NAME = "mtech_offers.db"

# def _sanitize_col_name(name: str) -> str:
#     """Sanitize Excel column names to valid SQLite column names."""
#     return "".join(c if (c.isalnum() or c == "_") else "_" for c in str(name)).lower()

# class RoundUploadWidget(QWidget):
#     """
#     Handles a single file upload + column-match section for a given round.
#     """
#     def __init__(self, title, required_map, table_name_fn, max_rounds=10):
#         super().__init__()
#         self.title = title
#         self.required_map = required_map  # [(db_col, human_label), ...]
#         self.table_name_fn = table_name_fn
#         self.max_rounds = max_rounds

#         self.file_path = None
#         self.df = None
#         self.col_map = {}  # db_col -> selected Excel column

#         # Layout
#         self.layout = QVBoxLayout()
#         self.setLayout(self.layout)

#         # Round selection dropdown
#         self.round_label = QLabel("Select Round:")
#         self.round_combo = QComboBox()
#         self.update_rounds()
#         self.layout.addWidget(self.round_label)
#         self.layout.addWidget(self.round_combo)

#         # File upload title + status
#         self.title_label = QLabel(f"{self.title}: <font color='red'>No file uploaded</font>")
#         self.layout.addWidget(self.title_label)

#         # Buttons row
#         row = QHBoxLayout()
#         self.upload_btn = QPushButton("Select File")
#         self.upload_btn.clicked.connect(self.select_file)
#         row.addWidget(self.upload_btn)

#         self.get_cols_btn = QPushButton("Get Column Names")
#         self.get_cols_btn.setEnabled(False)
#         self.get_cols_btn.clicked.connect(self.show_column_match_table)
#         row.addWidget(self.get_cols_btn)

#         self.save_btn = QPushButton("Save This File to DB")
#         self.save_btn.setEnabled(False)
#         self.save_btn.clicked.connect(self.save_to_db)
#         row.addWidget(self.save_btn)

#         self.layout.addLayout(row)

#         # Column mapping table
#         self.col_table = QTableWidget()
#         self.col_table.setColumnCount(2)
#         self.col_table.setHorizontalHeaderLabels(["Required Column", "Excel Column"])
#         self.layout.addWidget(self.col_table)
#         self.col_table.hide()

#     # ---------------- File Handling ----------------
#     def select_file(self):
#         file_path, _ = QFileDialog.getOpenFileName(
#             self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
#         )
#         if not file_path:
#             return

#         self.file_path = file_path
#         self.df = pd.read_excel(file_path)
#         self.title_label.setText(f"{self.title}: <font color='green'>{os.path.basename(file_path)}</font>")
#         self.get_cols_btn.setEnabled(True)

#     # ---------------- Column Matching ----------------
#     def show_column_match_table(self):
#         if self.df is None:
#             QMessageBox.warning(self, "No file", "Please select a file first.")
#             return

#         self.col_table.setRowCount(len(self.required_map))
#         for i, (db_col, label) in enumerate(self.required_map):
#             self.col_table.setItem(i, 0, QTableWidgetItem(label))
#             combo = QComboBox()
#             combo.addItems(self.df.columns.tolist())
#             self.col_table.setCellWidget(i, 1, combo)

#         self.col_table.show()
#         self.save_btn.setEnabled(True)

#     # ---------------- Save to DB ----------------
#     def save_to_db(self):
#         round_no = int(self.round_combo.currentText().split()[-1])  # e.g., "Round 2" -> 2
#         table_name = self.table_name_fn(round_no)

#         # Get mapping
#         self.col_map = {}
#         for i, (db_col, _) in enumerate(self.required_map):
#             widget = self.col_table.cellWidget(i, 1)
#             if widget:
#                 self.col_map[db_col] = widget.currentText()
#             else:
#                 self.col_map[db_col] = None

#         # Extract relevant columns
#         df_to_save = pd.DataFrame()
#         for db_col, excel_col in self.col_map.items():
#             if excel_col not in self.df.columns:
#                 QMessageBox.warning(self, "Column Missing", f"Column {excel_col} not found in Excel")
#                 return
#             df_to_save[db_col] = self.df[excel_col]

#         # Save to DB
#         conn = sqlite3.connect(DB_NAME)
#         cursor = conn.cursor()
#         # Create table if not exists
#         cols_sql = ", ".join([f"{_sanitize_col_name(c)} TEXT" for c in df_to_save.columns])
#         cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({cols_sql})")
#         # Clear previous data for this round
#         cursor.execute(f"DELETE FROM {table_name}")
#         # Insert new data
#         df_to_save.to_sql(table_name, conn, if_exists="append", index=False)
#         conn.commit()
#         conn.close()

#         QMessageBox.information(self, "Saved", f"Data saved to table: {table_name}")

#     # ---------------- Round Combo Update ----------------
#     def update_rounds(self):
#         self.round_combo.clear()
#         conn = sqlite3.connect(DB_NAME)
#         cursor = conn.cursor()
#         # Find max round generated in offers table
#         cursor.execute("SELECT MAX(round_no) FROM offers")
#         res = cursor.fetchone()
#         max_round = res[0] if res and res[0] else 0
#         # Populate dropdown up to next round
#         for i in range(1, max_round + 2):
#             self.round_combo.addItem(f"Round {i}")
#         conn.close()

#second code
# ui/round_upload_widget.py
# import os
# import sqlite3
# import pandas as pd
# from PySide6.QtCore import Qt
# from PySide6.QtWidgets import (
#     QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
#     QFileDialog, QComboBox, QTableWidget, QTableWidgetItem, QMessageBox, QSizePolicy
# )

# DB_NAME = "mtech_offers.db"

# def _sanitize_col_name(name: str) -> str:
#     """Sanitize SQL column names (letters, numbers, underscore only)."""
#     return "".join(c if c.isalnum() or c == "_" else "_" for c in str(name)).lower()


# class SingleFileUpload(QWidget):
#     """
#     Widget to handle single file upload + column mapping + save to DB
#     """
#     def __init__(self, title, required_map, table_name_fn):
#         super().__init__()
#         self.title = title
#         self.required_map = required_map  # list of tuples (db_col, human_label)
#         self.table_name_fn = table_name_fn  # function(round_no) -> table_name

#         self.file_path = None
#         self.df = None
#         self.col_map = {}

#         self.layout = QVBoxLayout()
#         self.setLayout(self.layout)

#         self.title_label = QLabel(f"{self.title}: <font color='red'>No file uploaded</font>")
#         self.layout.addWidget(self.title_label)

#         row = QHBoxLayout()
#         self.upload_btn = QPushButton("Select File")
#         self.upload_btn.clicked.connect(self.select_file)
#         row.addWidget(self.upload_btn)

#         self.get_cols_btn = QPushButton("Get Column Names")
#         self.get_cols_btn.setEnabled(False)
#         self.get_cols_btn.clicked.connect(self.show_column_match_table)
#         row.addWidget(self.get_cols_btn)

#         self.save_btn = QPushButton("Save to DB")
#         self.save_btn.setEnabled(False)
#         self.save_btn.clicked.connect(self.save_to_db)
#         row.addWidget(self.save_btn)

#         self.layout.addLayout(row)

#         self.table_widget = None

#     def select_file(self):
#         path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
#         if not path:
#             return
#         self.file_path = path
#         self.df = pd.read_excel(self.file_path)
#         self.title_label.setText(f"{self.title}: <font color='green'>{os.path.basename(path)}</font>")
#         self.get_cols_btn.setEnabled(True)

#     def show_column_match_table(self):
#         if self.table_widget:
#             self.layout.removeWidget(self.table_widget)
#             self.table_widget.deleteLater()
#             self.table_widget = None
#             self.col_map = {}

#         self.table_widget = QTableWidget()
#         self.table_widget.setColumnCount(2)
#         self.table_widget.setRowCount(len(self.required_map))
#         self.table_widget.setHorizontalHeaderLabels(["DB Column", "Excel Column"])
#         self.table_widget.verticalHeader().setVisible(False)

#         for i, (db_col, human_label) in enumerate(self.required_map):
#             self.table_widget.setItem(i, 0, QTableWidgetItem(human_label))
#             combo = QComboBox()
#             combo.addItems(self.df.columns.tolist())
#             combo.currentTextChanged.connect(lambda val, i=i, db=db_col: self.set_col_map(i, db, val))
#             self.table_widget.setCellWidget(i, 1, combo)
#             # preselect first column
#             self.set_col_map(i, db_col, self.df.columns[0])

#         self.layout.addWidget(self.table_widget)
#         self.save_btn.setEnabled(True)

#     def set_col_map(self, row, db_col, val):
#         self.col_map[db_col] = val

#     def save_to_db(self, round_no=None):
#         if not self.df or not self.col_map:
#             QMessageBox.warning(self, "Error", "No file or columns selected")
#             return
#         table_name = self.table_name_fn(round_no)
#         conn = sqlite3.connect(DB_NAME)
#         cursor = conn.cursor()

#         # Create table dynamically
#         cols = [f"{_sanitize_col_name(db_col)} TEXT" for db_col in self.col_map.keys()]
#         create_sql = f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join(cols)})"
#         cursor.execute(create_sql)

#         # Insert data
#         insert_cols = [ _sanitize_col_name(db_col) for db_col in self.col_map.keys()]
#         placeholders = ", ".join(["?"] * len(insert_cols))

#         rows = []
#         for _, row in self.df.iterrows():
#             rows.append([row[self.col_map[db_col]] for db_col in self.col_map.keys()])

#         cursor.executemany(
#             f"INSERT INTO {table_name} ({', '.join(insert_cols)}) VALUES ({placeholders})",
#             rows
#         )
#         conn.commit()
#         conn.close()
#         QMessageBox.information(self, "Saved", f"File saved to table {table_name}")

# class RoundUploadWidget(QWidget):
#     """
#     Handles a single file upload widget.
#     Can be used for multiple files by creating multiple instances.
#     """
#     def __init__(self, title=None, required_map=None, table_name_fn=None):
#         super().__init__()

#         self.title = title
#         self.required_map = required_map
#         self.table_name_fn = table_name_fn

#         self.layout = QVBoxLayout()
#         self.setLayout(self.layout)

#         # Create the upload widget
#         self.upload_widget = SingleFileUpload(
#             title=self.title,
#             required_map=self.required_map,
#             table_name_fn=self.table_name_fn
#         )
#         self.layout.addWidget(self.upload_widget)

#     def save_to_db(self, round_no=None):
    
#         if not self.upload_widget:
#             return

#         # Check inside the SingleFileUpload
#         if self.upload_widget.df is None or self.upload_widget.df.empty:
#             QMessageBox.warning(self, "Error", "No file uploaded or file is empty")
#             return

#         if not self.upload_widget.col_map:
#             QMessageBox.warning(self, "Error", "No columns selected")
#             return

#         # Call the actual save function
#         self.upload_widget.save_to_db(round_no)


#     def reset_widget(self):
#         """
#         Reset the upload widget for new file upload
#         """
#         if self.upload_widget:
#             self.upload_widget.file_path = None
#             self.upload_widget.df = None
#             self.upload_widget.col_map = {}
#             self.upload_widget.title_label.setText(f"{self.title}: <font color='red'>No file uploaded</font>")
#             self.upload_widget.get_cols_btn.setEnabled(False)
#             self.upload_widget.save_btn.setEnabled(False)
#             if self.upload_widget.table_widget:
#                 self.upload_widget.layout.removeWidget(self.upload_widget.table_widget)
#                 self.upload_widget.table_widget.deleteLater()
#                 self.upload_widget.table_widget = None
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
        # ✅ Correct DataFrame check
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "Error", "No file uploaded or file is empty")
            return
        if not self.col_map:
            QMessageBox.warning(self, "Error", "No columns selected")
            return

        table_name = self.table_name_fn(round_no)
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
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.upload_widget = SingleFileUpload(title, required_map, table_name_fn)
        self.layout.addWidget(self.upload_widget)

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
