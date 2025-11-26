REQUIRED_MAPPING_TARGETS = [
    "COAP",
    "App_no",
    "Email",
    "Full_Name",
    "MaxGATEScore_3yrs",
    "Pwd",
    "Ews",
    "Gender",
    "Category",
    "GATE22Score",
    "GATE23Score",
    "GATE24Score",
    "GATE22RollNo",
    "GATE23RollNo",
    "GATE24RollNo",
    "HSSC_per",
    "SSC_per",
    "Degree_Per_8th",
    "Degree_CGPA_8th"
]

# main_window.py
import sqlite3
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow,QDialog, QWidget, QVBoxLayout, QMessageBox, 
    QTabWidget, QPushButton, QFileDialog, QLabel, QComboBox, QTableWidget, 
    QTableWidgetItem, QScrollArea, QGroupBox, QToolBox, QHBoxLayout
)
import re, difflib, json
import numpy as np
from ui.update_dialog import UpdateDialog
# IMPORTANT CHANGE: Import the generic multi-round functions
from ui.rounds_manager import run_round, download_offers, upload_round_decisions 
import pandas as pd
from database import db_manager 
from ui.round_upload_widget import RoundUploadWidget
from ui.search_page import SearchPage
from ui.seat_matrix_upload import SeatMatrixUpload

DB_NAME = "mtech_offers.db"
def _coerce_df_for_sql(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert DataFrame values to SQLite-safe Python types:
    - Timestamp -> YYYY-MM-DD string
    - numpy scalars -> python native types
    - NaN -> None
    """
    out = df.copy()
    out = out.where(pd.notnull(out), None)

    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].apply(lambda x: x.strftime('%Y-%m-%d') if x else None)
            continue

        out[col] = out[col].apply(
            lambda v: (
                v.strftime('%Y-%m-%d') if hasattr(v, "to_pydatetime")
                else (v.item() if isinstance(v, np.generic) else v)
            )
        )
    return out

class MainWindow(QMainWindow):
    def __init__(self):
        from database import db_manager
        db_manager.init_db()
        super().__init__()
        self.setWindowTitle("MTech Offers Automation")
        self.resize(900, 600)
        self.total_rounds = 10 
        
        # Create tab widget
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Initialization tab
        self.init_tab = QWidget()
        self.setup_init_tab()
        self.tabs.addTab(self.init_tab, "Initialization")

        # Seat matrix tab
        self.seat_matrix_tab = SeatMatrixTab()
        self.tabs.addTab(self.seat_matrix_tab, "Seat Matrix")
        
        # Rounds tab
        self.rounds_tab = RoundsWidget(total_rounds=self.total_rounds)
        self.tabs.addTab(self.rounds_tab, "Rounds")
        
        # Search tab
        self.search_tab = SearchPage(db_path="mtech_offers.db")
        self.search_tab.updateRequested.connect(self.open_update_page) 
        self.tabs.addTab(self.search_tab, "Search")

    def setup_init_tab(self):
        layout = QVBoxLayout()
        self.init_tab.setLayout(layout)

        # Upload Excel button
        self.upload_btn = QPushButton("Upload Applicants Excel")
        self.upload_btn.clicked.connect(self.upload_excel)
        layout.addWidget(self.upload_btn)

        # Status label
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)
        
    def open_update_page(self, record: dict):
    # lazy import to avoid circulars
        coap = record.get("coap_id")
        if not coap:
            QMessageBox.warning(self, "Missing COAP", "Could not read COAP from the selected row.")
            return

        dlg = UpdateDialog(DB_NAME, coap, self)
        dlg.exec()

    def upload_excel(self):
        from ui.mapping_preview import MappingPreviewDialog

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            df = df.loc[:, df.columns.notnull()]
            df = df.loc[:, ~df.columns.duplicated()]
            df = df.where(pd.notnull(df), None)

            # Read DB columns
            conn = db_manager.get_connection()
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(candidates)")
            table_columns = [info[1] for info in cursor.fetchall()]
            conn.close()

            # Header normalization
            import re, difflib
            def norm(s):
                s = str(s).lower()
                s = re.sub(r'[^a-z0-9]+', ' ', s)
                return s.strip()

            src_norm_map = {norm(c): c for c in df.columns}
            src_norm_list = list(src_norm_map.keys())

            synonyms = {
                "app no": "App_no",
                "application number": "App_no",
                "full name": "Full_Name",
                "coap id": "COAP",
                "gate roll": "GATE_Roll_num",
                "max gate score": "MaxGATEScore_3yrs"
            }

            # Auto match
            mapping = {}
            for tgt in table_columns:
                tgt_norm = norm(tgt)

                # Exact
                if tgt_norm in src_norm_map:
                    mapping[tgt] = src_norm_map[tgt_norm]
                    continue

                # Synonym
                found = False
                for syn, real in synonyms.items():
                    if real == tgt and norm(syn) in src_norm_map:
                        mapping[tgt] = src_norm_map[norm(syn)]
                        found = True
                        break
                if found:
                    continue

                # Fuzzy
                best = None
                best_score = 0
                for src_norm in src_norm_list:
                    score = difflib.SequenceMatcher(None, tgt_norm, src_norm).ratio()
                    if score > best_score:
                        best_score = score
                        best = src_norm

                mapping[tgt] = src_norm_map[best] if best_score >= 0.65 else None

            # Preview dialog
            dlg = MappingPreviewDialog(mapping, df.columns, required_targets=REQUIRED_MAPPING_TARGETS, parent=self)

            if dlg.exec() != QDialog.Accepted:
                self.status_label.setText("Upload cancelled.")
                return

            final_map = dlg.get_final_mapping()
            rename_map = {}
            for tgt, src in final_map.items():
                if src:
                    rename_map[src] = tgt

            df = df.rename(columns=rename_map)

            # prepare insert columns (same as you did)
            insert_cols = [c for c in df.columns if c in table_columns]
            if not insert_cols:
                self.status_label.setText("No valid mapped columns!")
                return

            # convert all values to sqlite-safe Python types
            df_for_sql = _coerce_df_for_sql(df)

            # build SQL
            placeholders = ', '.join(['?'] * len(insert_cols))
            cols_sql = ', '.join([f'"{c}"' for c in insert_cols])   # quoted columns

            conn = db_manager.get_connection()
            cur = conn.cursor()
            try:
                for _, row in df_for_sql.iterrows():
                    vals = [row.get(c) for c in insert_cols]
                    cur.execute(
                        f'INSERT OR IGNORE INTO candidates ({cols_sql}) VALUES ({placeholders})',
                        vals
                    )
                conn.commit()
                self.status_label.setText("Excel uploaded & mapped successfully!")
            except Exception as e:
                conn.rollback()
                self.status_label.setText(f"Error saving to DB: {e}")
            finally:
                conn.close()


            self.status_label.setText("Excel uploaded & mapped successfully!")

        except Exception as e:
            self.status_label.setText(f"Error: {e}")
    
    # NOTE: setup_rounds_tab is no longer needed as RoundsWidget handles its own setup
# # main_window.py
# import sqlite3
# from PySide6.QtCore import Qt
# from PySide6.QtWidgets import (
#     QApplication, QMainWindow,QDialog, QWidget, QVBoxLayout, QMessageBox, 
#     QTabWidget, QPushButton, QFileDialog, QLabel, QComboBox, QTableWidget, 
#     QTableWidgetItem, QScrollArea, QGroupBox, QToolBox, QHBoxLayout
# )
# from ui.mapping_preview import MappingDialog
# from ui.update_dialog import UpdateDialog
# # IMPORTANT CHANGE: Import the generic multi-round functions
# from ui.rounds_manager import run_round, download_offers, upload_round_decisions 
# import pandas as pd
# from database import db_manager 
# from ui.round_upload_widget import RoundUploadWidget
# from ui.search_page import SearchPage
# from ui.seat_matrix_upload import SeatMatrixUpload


# DB_NAME = "mtech_offers.db"

# class MainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("MTech Offers Automation")
#         self.resize(900, 600)
#         self.total_rounds = 10 
        
#         # Create tab widget
#         self.tabs = QTabWidget()
#         self.setCentralWidget(self.tabs)

#         # Initialization tab
#         self.init_tab = QWidget()
#         self.setup_init_tab()
#         self.tabs.addTab(self.init_tab, "Initialization")

#         # Seat matrix tab
#         self.seat_matrix_tab = SeatMatrixTab()
#         self.tabs.addTab(self.seat_matrix_tab, "Seat Matrix")
        
#         # Rounds tab
#         self.rounds_tab = RoundsWidget(total_rounds=self.total_rounds)
#         self.tabs.addTab(self.rounds_tab, "Rounds")
        
#         # Search tab
#         self.search_tab = SearchPage(db_path="mtech_offers.db")
#         self.search_tab.updateRequested.connect(self.open_update_page) 
#         self.tabs.addTab(self.search_tab, "Search")

#     def setup_init_tab(self):
#         layout = QVBoxLayout()
#         self.init_tab.setLayout(layout)

#         # Upload Excel button
#         self.upload_btn = QPushButton("Upload Applicants Excel")
#         self.upload_btn.clicked.connect(self.upload_excel)
#         layout.addWidget(self.upload_btn)

#         # Status label
#         self.status_label = QLabel("")
#         layout.addWidget(self.status_label)
        
#     def open_update_page(self, record: dict):
#     # lazy import to avoid circulars
#         coap = record.get("coap_id")
#         if not coap:
#             QMessageBox.warning(self, "Missing COAP", "Could not read COAP from the selected row.")
#             return

#         dlg = UpdateDialog(DB_NAME, coap, self)
#         dlg.exec()

#     # def upload_excel(self):
#     #     file_path, _ = QFileDialog.getOpenFileName(
#     #         self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
#     #     )
#     #     if not file_path:
#     #         return

#     #     try:
#     #         # Read Excel
#     #         df = pd.read_excel(file_path)

#     #         # Remove empty or duplicate columns
#     #         df = df.loc[:, df.columns.notnull()]
#     #         df = df.loc[:, ~df.columns.duplicated()]

#     #         # Rename Excel columns to match database exactly
#     #         column_mapping = {
#     #             "Si NO": "Si_NO",
#     #             "App no": "App_no",
#     #             "Full Name": "Full_Name",
#     #             "Adm cat": "Adm_cat",
#     #             "MaxGATEScore out of 3 yrs": "MaxGATEScore_3yrs",
#     #             "HSSC(date)": "HSSC_date",
#     #             "HSSC(board)": "HSSC_board",
#     #             "HSSC(per)": "HSSC_per",
#     #             "SSC(date)": "SSC_date",
#     #             "SSC(board)": "SSC_board",
#     #             "SSC(per)": "SSC_per",
#     #             "Degree(PassingDate)": "Degree_PassingDate",
#     #             "Degree(Qualification)": "Degree_Qualification",
#     #             "Degree(Branch)": "Degree_Branch",
#     #             "Degree(OtherBranch)": "Degree_OtherBranch",
#     #             "Degree(Institute Name)": "Degree_Institute",
#     #             "Degree(CGPA-7thSem)": "Degree_CGPA_7th",
#     #             "Degree(CGPA-8thSem)": "Degree_CGPA_8th",
#     #             "Degree(Per-7thSem)": "Degree_Per_7th",
#     #             "Degree(Per-8thSem)": "Degree_Per_8th",
#     #             "GATE Roll num": "GATE_Roll_num",
#     #             "unnamed": "ExtraColumn"
#     #         }
#     #         df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns}, inplace=True)

#     #         # Convert datetime columns to string
#     #         for col in ['HSSC_date', 'SSC_date', 'Degree_PassingDate']:
#     #             if col in df.columns:
#     #                 df[col] = df[col].apply(
#     #                     lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) and not isinstance(x, str) else x
#     #                 )

#     #         # Connect to DB
#     #         conn = db_manager.get_connection()
#     #         cursor = conn.cursor()

#     #         # Get list of columns in DB
#     #         cursor.execute("PRAGMA table_info(candidates)")
#     #         table_columns = [info[1] for info in cursor.fetchall()]

#     #         # Only keep columns that exist in DB
#     #         insert_columns = [c for c in df.columns if c in table_columns]
#     #         placeholders = ', '.join(['?'] * len(insert_columns))

#     #         # Insert rows
#     #         for _, row in df.iterrows():
#     #             values = [row[c] for c in insert_columns]
#     #             cursor.execute(
#     #                 f'INSERT OR IGNORE INTO candidates ({", ".join(insert_columns)}) VALUES ({placeholders})',
#     #                 values
#     #             )

#     #         conn.commit()
#     #         conn.close()
#     #         self.status_label.setText("Excel data inserted successfully into database!")

#     #     except Exception as e:
#     #         self.status_label.setText(f"Error: {str(e)}")
    
#     # NOTE: setup_rounds_tab is no longer needed as RoundsWidget handles its own setup
#     def upload_excel(self):
#         file_path, _ = QFileDialog.getOpenFileName(
#             self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
#         )
#         if not file_path:
#             return

#         try:
#             # Step A: Read ONLY HEADERS to pass to the mapping dialog
#             df_headers = pd.read_excel(file_path, nrows=0)
#             excel_columns = df_headers.columns.tolist()

#         except Exception as e:
#             self.status_label.setText(f"Error reading Excel headers: {str(e)}")
#             return

#         # Step B: Show the Mapping Dialog
#         # NOTE: You must define REQUIRED_DB_FIELDS, likely imported from db_manager or mapping_dialog
#         # Pass the list of Excel column headers to the dialog.
#         mapping_dlg = MappingDialog(excel_columns, self)
        
#         if mapping_dlg.exec() != QDialog.Accepted:
#             self.status_label.setText("Excel upload cancelled by user.")
#             return
            
#         final_mapping = mapping_dlg.get_mapping()

#         # Step C: Read the full Excel file and apply the user's mapping
#         try:
#             # Read full data
#             df = pd.read_excel(file_path)

#             # Remove empty or duplicate columns (pre-cleanup still necessary)
#             df = df.loc[:, df.columns.notnull()]
#             df = df.loc[:, ~df.columns.duplicated()]

#             # Apply the mapping provided by the user in the UI
#             df.rename(columns=final_mapping, inplace=True)
            
#             # -----------------------------------------------------------
#             # IMPORTANT: Select only the columns that were mapped/renamed
#             # -----------------------------------------------------------
            
#             # Get the list of DB column names the user chose to map
#             required_db_cols = list(final_mapping.values()) 
            
#             # Filter the DataFrame to only include these columns
#             df_to_insert = df[[c for c in required_db_cols if c in df.columns]]

#             # Convert datetime columns to string (your original logic)
#             for col in ['HSSC_date', 'SSC_date', 'Degree_PassingDate']:
#                 if col in df_to_insert.columns:
#                     df_to_insert[col] = df_to_insert[col].apply(
#                         lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) and not isinstance(x, str) else x
#                     )

#             # Step D: Connect and Insert (Your existing insertion logic, using df_to_insert)
#             conn = db_manager.get_connection()
#             cursor = conn.cursor()

#             insert_columns = df_to_insert.columns.tolist()
#             placeholders = ', '.join(['?'] * len(insert_columns))

#             for _, row in df_to_insert.iterrows():
#                 # Use df_to_insert's columns, which are now correctly named DB fields
#                 cursor.execute(
#                     f'INSERT OR IGNORE INTO candidates ({", ".join(insert_columns)}) VALUES ({placeholders})',
#                     tuple(row)
#                 )

#             conn.commit()
#             conn.close()
#             self.status_label.setText("Excel data inserted successfully into database using custom mapping!")

#         except Exception as e:
            # self.status_label.setText(f"Error during import: {str(e)}")
# ----------------------------------------------------------------------
# SeatMatrixTab (No changes needed)
# ----------------------------------------------------------------------

class SeatMatrixTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout(self)
        self.toolbox = QToolBox()
        layout.addWidget(self.toolbox)

        self.categories = {
            "COMMON_PWD": ["COMMON_PWD"],
            "EWS": ["EWS_FandM", "EWS_FandM_PWD", "EWS_Female", "EWS_Female_PWD"],
            "GEN": ["GEN_FandM", "GEN_FandM_PWD", "GEN_Female", "GEN_Female_PWD"],
            "OBC": ["OBC_FandM", "OBC_FandM_PWD", "OBC_Female", "OBC_Female_PWD"],
            "SC": ["SC_FandM", "SC_FandM_PWD", "SC_Female", "SC_Female_PWD"],
            "ST": ["ST_FandM", "ST_FandM_PWD", "ST_Female", "ST_Female_PWD"],
        }

        self.tables = {}
        self.create_sections()

        # Save button
        btn_layout = QHBoxLayout()
        self.save_btn = QPushButton("Save Seat Matrix")
        self.save_btn.clicked.connect(self.save_matrix)
        btn_layout.addWidget(self.save_btn)
        manual_layout.addLayout(btn_layout)

        # Load initial state from DB
        self.load_matrix()

    def _on_upload_clicked(self):
        """Wrapper called when the Upload button is clicked.
        It calls the upload widget's upload flow and then reloads the seat matrix from DB.
        """
        try:
            # trigger the upload flow (this opens the file dialog inside SeatMatrixUpload)
            self.upload_widget.upload_excel()
        except Exception as e:
            # keep UX friendly: show a message but continue to attempt reload
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Upload Error", f"Upload failed: {e}")

        # Reload whatever is in DB now (works whether upload succeeded or not)
        self.load_matrix()
        self.tabs.addTab(self.manual_tab, "Manual Entry")

        # --- Tab 2: Upload Excel ---
        self.upload_tab = SeatMatrixUpload()
        self.tabs.addTab(self.upload_tab, "Upload Excel")

    def create_sections(self):
        """Create collapsible sections (QToolBox) for each main category."""
        for section, subcats in self.categories.items():
            table = QTableWidget()
            table.setRowCount(len(subcats))
            table.setColumnCount(3)
            table.setHorizontalHeaderLabels(["Set Seats", "Seats Allocated", "Seats Booked"])

            for i, sub in enumerate(subcats):
                header_item = QTableWidgetItem(sub)
                header_item.setFlags(header_item.flags() & ~Qt.ItemIsEditable)
                table.setVerticalHeaderItem(i, header_item)

                for j in range(3):
                    val = QTableWidgetItem("0")
                    if j != 0:  
                        val.setFlags(val.flags() & ~Qt.ItemIsEditable)
                    table.setItem(i, j, val)

            self.toolbox.addItem(table, section)
            self.tables[section] = table

    def load_matrix(self):
        """Load data from seat_matrix table into GUI."""
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT category, set_seats, seats_allocated, seats_booked FROM seat_matrix")
        data = cursor.fetchall()
        conn.close()

        # fill GUI with DB values
        for category, set_seats, seats_allocated, seats_booked in data:
            for section, table in self.tables.items():
                for r in range(table.rowCount()):
                    if table.verticalHeaderItem(r).text() == category:
                        table.blockSignals(True)
                        table.item(r, 0).setText(str(set_seats))
                        table.item(r, 1).setText(str(seats_allocated))
                        table.item(r, 2).setText(str(seats_booked))
                        table.blockSignals(False)

    def save_matrix(self):
        """Save data back to the database."""
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        for section, table in self.tables.items():
            for r in range(table.rowCount()):
                category = table.verticalHeaderItem(r).text()
                set_seats = int(table.item(r, 0).text())
                seats_allocated = int(table.item(r, 1).text())
                seats_booked = int(table.item(r, 2).text())
                cursor.execute("""
                    INSERT OR REPLACE INTO seat_matrix (category, set_seats, seats_allocated, seats_booked)
                    VALUES (?, ?, ?, ?)
                """, (category, set_seats, seats_allocated, seats_booked))
        conn.commit()
        conn.close()
        
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Saved Successfully")
        msg.setText("Seat Matrix data has been saved to the database successfully!")
        msg.exec()
class RoundsWidget(QWidget):
    def __init__(self, total_rounds=10):
        super().__init__()
        self.total_rounds = total_rounds
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        # ------------------ Round Selection ------------------
        round_layout = QHBoxLayout()
        round_layout.addWidget(QLabel("Select Round:"))
        self.round_combo = QComboBox()
        round_layout.addWidget(self.round_combo)
        self.layout.addLayout(round_layout)

        # ------------------ Single Combined Upload Widget ------------------
        # This widget contains all THREE uploads inside it
        self.upload_widget = RoundUploadWidget()
        self.layout.addWidget(self.upload_widget)

        # ------------------ Action Buttons ------------------
        btn_layout = QHBoxLayout()
        self.generate_btn = QPushButton("Generate Offers")
        self.generate_btn.clicked.connect(self.run_round)
        btn_layout.addWidget(self.generate_btn)

        self.download_btn = QPushButton("Download Offers")
        self.download_btn.clicked.connect(self.download_current_round_offers)
        btn_layout.addWidget(self.download_btn)

        self.reset_btn = QPushButton("Reset Uploaded Files")
        self.reset_btn.clicked.connect(self.reset_round)
        btn_layout.addWidget(self.reset_btn)

        self.layout.addLayout(btn_layout)

        # ------------------ Signals ------------------
        self.round_combo.currentIndexChanged.connect(self.update_ui_visibility)

        # Populate rounds
        self.refresh_rounds()
        self.update_ui_visibility()

    # ------------------ Logic ------------------
    def get_current_round(self):
        if self.round_combo.count() == 0:
            return 1
        return int(self.round_combo.currentText())
    def get_file_path(self):
        if self.upload_widget:
            return self.upload_widget.file_path
        return None
    def refresh_rounds(self):
        self.round_combo.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='offers'")
        if cursor.fetchone() is None:
            max_round = 0
        else:
            cursor.execute("SELECT MAX(round_no) FROM offers")
            max_round = cursor.fetchone()[0] or 0
        conn.close()

        start_round = 1
        end_round = (max_round + 1)
        end_round = min(end_round, self.total_rounds)

        for r in range(start_round, end_round + 1):
            self.round_combo.addItem(str(r))

        self.update_ui_visibility()
    
    def update_ui_visibility(self):
        if self.round_combo.count() == 0:
            return

        round_no = self.get_current_round()
        is_current_round_run = (round_no < self.round_combo.count())
        is_next_round_upload = (round_no > 1 and round_no == self.round_combo.count())

        # Clear previous uploads when moving to new upload round
        if is_next_round_upload:
            self.upload_widget.goa_widget.reset_widget()
            self.upload_widget.other_widget.reset_widget()
            self.upload_widget.cons_widget.reset_widget()

        if round_no == 1 and not is_current_round_run:
            # Round 1 requires NO upload
            self.upload_widget.setVisible(False)
            self.generate_btn.setEnabled(True)
            self.reset_btn.setVisible(False)

        elif is_current_round_run:
            # Past round already generated
            self.upload_widget.setVisible(False)
            self.generate_btn.setEnabled(False)
            self.reset_btn.setVisible(False)

        else:
            # Round > 1 → show upload section
            self.upload_widget.setVisible(True)
            self.generate_btn.setEnabled(True)
            self.reset_btn.setVisible(True)

    def run_round(self):
        round_no = self.get_current_round()

        if round_no > 1:
            # Check if all 3 files uploaded
            file_paths = [
                self.upload_widget.goa_widget.get_file_path(),
                self.upload_widget.other_widget.get_file_path(),
                self.upload_widget.cons_widget.get_file_path()
            ]

            if not all(file_paths):
                QMessageBox.critical(self, "Missing Files",
                                    f"Please upload all 3 decision files for Round {round_no - 1}.")
                return

            prev_round = round_no - 1

            # Upload decision files
            upload_round_decisions(
                round_no=prev_round,
                goa_widget=self.upload_widget.goa_widget,
                other_widget=self.upload_widget.other_widget,
                cons_widget=self.upload_widget.cons_widget
            )

        # Run allocation for current round
        run_round(round_no)

        # Update UI
        self.refresh_rounds()
        self.update_ui_visibility()

    def download_current_round_offers(self):
        round_no = self.get_current_round()
        download_offers(round_no)

    def reset_round(self):
        round_no = self.get_current_round()
        if round_no == 1:
            QMessageBox.warning(self, "Warning",
                                "Round 1 does not have decision uploads.")
            return

        prev_round = round_no - 1

        # Drop tables created when decisions were uploaded
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        try:
            cursor.execute(f"DROP TABLE IF EXISTS iit_goa_offers_round{prev_round}")
            cursor.execute(f"DROP TABLE IF EXISTS accepted_other_institute_round{prev_round}")
            cursor.execute(f"DROP TABLE IF EXISTS consolidated_decisions_round{prev_round}")
            conn.commit()
        finally:
            conn.close()

        # Reset UI
        self.upload_widget.goa_widget.reset_widget()
        self.upload_widget.other_widget.reset_widget()
        self.upload_widget.cons_widget.reset_widget()

        QMessageBox.information(
            self,
            "Reset Complete",
            f"Decision uploads for Round {prev_round} have been cleared!"
        )

# ----------------------------------------------------------------------
# RoundsWidget (UPDATED LOGIC)
# ---------------------------------------------------------------------
# class RoundsWidget(QWidget):
#     def __init__(self, total_rounds=10):
#         super().__init__()
#         self.total_rounds = total_rounds
#         self.layout = QVBoxLayout()
#         self.setLayout(self.layout)

#         # ------------------ Round Selection ------------------
#         round_layout = QHBoxLayout()
#         round_layout.addWidget(QLabel("Select Round:"))
#         self.round_combo = QComboBox()
#         round_layout.addWidget(self.round_combo)
#         self.layout.addLayout(round_layout)

#         # ------------------ File Upload Widgets ------------------
#         # File 1: IIT Goa Candidate Decision Report (Uses Mtech App No)
#         required_map_1 = [
#             ("Mtech App No", "Mtech App No"), 
#             ("Applicant Decision", "Applicant Decision")
#         ]
#         table_name_fn_1 = lambda round_no: f"iit_goa_offers_round{round_no}"
#         self.upload1 = RoundUploadWidget(
#             title="1. IIT Goa Offered Candidate Decision File",
#             required_map=required_map_1,
#             table_name_fn=table_name_fn_1
#         )
#         self.layout.addWidget(self.upload1)

#         # File 2: IIT Goa Offered But Accept and Freeze at Other Institutes (Uses Mtech App No)
#         required_map_2 = [
#             ("Mtech App No", "Mtech App No"),
#             ("Other Institute Decision", "Other Institute Decision")
#         ]
#         table_name_fn_2 = lambda round_no: f"accepted_other_institute_round{round_no}"
#         self.upload2 = RoundUploadWidget(
#             title="2. IIT Goa Offered But Accepted at Different Institute File",
#             required_map=required_map_2,
#             table_name_fn=table_name_fn_2
#         )
#         self.layout.addWidget(self.upload2)

#         # File 3: Consolidated Accept and Freeze Candidates Across All Institutes (Uses COAP Reg Id)
#         required_map_3 = [
#             ("COAP Reg Id", "COAP Reg Id"),
#             ("Applicant Decision", "Applicant Decision")
#         ]
#         table_name_fn_3 = lambda round_no: f"consolidated_decisions_round{round_no}"
#         self.upload3 = RoundUploadWidget(
#             title="3. Consolidated Decision File",
#             required_map=required_map_3,
#             table_name_fn=table_name_fn_3
#         )
#         self.layout.addWidget(self.upload3)

#         # ------------------ Action Buttons ------------------
#         btn_layout = QHBoxLayout()
#         self.generate_btn = QPushButton("Generate Offers")
#         self.generate_btn.clicked.connect(self.run_round)
#         btn_layout.addWidget(self.generate_btn)

#         self.download_btn = QPushButton("Download Offers")
#         self.download_btn.clicked.connect(self.download_current_round_offers)
#         btn_layout.addWidget(self.download_btn)

#         self.reset_btn = QPushButton("Reset Uploaded Files")
#         self.reset_btn.clicked.connect(self.reset_round)
#         btn_layout.addWidget(self.reset_btn)

#         self.layout.addLayout(btn_layout)

#         # ------------------ Signals ------------------
#         self.round_combo.currentIndexChanged.connect(self.update_ui_visibility)

#         # Populate combo box after widgets are created
#         self.refresh_rounds()
#         self.update_ui_visibility()

#     # ------------------ Logic ------------------
#     def get_current_round(self):
#         """Return selected round as int."""
#         if self.round_combo.count() == 0:
#             return 1
#         return int(self.round_combo.currentText())

#     def refresh_rounds(self):
#         """Populate dropdown based on already generated rounds."""
#         self.round_combo.clear()
#         conn = sqlite3.connect(DB_NAME)
#         cursor = conn.cursor()

#         cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='offers'")
#         if cursor.fetchone() is None:
#             max_round = 0
#         else:
#             cursor.execute("SELECT MAX(round_no) FROM offers")
#             max_round = cursor.fetchone()[0] or 0
#         conn.close()

#         start_round = 1
#         end_round = (max_round + 1) if max_round else 1
#         end_round = min(end_round, self.total_rounds)

#         for r in range(start_round, end_round + 1):
#             self.round_combo.addItem(str(r))

#         self.update_ui_visibility()

#     def update_ui_visibility(self):
#         """Control visibility of upload widgets and buttons based on round."""
#         if not hasattr(self, 'upload1') or self.round_combo.count() == 0:
#             return

#         round_no = self.get_current_round()
        
#         # Determine if the current round has already been run (i.e., is not the last one in the combo box)
#         is_current_round_run = (round_no < self.round_combo.count())

#         # Check if the selected round is the next un-run round (where uploads are required)
#         is_next_upload_round = (round_no > 1 and round_no == self.round_combo.count())
        
#         # --- FIX: Clear old file paths when selecting a new round for upload ---
#         if is_next_upload_round:
#             # When we select a new round that requires uploads (Round 2, 3, etc.), 
#             # we must reset the old file paths from the previous round (Round 1, 2, etc.)
#             for upload_widget in [self.upload1, self.upload2, self.upload3]:
#                 upload_widget.reset_widget() 
#         # ---------------------------------------------------------------------

#         if round_no == 1 and not is_current_round_run:
#             # Round 1 (unrun) — show only offer generation + download
#             self.upload1.setVisible(False)
#             self.upload2.setVisible(False)
#             self.upload3.setVisible(False)
#             self.generate_btn.setEnabled(True)
#             self.generate_btn.setText(" Generate Round 1 Offers")
#             self.reset_btn.setVisible(False)
#         elif is_current_round_run:
#             # Already run round (show only download button)
#             self.upload1.setVisible(False)
#             self.upload2.setVisible(False)
#             self.upload3.setVisible(False)
#             self.generate_btn.setEnabled(False)
#             self.generate_btn.setText(f"Round {round_no} Already Generated")
#             self.reset_btn.setVisible(False)
#         else:
#             # Rounds > 1 (unrun) — show upload widgets + generate button
#             self.upload1.setVisible(True)
#             self.upload2.setVisible(True)
#             self.upload3.setVisible(True)
#             self.generate_btn.setEnabled(True)
#             self.generate_btn.setText(f"Generate Round {round_no} Offers")
#             self.reset_btn.setVisible(True)
#     def run_round(self):
#         """Run allocation for the current round."""
#         round_no = self.get_current_round()
        
#         # 1. Validation and Decision Upload (for Round 2 and above)
#         if round_no > 1:
#             prev_round = round_no - 1
#             # Check if all files are selected for decision upload
#             file_paths = [u.get_file_path() for u in [self.upload1, self.upload2, self.upload3]]
#             if not all(file_paths):
#                 QMessageBox.critical(self, "Missing Files", f"Please upload the three decision files for **Round {round_no - 1}** before running Round {round_no}.")
#                 return
            
#             try:
#                 # Upload the decisions of the PREVIOUS round (round_no - 1)
#                 upload_round_decisions(
#                     round_no=prev_round,
#                     iit_goa_report=file_paths[0],
#                     other_iit_report=file_paths[1],
#                     consolidated_report=file_paths[2]
#                 )
                
#             except Exception as e:
#                 QMessageBox.critical(self, "Upload Error", f"Failed to upload decisions for Round {round_no - 1}:\n{e}")
#                 return
                
#         # 2. Run Allocation
#         run_round(round_no) 

#         # 3. Finalize
#         self.refresh_rounds() # Refresh to show the next round option
#         self.update_ui_visibility() # Update button state

#     def download_current_round_offers(self):
#         """Download offers for the current round."""
#         round_no = self.get_current_round()
#         download_offers(round_no)

#     def reset_round(self):
#         """Reset uploaded files and their DB tables for current round."""
#         round_no = self.get_current_round()
#         if round_no == 1:
#             QMessageBox.warning(self, "Warning", "Cannot reset Round 1 uploads. It does not require decision files.")
#             return

#         if QMessageBox.question(self, "Confirm Reset", 
#                                 f"Are you sure you want to delete the uploaded decision files and corresponding database tables for **Round {round_no - 1}**?", 
#                                 QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
#             return

#         # We reset the files and tables for the PREVIOUS round (N-1)
#         prev_round = round_no - 1
#         for upload in [self.upload1, self.upload2, self.upload3]:
#             table_name = upload.table_name_fn(prev_round)
#             conn = sqlite3.connect(DB_NAME)
#             cursor = conn.cursor()
#             try:
#                 cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
#                 conn.commit()
#                 upload.reset_widget()
#             except Exception as e:
#                 QMessageBox.critical(self, "DB Error", f"Failed to drop table {table_name}: {e}")
#             finally:
#                 conn.close()
        
#         QMessageBox.information(self, "Reset Complete", f"Decision uploads and tables for Round {prev_round} cleared!")