# main_window.py
import sqlite3
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,QMessageBox, 
    QTabWidget, QPushButton, QFileDialog, QLabel,QComboBox, QTableWidget, QTableWidgetItem, QScrollArea, QGroupBox, QPushButton, QToolBox, QHBoxLayout
)
from ui.update_dialog import UpdateDialog
from ui.rounds_manager import run_round_1, download_offers
import pandas as pd
from database import db_manager  # your module with get_connection()
from ui.round_upload_widget import RoundUploadWidget
from ui.search_page import SearchPage

DB_NAME = "mtech_offers.db"

class MainWindow(QMainWindow):
    def __init__(self):
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

        # Seat matrix placeholder
        self.tabs.addTab(QWidget(), "Seat Matrix")

        # Rounds tab (your custom widget)
        self.rounds_tab = RoundsWidget(total_rounds=self.total_rounds)
        self.tabs.addTab(self.rounds_tab, "Rounds")
        
        self.search_tab = SearchPage(db_path="mtech_offers.db")
        self.search_tab.updateRequested.connect(self.open_update_page)  # define this slot to show the next page
        self.tabs.addTab(self.search_tab, "Search")

        # Properly initialize seat matrix
        self.seat_matrix_tab = SeatMatrixTab()
        self.tabs.removeTab(1)
        self.tabs.insertTab(1, self.seat_matrix_tab, "Seat Matrix")
        
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
        
    # def open_update_page(self, record: dict):
    #     """
    #     record: dict containing the row clicked in SearchPage
    #     You can show a QMessageBox, populate a new tab, or open a new window
    #     """
    #     # Example: just show the record for now
    #     from PySide6.QtWidgets import QMessageBox
    #     msg = QMessageBox()
    #     msg.setWindowTitle("Update Requested")
    #     msg.setText(f"UPDATE requested for:\n{record}")
    #     msg.exec()

    def upload_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if not file_path:
            return

        try:
            # Read Excel
            df = pd.read_excel(file_path)

            # Remove empty or duplicate columns
            df = df.loc[:, df.columns.notnull()]
            df = df.loc[:, ~df.columns.duplicated()]

            # Rename Excel columns to match database exactly
            column_mapping = {
                "Si NO": "Si_NO",
                "App no": "App_no",
                "Full Name": "Full_Name",
                "Adm cat": "Adm_cat",
                "MaxGATEScore out of 3 yrs": "MaxGATEScore_3yrs",
                "HSSC(date)": "HSSC_date",
                "HSSC(board)": "HSSC_board",
                "HSSC(per)": "HSSC_per",
                "SSC(date)": "SSC_date",
                "SSC(board)": "SSC_board",
                "SSC(per)": "SSC_per",
                "Degree(PassingDate)": "Degree_PassingDate",
                "Degree(Qualification)": "Degree_Qualification",
                "Degree(Branch)": "Degree_Branch",
                "Degree(OtherBranch)": "Degree_OtherBranch",
                "Degree(Institute Name)": "Degree_Institute",
                "Degree(CGPA-7thSem)": "Degree_CGPA_7th",
                "Degree(CGPA-8thSem)": "Degree_CGPA_8th",
                "Degree(Per-7thSem)": "Degree_Per_7th",
                "Degree(Per-8thSem)": "Degree_Per_8th",
                "GATE Roll num": "GATE_Roll_num",
                "unnamed": "ExtraColumn"
            }
            df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns}, inplace=True)

            # Convert datetime columns to string
            for col in ['HSSC_date', 'SSC_date', 'Degree_PassingDate']:
                if col in df.columns:
                    df[col] = df[col].apply(
                        lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) and not isinstance(x, str) else x
                    )

            # Connect to DB
            conn = db_manager.get_connection()
            cursor = conn.cursor()

            # Get list of columns in DB
            cursor.execute("PRAGMA table_info(candidates)")
            table_columns = [info[1] for info in cursor.fetchall()]

            # Only keep columns that exist in DB
            insert_columns = [c for c in df.columns if c in table_columns]
            placeholders = ', '.join(['?'] * len(insert_columns))

            # Insert rows
            for _, row in df.iterrows():
                values = [row[c] for c in insert_columns]
                cursor.execute(
                    f'INSERT OR IGNORE INTO candidates ({", ".join(insert_columns)}) VALUES ({placeholders})',
                    values
                )

            conn.commit()
            conn.close()
            self.status_label.setText("Excel data inserted successfully into database!")

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")
    
    # ------------------ Rounds Tab ------------------
    def setup_rounds_tab(self):
        layout = QVBoxLayout()
        self.rounds_tab.setLayout(layout)

        # Upload round file
        self.upload_round_btn = QPushButton("Upload Round Excel (placeholder)")
        layout.addWidget(self.upload_round_btn)

        # Status label
        self.round_status_label = QLabel("")
        layout.addWidget(self.round_status_label)
        
        # Table to show round offers
        self.round_table = QTableWidget()
        layout.addWidget(self.round_table)

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
        self.save_btn = QPushButton("💾 Save Seat Matrix")
        self.save_btn.clicked.connect(self.save_matrix)
        btn_layout.addWidget(self.save_btn)
        layout.addLayout(btn_layout)

        self.load_matrix()

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
                    if j != 0:  # only 'Set Seats' editable
                        val.setFlags(val.flags() & ~Qt.ItemIsEditable)
                    table.setItem(i, j, val)

            # 🔁 Auto-update seats_allocated when set_seats edited
            # table.itemChanged.connect(self.handle_item_changed)

            self.toolbox.addItem(table, section)
            self.tables[section] = table

    # def handle_item_changed(self, item):
    #     """Automatically update 'Seats Allocated' when 'Set Seats' changes."""
    #     if item.column() == 0:  # Set Seats column
    #         try:
    #             new_value = int(item.text())
    #             table = item.tableWidget()
    #             table.item(item.row(), 1).setText(str(new_value))  # copy to Seats Allocated
    #         except ValueError:
    #             pass  # ignore invalid (non-integer) entries

    def load_matrix(self):
        """Load data from seat_matrix table into GUI."""
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT category, set_seats, seats_allocated, seats_booked FROM seat_matrix")
        data = cursor.fetchall()
        conn.close()

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
        # ✅ Show confirmation message
        from PySide6.QtWidgets import QMessageBox
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Saved Successfully")
        msg.setText("✅ Seat Matrix data has been saved to the database successfully!")
        msg.exec()
        
# class RoundsWidget(QWidget):
#     def __init__(self,total_rounds=10):
#         super().__init__()
#         self.total_rounds = total_rounds
#         layout = QVBoxLayout()
#         self.round_upload_widget = RoundUploadWidget(max_rounds=self.total_rounds)
#         layout.addWidget(self.round_upload_widget)

#         layout.addWidget(QLabel("Round 1 Allocation"))

#         self.round1_btn = QPushButton("Run Round 1 Allocation")
#         self.round1_btn.clicked.connect(run_round_1)
#         layout.addWidget(self.round1_btn)

#         self.download_btn = QPushButton("Download Round 1 Offers")
#         self.download_btn.clicked.connect(lambda: download_offers(1))
#         layout.addWidget(self.download_btn)

#         self.setLayout(layout)
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
        self.refresh_rounds()
        self.layout.addLayout(round_layout)

        # ------------------ File Upload Widgets ------------------
        # File 1: IIT Goa Offered Candidate Decision File
        required_map_1 = [
            ("Mtech_App_no", "Mtech Application No"),
            ("Applicant_Decision", "Applicant Decision")
        ]
        table_name_fn_1 = lambda round_no: f"iit_goa_offers_round{round_no}"
        self.upload1 = RoundUploadWidget(
            title="IIT Goa Offered Candidate Decision File",
            required_map=required_map_1,
            table_name_fn=table_name_fn_1
        )
        self.layout.addWidget(self.upload1)

        # File 2: IIT Goa Offered But Accepted at Different Institute File
        required_map_2 = [
            ("Mtech_App_no", "Mtech Application No"),
            ("Other_Institute_Decision", "Other Institute Decision")
        ]
        table_name_fn_2 = lambda round_no: f"accepted_other_institute_round{round_no}"
        self.upload2 = RoundUploadWidget(
            title="IIT Goa Offered But Accepted at Different Institute File",
            required_map=required_map_2,
            table_name_fn=table_name_fn_2
        )
        self.layout.addWidget(self.upload2)

        # File 3: Consolidated Decision File
        required_map_3 = [
            ("COAP_Reg_ID", "COAP Reg ID"),
            ("Applicant_Decision", "Applicant Decision")
        ]
        table_name_fn_3 = lambda round_no: f"consolidated_decisions_round{round_no}"
        self.upload3 = RoundUploadWidget(
            title="Consolidated Decision File",
            required_map=required_map_3,
            table_name_fn=table_name_fn_3
        )
        self.layout.addWidget(self.upload3)

        # ------------------ Action Buttons ------------------
        btn_layout = QHBoxLayout()
        # Run Round 1 Allocation (can replace with dynamic round logic later)
        self.round_btn = QPushButton("Run Round Allocation")
        self.round_btn.clicked.connect(self.run_round)
        btn_layout.addWidget(self.round_btn)

        # Download offers button
        self.download_btn = QPushButton("Download Offers")
        self.download_btn.clicked.connect(self.download_current_round_offers)
        btn_layout.addWidget(self.download_btn)

        # Reset files & DB button
        self.reset_btn = QPushButton("Reset Uploaded Files")
        self.reset_btn.clicked.connect(self.reset_round)
        btn_layout.addWidget(self.reset_btn)

        self.layout.addLayout(btn_layout)

    # ------------------ Functions ------------------
    def get_current_round(self):
        """Return selected round as int"""
        return int(self.round_combo.currentText())

    def refresh_rounds(self):
        """Populate dropdown based on already generated rounds"""
        self.round_combo.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        # Check max round in offers table
        # cursor.execute("SELECT MAX(round_no) FROM offers")
        # max_round = cursor.fetchone()[0]
        # conn.close()
        # Check if offers table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='offers'")
        if cursor.fetchone() is None:
            max_round = 0  # no offers generated yet
        else:
            cursor.execute("SELECT MAX(round_no) FROM offers")
            max_round = cursor.fetchone()[0] or 0

        conn.close()
        start_round = 1
        end_round = (max_round + 1) if max_round else 1
        end_round = min(end_round, self.total_rounds)
        for r in range(start_round, end_round + 1):
            self.round_combo.addItem(str(r))

    def run_round(self):
        """Placeholder: call your round allocation logic"""
        round_no = self.get_current_round()
        # Here you would implement Round 2, Round 3, etc. allocation logic
        run_round_1() if round_no == 1 else QMessageBox.information(self, "Info", f"Run round {round_no} logic here")
        # Refresh dropdown after allocation
        self.refresh_rounds()

    def download_current_round_offers(self):
        round_no = self.get_current_round()
        download_offers(round_no)

    def reset_round(self):
        """Delete uploaded files from DB and reset widgets"""
        round_no = self.get_current_round()
        for upload in [self.upload1, self.upload2, self.upload3]:
            table_name = upload.table_name_fn(round_no)
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
            conn.commit()
            conn.close()
            upload.reset_widget()
        QMessageBox.information(self, "Reset", f"Round {round_no} uploads and DB tables cleared!")