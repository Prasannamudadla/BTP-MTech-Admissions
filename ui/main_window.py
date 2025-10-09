# main_window.py
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTabWidget, QPushButton, QFileDialog, QLabel, QTableWidget, QTableWidgetItem, QScrollArea, QGroupBox, QPushButton, QToolBox, QHBoxLayout
)
import pandas as pd
from database import db_manager  # your module with get_connection()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MTech Offers Automation")
        self.resize(900, 600)

        # Create tab widget
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Create tabs
        self.init_tab = QWidget()
        # self.seat_matrix_tab = QWidget()
        # self.rounds_tab = QWidget()

        # self.tabs.addTab(self.init_tab, "Initialization")
        # self.tabs.addTab(self.seat_matrix_tab, "Seat Matrix")
        # self.tabs.addTab(self.rounds_tab, "Rounds")

        # # Setup Initialization tab
        # self.setup_init_tab()
        # self.seat_matrix_tab = SeatMatrixTab()
        # self.tabs.addTab(self.seat_matrix_tab, "Seat Matrix")
        # #self.setup_rounds_tab()       # Add this line
        self.rounds_tab = QWidget()

        self.tabs.addTab(self.init_tab, "Initialization")
        self.tabs.addTab(QWidget(), "Seat Matrix")  # placeholder will be replaced below
        self.tabs.addTab(self.rounds_tab, "Rounds")

        self.setup_init_tab()

        # Properly initialize seat matrix
        self.seat_matrix_tab = SeatMatrixTab()
        self.tabs.removeTab(1)  # remove placeholder
        self.tabs.insertTab(1, self.seat_matrix_tab, "Seat Matrix")

        # Now setup rounds tab
        self.setup_rounds_tab()
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
            table.itemChanged.connect(self.handle_item_changed)

            self.toolbox.addItem(table, section)
            self.tables[section] = table

    def handle_item_changed(self, item):
        """Automatically update 'Seats Allocated' when 'Set Seats' changes."""
        if item.column() == 0:  # Set Seats column
            try:
                new_value = int(item.text())
                table = item.tableWidget()
                table.item(item.row(), 1).setText(str(new_value))  # copy to Seats Allocated
            except ValueError:
                pass  # ignore invalid (non-integer) entries

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