# main_window.py
import sqlite3
from PySide6.QtCore import Qt,Signal
from PySide6.QtWidgets import (
    QApplication, QMainWindow,QDialog, QWidget, QVBoxLayout, QMessageBox, 
    QTabWidget, QPushButton, QFileDialog, QLabel, QComboBox, QTableWidget, 
    QTableWidgetItem, QScrollArea, QGroupBox, QToolBox, QHBoxLayout,QToolButton,QSizePolicy
)
import re, difflib, json
import numpy as np
from ui.update_dialog import UpdateDialog
from ui.rounds_manager import run_round, download_offers, upload_round_decisions 
import pandas as pd
import re
from database import db_manager 
from ui.round_upload_widget import RoundUploadWidget
from ui.search_page import SearchPage
from ui.seat_matrix_upload import SeatMatrixUpload
import datetime

def dynamic_required_gate_cols():
    current = datetime.datetime.now().year % 100 
    years = [current, current - 1, current - 2]

    cols = []
    for y in years:
        cols += [
            f"GATE{y}Score",
            f"GATE{y}RollNo"
        ]
    return cols

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
] + dynamic_required_gate_cols() + [
    "HSSC_per",
    "SSC_per",
    "Degree_Per_8th",
    "Degree_CGPA_8th"
]

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
        self.total_rounds = 11
        
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
        self.rounds_tab.roundsRefreshed.connect(self.seat_matrix_tab.load_matrix)
        # Search tab
        self.search_tab = SearchPage(db_path="mtech_offers.db")
        self.search_tab.updateRequested.connect(self.open_update_page) 
        self.tabs.addTab(self.search_tab, "Search")

    def setup_init_tab(self):
        layout = QVBoxLayout()
        self.init_tab.setLayout(layout)
        
        self.status_label = QLabel("") 
        layout.addWidget(self.status_label)
        # Upload Excel button
        self.upload_btn = QPushButton("Upload Applicants Excel")
        self.upload_btn.clicked.connect(self.upload_excel)
        layout.addWidget(self.upload_btn)
        
        layout.addSpacing(40)
        
        self.reset_db_btn = QPushButton("Reset All Database Data")
        self.reset_db_btn.clicked.connect(self.reset_all_data)
        layout.addWidget(self.reset_db_btn)
        
        # --- Final step ---
        # Call a new method to check DB state and update UI labels/buttons
        self.update_init_tab_state()

    def reset_all_data(self):
        """Handles the full deletion and re-initialization of the database."""
        
        reply = QMessageBox.question(self, 'Confirm Full Reset',
            "**WARNING:** Are you absolutely sure you want to **DELETE ALL DATA** (Candidates, Seat Matrix, Rounds) and start over? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            
        if reply == QMessageBox.StandardButton.No:
            return

        try:
            db_manager.reset_db_data()
            QMessageBox.information(self, "Reset Successful", 
                                    "The database has been completely deleted and re-initialized. You must now upload the Applicants Excel and Seat Matrix again.")
            
            # Refresh all relevant UI components after a full reset
            self.update_init_tab_state()
            self.seat_matrix_tab.reset_upload_status()
            self.seat_matrix_tab.load_matrix() # Reload the empty seat matrix table
            self.rounds_tab.refresh_rounds() # Reset rounds
            
        except Exception as e:
            QMessageBox.critical(self, "Reset Error", f"A critical error occurred during database reset:\n{e}")
            
    def check_db_state(self):
        """Checks if candidates or seat matrix data exists in the database."""
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        
        # Check candidates table
        try:
            cursor.execute("SELECT COUNT(*) FROM candidates")
            candidates_count = cursor.fetchone()[0]
        except:
            candidates_count = 0
            
        # Check seat_matrix table
        try:
            cursor.execute("SELECT COUNT(*) FROM seat_matrix")
            matrix_count = cursor.fetchone()[0]
        except:
            matrix_count = 0
            
        conn.close()
        
        return candidates_count, matrix_count

    def update_init_tab_state(self):
        """Updates the Initialization tab UI based on DB content."""
        candidates_count, matrix_count = self.check_db_state()
        if candidates_count == 0:
            # No data present, hide the reset button.
            self.upload_btn.show()
            self.reset_db_btn.hide()
            self.reset_db_btn.setEnabled(False)
            
            # Inform the user that data needs to be uploaded.
            status_text = "No candidate data found. Please upload candidates file ."
            self.status_label.setStyleSheet("color: black; font-weight: bold;")
            
        else:
            self.upload_btn.hide()
            # Data exists, enable and show the reset button.
            self.reset_db_btn.show()
            self.reset_db_btn.setEnabled(True)
            
            # Inform the user about the current data count.
            status_text = (
                f"Database successfully initialized "
            )
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
            
        self.status_label.setText(status_text)
             
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

            current = datetime.datetime.now().year % 100  
            years = [current, current - 1, current - 2]

            for y in years:
                synonyms.update({
                    f"gate {y} score": f"GATE{y}Score",
                    f"gate {y} marks": f"GATE{y}Score",
                    f"gate {y} roll": f"GATE{y}RollNo",
                    f"gate {y} roll number": f"GATE{y}RollNo",
                    f"gate {y} rollno": f"GATE{y}RollNo",
                })
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

            self.update_init_tab_state()
            self.status_label.setText("Excel uploaded & mapped successfully!")

        except Exception as e:
            self.status_label.setText(f"Error: {e}")

class SeatMatrixTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout(self)

        # top: Upload widget (uses your existing seat_matrix_upload module)
        self.upload_widget = SeatMatrixUpload()
        layout.addWidget(self.upload_widget)

        # Connect the upload button so that after upload completes we reload the UI.
        # Note: SeatMatrixUpload.upload_excel does the DB writing and sets a status label.
        # We call load_matrix() afterwards to refresh the visible tables.
        self.upload_widget.upload_btn.clicked.connect(self._on_upload_clicked)

        # Separator / info
        info = QLabel("Or edit seat counts below and click Save Seat Matrix")
        layout.addWidget(info)
        
        # Collapsible sections using QToolBox
        # self.toolbox = QToolBox()
        # layout.addWidget(self.toolbox)
        self.accordion = QWidget()
        self.accordion_layout = QVBoxLayout(self.accordion)
        self.accordion_layout.setContentsMargins(0, 0, 0, 0)
        self.accordion_layout.setSpacing(6)
        layout.addWidget(self.accordion)
        
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
        layout.addLayout(btn_layout)

        self.is_rounds_started = False
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
    def create_sections(self):
        """
        Build a simple accordion: a QToolButton header (checkable) + a content widget (table).
        All sections start collapsed. Clicking a header toggles its content. When a header is
        opened other sections will close (accordion behavior). Click same header again to collapse.
        """
        # clear any previous content (safe if re-run)
        for i in reversed(range(self.accordion_layout.count())):
            w = self.accordion_layout.itemAt(i).widget()
            if w:
                w.setParent(None)

        # containers to keep references
        self.header_buttons = {}  # section -> QToolButton
        self.content_widgets = {}  # section -> QWidget (contains the QTableWidget)
        self.tables = {}  # keep your existing mapping section -> QTableWidget

        for section, subcats in self.categories.items():
            # Header (toggle button)
            header = QToolButton()
            header.setText(section)
            header.setCheckable(True)
            header.setChecked(False)  # collapsed initially
            header.setToolButtonStyle(Qt.ToolButtonTextOnly)
            header.setStyleSheet("""
                QToolButton {
                    text-align: center;
                    padding: 10px;
                    font-weight: 600;
                    border: 1px solid #e6e6e6;
                    background: #ffffff;
                }
                QToolButton:checked {
                    background: #fafafa;
                }
            """)
            header.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
            self.accordion_layout.addWidget(header)

            # Content container (will hold the QTableWidget)
            content = QWidget()
            content_layout = QVBoxLayout(content)
            content_layout.setContentsMargins(8, 8, 8, 8)
            content_layout.setSpacing(6)

            table = QTableWidget()
            table.setRowCount(len(subcats))
            table.setColumnCount(3)
            table.setHorizontalHeaderLabels(["Set Seats", "Seats Allocated", "Seats Booked"])

            # Fill rows / vertical headers
            for row_idx, sub in enumerate(subcats):
                header_item = QTableWidgetItem(sub)
                header_item.setFlags(header_item.flags() & ~Qt.ItemIsEditable)
                table.setVerticalHeaderItem(row_idx, header_item)
                for col in range(3):
                    val = QTableWidgetItem("0")
                    if col != 0:
                        val.setFlags(val.flags() & ~Qt.ItemIsEditable)
                    table.setItem(row_idx, col, val)

            # Layout / sizing for neatness
            table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            table.setMaximumHeight(180)
            # column resize behaviour (optional tweak)
            try:
                from PySide6.QtWidgets import QHeaderView
                table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
                table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
                table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            except Exception:
                pass

            content_layout.addWidget(table)
            content.setVisible(False)  # collapsed initially
            self.accordion_layout.addWidget(content)

            # store references
            self.header_buttons[section] = header
            self.content_widgets[section] = content
            self.tables[section] = table

            # connect signals: toggle content when header toggles
            header.toggled.connect(lambda checked, s=section: self._on_header_toggled(s, checked))

            # when user edits 'Set Seats' (col 0) update header label
            table.itemChanged.connect(lambda item, _table=table, _section=section:
                                    self.on_table_item_changed(item, _table, _section))

        # spacer to push Save button to bottom if desired
        self.accordion_layout.addStretch()
        # update header labels based on initial values (zeroes)
        self.update_all_section_labels()
        
    def _on_header_toggled(self, section: str, checked: bool):
        """
        When a header is toggled:
        - if checked=True: show its content and uncheck/close other sections (accordion behaviour)
        - if checked=False: hide its content
        """
        # If opening this section, close others
        if checked:
            for sec, btn in self.header_buttons.items():
                if sec != section:
                    # block signals to avoid recursion
                    btn.blockSignals(True)
                    btn.setChecked(False)
                    btn.blockSignals(False)
                    self.content_widgets[sec].setVisible(False)
            # show this content
            self.content_widgets[section].setVisible(True)
        else:
            # hide this content
            self.content_widgets[section].setVisible(False)
            
    def _on_toolbox_current_changed(self, index: int):
        # Show only the currently selected widget; hide the rest
        for i in range(self.toolbox.count()):
            w = self.toolbox.widget(i)
            w.setVisible(i == index)
            
    def on_table_item_changed(self, item: QTableWidgetItem, table: QTableWidget, section: str):
        """Update the section label when Set Seats (col 0) changes."""
        if item is None:
            return
        if item.column() != 0:
            return
        # don't react to programmatic changes
        if table.signalsBlocked():
            return
        self.update_section_label(section, table)
              
    def update_section_label(self, section: str, table: QTableWidget):
        """
        Build a compact summary string showing only subcategories with seats > 0.
        If nothing has seats, just show the section name.
        Also truncate very long headers with ellipsis.
        """
        parts = []
        for r in range(table.rowCount()):
            vheader = table.verticalHeaderItem(r)
            if vheader is None:
                continue
            name = vheader.text().strip()
            item = table.item(r, 0)
            try:
                seats = int(item.text()) if item and item.text() != "" else 0
            except Exception:
                seats = 0
            # Only include non-zero seats in the header summary
            if seats > 0:
                parts.append(f"{name} ({seats})")

        # If there are non-zero items show them, otherwise just the section name
        if parts:
            summary = f"{section} : " + "  ".join(parts)
        else:
            summary = section

        # Truncate to keep header compact (adjust length as needed)
        MAX_LEN = 80
        if len(summary) > MAX_LEN:
            summary = summary[:MAX_LEN - 3] + "..."

        btn = self.header_buttons.get(section)
        if btn:
            btn.setText(summary)
            
    def update_all_section_labels(self):
        """Update labels for all created tables."""
        for section, table in self.tables.items():
            table.blockSignals(True)
            self.update_section_label(section, table)
            table.blockSignals(False)
                   
    def reset_upload_status(self): # <--- NEW METHOD
        """Resets the status message of the upload widget."""
        try:
            # Assuming SeatMatrixUpload has a public method or attribute to clear the status.
            # We'll assume the internal status label is named 'status_label' and we can set its text.
            # If your SeatMatrixUpload class has a different way to clear the message, adjust this line.
            self.upload_widget.reset_status()
        except AttributeError:
            # Fallback if status_label is private or named differently.
            # If the upload widget is a simple QWidget, this might not work.
            # If this still fails, you'll need to modify the SeatMatrixUpload class itself.
            pass
    
    def check_offers_exist(self):
        """Checks if any offers exist in the 'offers' table."""
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        
        # Check if 'offers' table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='offers'")
        if cursor.fetchone() is None:
            conn.close()
            return False

        # Check if offers exist
        cursor.execute("SELECT COUNT(*) FROM offers")
        count = cursor.fetchone()[0]
        conn.close()
        return count > 0
    
    def load_matrix(self):
        """Load data from seat_matrix table into GUI."""
        self.is_rounds_started = self.check_offers_exist()
        is_editable = not self.is_rounds_started
        # 1. Update the Upload Widget
        self.upload_widget.upload_btn.setEnabled(is_editable)
        self.upload_widget.setEnabled(is_editable)

        # 2. Update the Manual Save Button
        self.save_btn.setEnabled(is_editable)
        
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT category, set_seats, seats_allocated, seats_booked FROM seat_matrix")
            data = cursor.fetchall()
        except Exception:
            data = []
        finally:
            conn.close()
            
        for section, table in self.tables.items():
            for r in range(table.rowCount()):
                for c in range(table.columnCount()):
                    # Column 0 ('Set Seats') is editable, Columns 1 & 2 are not.
                    # We only need to reset the text value.
                    table.blockSignals(True)
                    table.item(r, c).setText("0")
                    table.blockSignals(False)
                    
                    # Set the 'Set Seats' column (index 0) editability
                    if c == 0:
                        item_flags = table.item(r, c).flags()
                        if is_editable:
                            # Enable: Keep it editable
                            table.item(r, c).setFlags(item_flags | Qt.ItemIsEditable)
                        else:
                            # Disable: Make it non-editable
                            table.item(r, c).setFlags(item_flags & ~Qt.ItemIsEditable)
                    # Columns 1 & 2 are always non-editable (set in create_sections)
                    
                    table.blockSignals(False)
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
        self.update_all_section_labels()
        

    def save_matrix(self):
        """Save data back to the database."""
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        for section, table in self.tables.items():
            for r in range(table.rowCount()):
                category = table.verticalHeaderItem(r).text()
                # defensive: ensure numeric parse
                try:
                    set_seats = int(table.item(r, 0).text())
                except Exception:
                    set_seats = 0
                try:
                    seats_allocated = int(table.item(r, 1).text())
                except Exception:
                    seats_allocated = 0
                try:
                    seats_booked = int(table.item(r, 2).text())
                except Exception:
                    seats_booked = 0

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
    roundsRefreshed = Signal()
    def __init__(self, total_rounds=11):
        super().__init__()
        self.total_rounds = total_rounds
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)
        
        # from PySide6.QtCore import Signal
        # self.roundsRefreshed = Signal()

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
        
        self.delete_round_btn = QPushButton("Reset Data") # NEW BUTTON
        self.delete_round_btn.clicked.connect(self.delete_round_data) # NEW method
        btn_layout.addWidget(self.delete_round_btn)

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
        self.roundsRefreshed.emit()
    
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
        # Round 1 (Ungenerated)
            self.upload_widget.setVisible(False)
            self.generate_btn.setEnabled(True)
            self.download_btn.setVisible(False)
            # self.reset_uploads_btn.setVisible(False)
            self.delete_round_btn.setVisible(False) # NEW

        elif is_current_round_run:
            # Past round already generated (e.g., you select Round 3, but Round 4 exists)
            self.upload_widget.setVisible(False)
            self.generate_btn.setEnabled(False)
            self.download_btn.setVisible(True)
            # self.reset_uploads_btn.setVisible(False) 
            self.delete_round_btn.setVisible(True) # NEW: Allow deleting offers for this past round

        elif round_no == 1 and is_current_round_run:
            # Round 1 (Generated)
            self.upload_widget.setVisible(False)
            self.generate_btn.setEnabled(False)
            self.download_btn.setVisible(True)
            # self.reset_uploads_btn.setVisible(False)
            self.delete_round_btn.setVisible(True) # NEW: Allow deleting offers for R1

        else: # is_next_round_upload is True (Round > 1, ready for upload)
            # Round > 1 → show upload section
            self.upload_widget.setVisible(True)
            self.generate_btn.setEnabled(True)
            self.download_btn.setVisible(False)
            # self.reset_uploads_btn.setVisible(True) # Show the reset button to clear uploads
            self.delete_round_btn.setVisible(False) # Hide the delete button

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
    # main_window.py (Inside RoundsWidget)

    def reset_uploads(self): # RENAMED from reset_round
        round_no = self.get_current_round()
        if round_no == 1:
            QMessageBox.warning(self, "Warning",
                                "Round 1 does not have decision uploads.")
            return

        # Clear UI for the next round's uploads
        self.upload_widget.goa_widget.reset_widget()
        self.upload_widget.other_widget.reset_widget()
        self.upload_widget.cons_widget.reset_widget()
        
        QMessageBox.information(
            self,
            "Uploads Cleared",
            f"File uploads have been cleared for Round {round_no}"
    )
    # main_window.py (Inside RoundsWidget)

    def delete_round_data(self):
        round_no = self.get_current_round()
        
        # 1. Confirmation
        reply = QMessageBox.question(self, 'Confirm Deletion',
            f"Are you sure you want to delete ALL GENERATED OFFERS for Round {round_no}? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                
        if reply == QMessageBox.StandardButton.No:
            return

        # 2. Delete offers from the main 'offers' table
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        try:
            # Delete the offers generated by run_round(round_no)
            cursor.execute("DELETE FROM offers WHERE round_no = ?", (round_no,))
            if round_no > 1:
                iit_goa_table = f"iit_goa_offers_round{round_no-1}"
                accepted_other_table = f"accepted_other_institute_round{round_no-1}"
                consolidated_table = f"consolidated_decisions_round{round_no-1}"
                cursor.execute(f"DELETE FROM {iit_goa_table}")
                cursor.execute(f"DELETE FROM {accepted_other_table}")
                cursor.execute(f"DELETE FROM {consolidated_table}")
            conn.commit()
            QMessageBox.information(self, "Success", f"All generated offers for Round {round_no} have been deleted.")

        except Exception as e:
            QMessageBox.critical(self, "DB Error", f"Could not delete round {round_no} offers:\n{e}")
            conn.rollback()
            return
        finally:
            conn.close()

        # 3. Refresh UI
        self.refresh_rounds() # Re-populates the dropdown based on MAX(round_no) in 'offers' table
        # The current round is now the new MAX + 1, so the UI should switch to the upload view for the next round
        self.update_ui_visibility()
        
