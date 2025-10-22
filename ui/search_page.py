# # search_page_final.py
# import sqlite3
# from pathlib import Path
# from typing import Dict, Optional, Tuple

# from PySide6.QtCore import Qt
# from PySide6.QtGui import QBrush, QColor
# from PySide6.QtWidgets import (
#     QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QHBoxLayout, QVBoxLayout,
#     QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView
# )


# class SearchPage(QWidget):
#     """
#     Local desktop 'Search' page (no web calls).
#     Inputs: COAP ID (text), Category (GEN/OBC/SC/ST), Gender (Male/Female)
#     Data source: sqlite3 database at mtech_offers.db
#     Output table columns:
#         COAP ID | Application Number | Category | Gender | Max Gate Score | PWD | EWS
#     """
#     def __init__(self, db_path: Optional[str | Path] = None, parent: Optional[QWidget] = None):
#         super().__init__(parent)
#         self.setObjectName("SearchPage")

#         # DB path
#         self.db_path = Path(db_path) if db_path else Path.cwd() / "mtech_offers.db"

#         # ---- Top filter row ----
#         self.coap_input = QLineEdit()
#         self.coap_input.setPlaceholderText("COAP ID")
#         self.coap_input.setClearButtonEnabled(True)
#         self.coap_input.setMinimumWidth(200)

#         self.category_combo = QComboBox()
#         self.category_combo.addItem("Category")
#         self.category_combo.addItems(["GEN", "OBC", "SC", "ST"])

#         self.gender_combo = QComboBox()
#         self.gender_combo.addItem("Gender")
#         self.gender_combo.addItems(["Male", "Female"])

#         self.find_btn = QPushButton("🔍 FIND DETAILS")
#         self.find_btn.setDefault(True)
#         self.find_btn.clicked.connect(self._on_find_clicked)

#         top = QHBoxLayout()
#         top.addWidget(self.coap_input, 2)
#         top.addWidget(self.category_combo, 1)
#         top.addWidget(self.gender_combo, 1)
#         top.addWidget(self.find_btn, 0, Qt.AlignLeft)
#         top.addStretch(1)

#         # ---- Results table ----
#         self.table = QTableWidget(0, 7, self)
#         self.table.setHorizontalHeaderLabels([
#             "COAP ID", "Application Number", "Category", "Gender",
#             "Max Gate Score", "PWD", "EWS"
#         ])
#         self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
#         self.table.horizontalHeader().setHighlightSections(False)
#         self.table.verticalHeader().setVisible(False)
#         self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
#         self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
#         self.table.setSelectionMode(QAbstractItemView.SingleSelection)
#         self.table.setAlternatingRowColors(True)
#         self.table.setShowGrid(True)

#         # empty state
#         self.empty_label = QLabel("No Results Found")
#         self.empty_label.setAlignment(Qt.AlignCenter)
#         self.empty_label.setStyleSheet("color:#666; font-size:14px; padding:16px;")

#         wrapper = QVBoxLayout(self)
#         wrapper.addLayout(top)
#         wrapper.addWidget(self.table)
#         wrapper.addWidget(self.empty_label)
#         self._set_empty(True)

#         # Enter to search
#         self.coap_input.returnPressed.connect(self._on_find_clicked)
#         self.category_combo.activated.connect(lambda *_: self._focus_find())
#         self.gender_combo.activated.connect(lambda *_: self._focus_find())

#     # ---------- UI helpers ----------
#     def _focus_find(self):
#         self.find_btn.setFocus(Qt.OtherFocusReason)

#     def _set_empty(self, is_empty: bool):
#         self.table.setVisible(not is_empty)
#         self.empty_label.setVisible(is_empty)

#     # ---------- DB helpers ----------
#     def _connect(self) -> sqlite3.Connection:
#         if not self.db_path.exists():
#             raise FileNotFoundError(f"Database not found: {self.db_path}")
#         conn = sqlite3.connect(str(self.db_path))
#         conn.row_factory = sqlite3.Row
#         return conn

#     # ---------- Actions ----------
#     def _on_find_clicked(self):
#         coap_id = self.coap_input.text().strip()
#         category = self.category_combo.currentText().strip()
#         gender = self.gender_combo.currentText().strip()

#         errs = []
#         if not coap_id:
#             errs.append("COAP ID")
#         if category not in {"GEN", "OBC", "SC", "ST"}:
#             errs.append("Category")
#         if gender not in {"Male", "Female"}:
#             errs.append("Gender")

#         if errs:
#             self._show_error_row(f"Please provide: {', '.join(errs)}")
#             return

#         try:
#             conn = self._connect()
#             cur = conn.execute("""
#                 SELECT COAP AS coap_id, App_no AS application_number, Category AS category,
#                        Gender AS gender, MaxGATEScore_3yrs AS max_gate_score,
#                        Pwd AS pwd, Ews AS ews
#                 FROM candidates
#                 WHERE COAP = ? AND Category = ? AND Gender = ?
#                 LIMIT 50;
#             """, (coap_id, category, gender))
#             rows = list(cur)
#             conn.close()
#         except Exception as e:
#             self._show_error_row(f"DB error: {e}")
#             return

#         if not rows:
#             self._set_empty(True)
#             return

#         self._populate_table(rows)

#     def _show_error_row(self, message: str):
#         self.table.setRowCount(0)
#         self._set_empty(False)
#         self.table.setRowCount(1)
#         item = QTableWidgetItem(message)
#         item.setForeground(QBrush(QColor('gray')))
#         self.table.setSpan(0, 0, 1, self.table.columnCount())
#         self.table.setItem(0, 0, item)

#     def _populate_table(self, rows: list[sqlite3.Row]):
#         self.table.clearSpans()
#         self.table.setRowCount(0)
#         self._set_empty(False)

#         for row in rows:
#             r = self.table.rowCount()
#             self.table.insertRow(r)

#             def _set(c, v):
#                 item = QTableWidgetItem("" if v is None else str(v))
#                 item.setTextAlignment(Qt.AlignCenter)
#                 self.table.setItem(r, c, item)

#             _set(0, row["coap_id"])
#             _set(1, row["application_number"])
#             _set(2, row["category"])
#             _set(3, row["gender"])
#             _set(4, row["max_gate_score"])
#             _set(5, row["pwd"])
#             _set(6, row["ews"])

#         self.table.resizeColumnsToContents()
#         self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
#         self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
#         self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)


# # ---------- Quick test ----------
# # if __name__ == "__main__":
# #     from PySide6.QtWidgets import QApplication, QMainWindow
# #     import sys

# #     app = QApplication(sys.argv)
# #     win = QMainWindow()
# #     win.setWindowTitle("M.Tech Admissions — Search")
# #     search = SearchPage()
# #     win.setCentralWidget(search)
# #     win.resize(1000, 520)
# #     win.show()
# #     sys.exit(app.exec())
# ui/search_page.py
import sqlite3
from pathlib import Path
from typing import Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QHBoxLayout, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QToolButton
)

class SearchPage(QWidget):
    """
    Search by COAP + Category + Gender (exact).
    Shows: COAP | App_no | Category | Gender | MaxGATEScore_3yrs | Pwd | Ews
    Emits updateRequested(dict) when UPDATE is clicked (dict contains coap/category/gender/app_no etc.)
    """
    updateRequested = Signal(dict)

    def __init__(self, db_path: Optional[str | Path] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("SearchPage")
        self.db_path = Path(db_path) if db_path else Path.cwd() / "mtech_offers.db"

        # ---- Filters ----
        self.coap_input = QLineEdit()
        self.coap_input.setPlaceholderText("COAP ID")
        self.coap_input.setClearButtonEnabled(True)
        self.coap_input.setMinimumWidth(200)

        self.category_combo = QComboBox()
        self.category_combo.addItems(["Category", "GEN", "OBC", "SC", "ST"])

        self.gender_combo = QComboBox()
        self.gender_combo.addItems(["Gender", "Male", "Female"])

        self.find_btn = QPushButton("🔍  FIND DETAILS")
        self.find_btn.setDefault(True)
        self.find_btn.clicked.connect(self._on_find_clicked)

        top = QHBoxLayout()
        top.addWidget(self.coap_input, 2)
        top.addWidget(self.category_combo, 1)
        top.addWidget(self.gender_combo, 1)
        top.addWidget(self.find_btn, 0, Qt.AlignLeft)
        top.addStretch(1)

        # ---- Results table ----
        self.table = QTableWidget(0, 8, self)
        self.table.setHorizontalHeaderLabels([
            "COAP ID", "Application Number", "Category", "Gender",
            "Max Gate Score", "PWD", "EWS", "Action"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)

        # Empty state
        self.empty_label = QLabel("No Results Found")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet("color:#666; font-size:14px; padding:16px;")

        wrapper = QVBoxLayout(self)
        wrapper.addLayout(top)
        wrapper.addWidget(self.table)
        wrapper.addWidget(self.empty_label)
        self._set_empty(True)

        self.coap_input.returnPressed.connect(self._on_find_clicked)
        self.category_combo.activated.connect(lambda *_: self.find_btn.setFocus(Qt.OtherFocusReason))
        self.gender_combo.activated.connect(lambda *_: self.find_btn.setFocus(Qt.OtherFocusReason))

    # ---------- Helpers ----------
    def _set_empty(self, is_empty: bool):
        # self.table.setVisible(not is_empty := is_empty is False)  # small flair to keep line short
        # readability fallback
        self.table.setVisible(not is_empty)
        self.empty_label.setVisible(is_empty)
        
    #     def _set_empty(self, is_empty: bool):
    #     self.table.setVisible(not is_empty)
    #     self.empty_label.setVisible(is_empty)

    def _connect(self) -> sqlite3.Connection:
        if not self.db_path.exists():
            raise FileNotFoundError(f"Database not found: {self.db_path}")
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    # ---------- Actions ----------
    def _on_find_clicked(self):
        coap_id = self.coap_input.text().strip()
        category = self.category_combo.currentText().strip()
        gender = self.gender_combo.currentText().strip()

        errs = []
        if not coap_id:
            errs.append("COAP ID")
        if category not in {"GEN", "OBC", "SC", "ST"}:
            errs.append("Category")
        if gender not in {"Male", "Female"}:
            errs.append("Gender")

        if errs:
            self._show_error_row(f"Please provide: {', '.join(errs)}")
            return

        try:
            conn = self._connect()
            cur = conn.execute("""
                SELECT
                    COAP               AS coap_id,
                    App_no             AS application_number,
                    Category           AS category,
                    Gender             AS gender,
                    MaxGATEScore_3yrs  AS max_gate_score,
                    Pwd                AS pwd,
                    Ews                AS ews
                FROM candidates
                WHERE COAP = ? AND Category = ? AND Gender = ?
                LIMIT 50;
            """, (coap_id, category, gender))
            rows = list(cur)
            conn.close()
        except Exception as e:
            self._show_error_row(f"DB error: {e}")
            return

        if not rows:
            self._set_empty(True)
            return

        self._populate_table(rows)

    def _show_error_row(self, message: str):
        self.table.setRowCount(0)
        self._set_empty(False)
        self.table.setRowCount(1)
        item = QTableWidgetItem(message)
        item.setForeground(QBrush(QColor('gray')))
        self.table.setSpan(0, 0, 1, self.table.columnCount())
        self.table.setItem(0, 0, item)

    def _populate_table(self, rows: list[sqlite3.Row]):
        self.table.clearSpans()
        self.table.setRowCount(0)
        self._set_empty(False)

        for row in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)

            def _set(c, v):
                item = QTableWidgetItem("" if v is None else str(v))
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r, c, item)

            _set(0, row["coap_id"])
            _set(1, row["application_number"])
            _set(2, row["category"])
            _set(3, row["gender"])
            _set(4, row["max_gate_score"])
            _set(5, row["pwd"])
            _set(6, row["ews"])

            # UPDATE button
            btn = QToolButton()
            btn.setText("UPDATE")
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _, rr=dict(row): self.updateRequested.emit(rr))
            self.table.setCellWidget(r, 7, btn)

        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)