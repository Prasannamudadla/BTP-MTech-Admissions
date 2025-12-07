import sqlite3
from pathlib import Path
from typing import Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QHBoxLayout, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QToolButton, QMessageBox
)

class SearchPage(QWidget):
    """
    Search by COAP ID (partial match supported).
    Shows: COAP | App_no | Category | Gender | MaxGATEScore_3yrs | Pwd | Ews
    Emits updateRequested(dict) when UPDATE is clicked.
    """
    updateRequested = Signal(dict)

    def __init__(self, db_path: Optional[str | Path] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("SearchPage")
        self.db_path = Path(db_path) if db_path else Path.cwd() / "mtech_offers.db"

        # ---- Filters ----
        self.coap_input = QLineEdit()
        self.coap_input.setPlaceholderText("Enter COAP ID")
        self.coap_input.setClearButtonEnabled(True)
        self.coap_input.setMinimumWidth(300)
        
        # Removed self.category_combo and self.gender_combo

        self.find_btn = QPushButton("SEARCH COAP")
        self.find_btn.setDefault(True)
        self.find_btn.clicked.connect(self._on_find_clicked)

        top = QHBoxLayout()
        top.addWidget(QLabel("COAP ID:"))
        top.addWidget(self.coap_input, 2)
        # Removed category/gender widgets
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
        self.empty_label = QLabel("Enter a COAP ID and click Search.")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet("color:#666; font-size:14px; padding:16px;")

        wrapper = QVBoxLayout(self)
        wrapper.addLayout(top)
        wrapper.addWidget(self.table)
        wrapper.addWidget(self.empty_label)
        self._set_empty(True)

        self.coap_input.returnPressed.connect(self._on_find_clicked)

    # ---------- Helpers ----------
    def _set_empty(self, is_empty: bool, message: Optional[str] = None):
        self.table.setVisible(not is_empty)
        self.empty_label.setVisible(is_empty)
        if message:
            self.empty_label.setText(message)

    def _connect(self) -> sqlite3.Connection:
        if not self.db_path.exists():
            # Show error box if DB is missing
            QMessageBox.critical(self, "Database Error", f"Database not found: {self.db_path}")
            raise FileNotFoundError(f"Database not found: {self.db_path}")
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    # ---------- Actions ----------
    def _on_find_clicked(self):
        coap_id = self.coap_input.text().strip()
        
        if not coap_id:
            self._set_empty(True, "Please enter a COAP ID to search.")
            return

        try:
            conn = self._connect()
            # SQL query now uses LIKE for partial matching
            cur = conn.execute("""
                SELECT
                    COAP AS coap_id,
                    App_no AS application_number,
                    Category AS category,
                    Gender AS gender,
                    MaxGATEScore_3yrs AS max_gate_score,
                    Pwd AS pwd,
                    Ews AS ews
                FROM candidates
                WHERE COAP LIKE ?
                ORDER BY COAP
                LIMIT 50;
            """, (f"%{coap_id}%",)) # Wrap input in % for LIKE query
            rows = list(cur)
            conn.close()
        except Exception as e:
            self._show_error_row(f"DB error: {e}")
            return

        if not rows:
            self._set_empty(True, f"No candidates found matching COAP ID: '{coap_id}'")
            return

        self._populate_table(rows)

    def _show_error_row(self, message: str):
        self.table.setRowCount(0)
        self._set_empty(False)
        self.table.setRowCount(1)
        item = QTableWidgetItem(message)
        item.setForeground(QBrush(QColor('red')))
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
            # Pass the entire row dictionary when the button is clicked
            btn.clicked.connect(lambda _, rr=dict(row): self.updateRequested.emit(rr))
            self.table.setCellWidget(r, 7, btn)

        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)