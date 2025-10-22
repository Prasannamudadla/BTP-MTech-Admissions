# ui/update_dialog.py
import sqlite3
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QGridLayout, QVBoxLayout, QLabel, QPushButton, QScrollArea
)

# Map UI labels -> DB columns (None => show "NULL")
FIELD_MAP = {
    "FullName":               "Full_Name",
    "ApplicationNumber":      "App_no",
    "COAP":                   "COAP",
    "Email":                  "Email",
    "MaxGateScore":           "MaxGATEScore_3yrs",
    "Gender":                 "Gender",
    "Category":               "Category",
    "EWS":                    "Ews",
    "PWD":                    "Pwd",
    "SSCper":                 "SSC_per",
    "HSSCper":                "HSSC_per",
    "DegreeCGPA8thSem":       "Degree_CGPA_8th",
    "Offered":                None,     # not in candidates -> NULL
    "Accepted":               None,
    "OfferCat":               None,
    "isOfferPwd":             None,
    "OfferedRound":           None,
    "RetainRound":            None,
    "RejectOrAcceptRound":    None,
}

class UpdateDialog(QDialog):
    def __init__(self, db_path: str | Path, coap_id: str, parent: QWidget = None):
        super().__init__(parent)
        self.setWindowTitle(f"Candidate Details — {coap_id}")
        self.resize(720, 520)
        self.db_path = Path(db_path)
        self.coap_id = coap_id

        root = QVBoxLayout(self)

        # scrollable content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        grid = QGridLayout(content)
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(12)
        scroll.setWidget(content)

        # load data
        data = self._load_record()

        # build cards like the screenshot (label on top, value below, boxed)
        row, col = 0, 0
        for label, colname in FIELD_MAP.items():
            value = "NULL"
            if colname:
                v = data.get(colname)
                value = "NULL" if (v is None or v == "") else str(v)

            card = self._make_card(label, value)
            grid.addWidget(card, row, col)
            col += 1
            if col == 3:  # 3 cards per row feels right visually
                col = 0
                row += 1

        # close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)

        root.addWidget(scroll)
        root.addWidget(close_btn, 0, Qt.AlignRight)

    def _connect(self) -> sqlite3.Connection:
        if not self.db_path.exists():
            raise FileNotFoundError(f"Database not found: {self.db_path}")
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _load_record(self) -> dict:
        """Fetch entire row for this COAP ID."""
        try:
            conn = self._connect()
            cur = conn.execute("PRAGMA table_info(candidates)")
            cols = [r[1] for r in cur.fetchall()]

            cur = conn.execute(f"""
                SELECT {", ".join(cols)}
                FROM candidates
                WHERE COAP = ?
                LIMIT 1
            """, (self.coap_id,))
            row = cur.fetchone()
            conn.close()
            return dict(row) if row else {}
        except Exception:
            return {}

    def _make_card(self, label: str, value: str) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        title = QLabel(label)
        title.setStyleSheet("font-weight:600; color:#333;")
        val = QLabel(value)
        val.setStyleSheet("color:#111; background:#fafafa; border:1px solid #ddd; padding:8px; border-radius:6px;")
        v.addWidget(title)
        v.addWidget(val)
        return w