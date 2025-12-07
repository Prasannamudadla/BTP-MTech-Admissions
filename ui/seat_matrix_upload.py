# ui/seat_matrix_upload.py
import pandas as pd
from PySide6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from database import db_manager

class SeatMatrixUpload(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        self.setLayout(layout)

        self.upload_btn = QPushButton(" Upload Seat Matrix Excel")
        layout.addWidget(self.upload_btn)

        self.status = QLabel("")
        layout.addWidget(self.status)

    def upload_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Seat Matrix Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return

        try:
            df = pd.read_excel(path)
            expected = ["category", "set_seats", "seats_allocated", "seats_booked"]
            if not all(col in df.columns for col in expected):
                QMessageBox.warning(self, "Invalid Format", f"Excel must contain: {', '.join(expected)}")
                return
            
            df.fillna(0, inplace=True)  

            conn = db_manager.get_connection()
            cursor = conn.cursor()
            for _, row in df.iterrows():
                cursor.execute("""
                    INSERT OR REPLACE INTO seat_matrix (category, set_seats, seats_allocated, seats_booked)
                    VALUES (?, ?, ?, ?)
                """, (row["category"], int(row["set_seats"]), int(row["seats_allocated"]), int(row["seats_booked"])))
            conn.commit()
            conn.close()

            self.status.setText("Seat matrix uploaded successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            
    def reset_status(self): 
        """Clears the status message shown in this widget."""
        self.status.setText("")
        self.status.setStyleSheet("")