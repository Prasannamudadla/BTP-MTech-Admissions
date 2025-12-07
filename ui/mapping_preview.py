from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QHBoxLayout, QLabel, QComboBox
)

class MappingPreviewDialog(QDialog):
    def __init__(self, mapping, src_columns, required_targets=None, parent=None):
        """
        mapping: dict (db_column -> suggested excel source)
        src_columns: Excel columns list
        required_targets: list of DB columns to show in UI
        """
        super().__init__(parent)
        self.setWindowTitle("Confirm Column Mapping")
        self.resize(700, 400)

        self.all_mapping = mapping.copy()
        self.src_columns = list(src_columns)

        # If not provided, show all mapping (fallback)
        if required_targets is None:
            required_targets = list(mapping.keys())

        # Only show these required DB fields in UI
        self.display_targets = required_targets  

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Select correct Excel column for each required database field:"))

        # UI table (only required rows)
        self.table = QTableWidget(len(self.display_targets), 2)
        self.table.setHorizontalHeaderLabels(["DB Column", "Mapped Column"])

        for i, tgt in enumerate(self.display_targets):
            # Pre-filled suggestion
            suggested_src = self.all_mapping.get(tgt)

            # DB column name
            self.table.setItem(i, 0, QTableWidgetItem(tgt))

            # Dropdown of Excel columns
            combo = QComboBox()
            combo.addItem("")  # Allow empty
            for col in self.src_columns:
                combo.addItem(col)

            # Auto-select suggested column
            if suggested_src:
                combo.setCurrentText(suggested_src)

            self.table.setCellWidget(i, 1, combo)

        layout.addWidget(self.table)

        # OK / Cancel buttons
        btns = QHBoxLayout()
        ok = QPushButton("OK")
        cancel = QPushButton("Cancel")
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        btns.addWidget(ok)
        btns.addWidget(cancel)
        layout.addLayout(btns)

    def get_final_mapping(self):
        """
        Return mapping for ALL DB columns:
        - required targets: use user selection
        - non-required: keep autosuggested values
        """
        final = self.all_mapping.copy()

        for r in range(self.table.rowCount()):
            tgt = self.table.item(r, 0).text()   # DB field
            combo = self.table.cellWidget(r, 1)
            choice = combo.currentText().strip()
            final[tgt] = choice if choice else None

        return final