from PySide6.QtCore import QThread, Signal
import pandas as pd
import sqlite3
from database.db_manager import DB_NAME

class ExcelWorker(QThread):
    progress = Signal(str)
    finished = Signal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        self.progress.emit("Reading Excel file...")
        try:
            df = pd.read_excel(self.file_path)

            # Remove unnamed columns
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

            # Fill missing values
            df = df.where(pd.notnull(df), None)

            # Map Excel columns to DB columns
            column_mapping = {
                "Si NO": "Si_NO",
                "App no": "App_no",
                "Email": "Email",
                "Full Name": "Full_Name",
                "Adm cat": "Adm_cat",
                "Pwd": "Pwd",
                "Ews": "Ews",
                "Gender": "Gender",
                "Category": "Category",
                "COAP": "COAP",
                "GATE22RollNo": "GATE22RollNo",
                "GATE22Rank": "GATE22Rank",
                "GATE22Score": "GATE22Score",
                "GATE22Disc": "GATE22Disc",
                "GATE21RollNo": "GATE21RollNo",
                "GATE21Rank": "GATE21Rank",
                "GATE21Score": "GATE21Score",
                "GATE21Disc": "GATE21Disc",
                "GATE20RollNo": "GATE20RollNo",
                "GATE20Rank": "GATE20Rank",
                "GATE20Score": "GATE20Score",
                "GATE20Disc": "GATE20Disc",
                "MaxGATEScore out of 3 yrs": "MaxGATEScore_3yrs",
                "HSSC(board)": "HSSC_board",
                "HSSC(date)": "HSSC_date",
                "HSSC(per)": "HSSC_per",
                "SSC(board)": "SSC_board",
                "SSC(date)": "SSC_date",
                "SSC(per)": "SSC_per",
                "Degree(Qualification)": "Degree_Qualification",
                "Degree(PassingDate)": "Degree_PassingDate",
                "Degree(Branch)": "Degree_Branch",
                "Degree(OtherBranch)": "Degree_OtherBranch",
                "Degree(Institute Name)": "Degree_Institute",
                "Degree(CGPA-7thSem)": "Degree_CGPA_7th",
                "Degree(CGPA-8thSem)": "Degree_CGPA_8th",
                "Degree(Per-7thSem)": "Degree_Per_7th",
                "Degree(Per-8thSem)": "Degree_Per_8th",
                "unnamed": "ExtraColumn",
                "GATE Roll num": "GATE_Roll_num"
            }

            # Only rename columns that exist
            df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns}, inplace=True)

        except Exception as e:
            self.finished.emit(f"Failed to read Excel: {e}")
            return

        self.progress.emit("Inserting into database...")
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()

            columns = ", ".join([f'"{c}"' for c in df.columns])
            placeholders = ", ".join("?" * len(df.columns))

            for _, row in df.iterrows():
                cursor.execute(f'INSERT OR IGNORE INTO candidates ({columns}) VALUES ({placeholders})', tuple(row))

            conn.commit()
            conn.close()
            self.finished.emit("Excel data inserted successfully!")
        except Exception as e:
            self.finished.emit(f"Database error: {e}")
