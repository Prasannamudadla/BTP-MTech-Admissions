# import os
# import sqlite3
# import pandas as pd
# import sys
# DB_NAME = "mtech_offers.db"
# # EXCEL_FILE = "ApplicantData_withCOAPcorr_maxGateRoll.xlsx"  # Update path if needed

# # --------------------------
# # Step 1: Delete old DB
# # --------------------------
# if os.path.exists(DB_NAME):
#     os.remove(DB_NAME)
#     print("Old database deleted.")

# # --------------------------
# # Step 2: Create candidates table
# # --------------------------
# def create_candidates_table():
#     conn = sqlite3.connect(DB_NAME)
#     cursor = conn.cursor()

#     cursor.execute("""
#     CREATE TABLE IF NOT EXISTS candidates (
#         Si_NO INTEGER,
#         App_no TEXT PRIMARY KEY,
#         Email TEXT,
#         Full_Name TEXT,
#         Adm_cat TEXT,
#         Pwd TEXT,
#         Ews TEXT,
#         Gender TEXT,
#         Category TEXT,
#         COAP TEXT,
#         GATE22RollNo TEXT,
#         GATE22Rank INTEGER,
#         GATE22Score REAL,
#         GATE22Disc TEXT,
#         GATE21RollNo TEXT,
#         GATE21Rank INTEGER,
#         GATE21Score REAL,
#         GATE21Disc TEXT,
#         GATE20RollNo TEXT,
#         GATE20Rank INTEGER,
#         GATE20Score REAL,
#         GATE20Disc TEXT,
#         MaxGATEScore_3yrs REAL,
#         HSSC_board TEXT,
#         HSSC_date TEXT,
#         HSSC_per REAL,
#         SSC_board TEXT,
#         SSC_date TEXT,
#         SSC_per REAL,
#         Degree_Qualification TEXT,
#         Degree_PassingDate TEXT,
#         Degree_Branch TEXT,
#         Degree_OtherBranch TEXT,
#         Degree_Institute TEXT,
#         Degree_CGPA_7th REAL,
#         Degree_CGPA_8th REAL,
#         Degree_Per_7th REAL,
#         Degree_Per_8th REAL,
#         ExtraColumn TEXT,
#         GATE_Roll_num TEXT
#     )
#     """)
#     conn.commit()
#     conn.close()
#     print("Candidates table created successfully.")

# create_candidates_table()

# # --------------------------
# # Step 2b: Create seat_matrix table
# # --------------------------
# def create_seat_matrix_table():
#     conn = sqlite3.connect(DB_NAME)
#     cursor = conn.cursor()

#     cursor.execute("""
#     CREATE TABLE IF NOT EXISTS seat_matrix (
#         category TEXT PRIMARY KEY,
#         set_seats INTEGER DEFAULT 0,
#         seats_allocated INTEGER DEFAULT 0, 
#         seats_booked INTEGER DEFAULT 0
#     )
#     """)
#     conn.commit()
#     conn.close()
#     print("Seat matrix table created successfully.")

# create_seat_matrix_table()


# --------------------------
# Step 3: Read and clean Excel
# --------------------------
# if len(sys.argv) < 2:
#     print("Error: Please provide the Excel file path as a command-line argument.")
#     print("Usage: python full_setup.py /path/to/your/file.xlsx")
#     sys.exit(1)

# EXCEL_FILE = sys.argv[1]

# try:
#     df = pd.read_excel(EXCEL_FILE)

#     # Remove completely empty columns
#     df = df.loc[:, df.columns.notnull()]

#     # Rename blank or unnamed columns to ExtraColumn
#     df.rename(columns=lambda x: 'ExtraColumn' if str(x).strip() == '' else x, inplace=True)

#     # Remove duplicate columns
#     df = df.loc[:, ~df.columns.duplicated()]

#     # Fill missing values with None
#     df = df.where(pd.notnull(df), None)

#     # Rename Excel columns to match SQLite table
#     column_mapping = {
#         "Si NO": "Si_NO",
#         "App no": "App_no",
#         "Full Name": "Full_Name",
#         "Adm cat": "Adm_cat",
#         "MaxGATEScore out of 3 yrs": "MaxGATEScore_3yrs",
#         "HSSC(date)": "HSSC_date",
#         "HSSC(board)": "HSSC_board",
#         "HSSC(per)": "HSSC_per",
#         "SSC(date)": "SSC_date",
#         "SSC(board)": "SSC_board",
#         "SSC(per)": "SSC_per",
#         "Degree(PassingDate)": "Degree_PassingDate",
#         "Degree(Qualification)": "Degree_Qualification",
#         "Degree(Branch)": "Degree_Branch",
#         "Degree(OtherBranch)": "Degree_OtherBranch",
#         "Degree(Institute Name)": "Degree_Institute",
#         "Degree(CGPA-7thSem)": "Degree_CGPA_7th",
#         "Degree(CGPA-8thSem)": "Degree_CGPA_8th",
#         "Degree(Per-7thSem)": "Degree_Per_7th",
#         "Degree(Per-8thSem)": "Degree_Per_8th",
#         "GATE Roll num": "GATE_Roll_num",
#         "unnamed": "ExtraColumn"
#     }


#     # Apply only existing columns
#     df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns}, inplace=True)

#     # Convert datetime columns to string for SQLite
#     datetime_cols = ['HSSC_date', 'SSC_date', 'Degree_PassingDate']
#     for col in datetime_cols:
#         if col in df.columns:
#             df[col] = df[col].apply(
#                 lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) and not isinstance(x, str) else x
#             )

#     print("Excel read successfully with columns:", df.columns.tolist())

# except Exception as e:
#     print(f"Failed to read Excel: {e}")
#     exit()

# # --------------------------
# # Step 4: Insert into SQLite
# # --------------------------
# try:
#     conn = sqlite3.connect(DB_NAME)
#     cursor = conn.cursor()

#     columns = ", ".join([f'"{c}"' for c in df.columns])
#     placeholders = ", ".join("?" * len(df.columns))

#     for _, row in df.iterrows():
#         cursor.execute(f'INSERT OR IGNORE INTO candidates ({columns}) VALUES ({placeholders})', tuple(row))

#     conn.commit()
#     conn.close()
#     print("Excel data inserted into database successfully!")

# except Exception as e:
#     print(f"Database error: {e}")
