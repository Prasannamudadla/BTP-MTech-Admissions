# import sqlite3
# import pandas as pd
# from PySide6.QtWidgets import QMessageBox

# DB_NAME = "mtech_offers.db"
# def _read_maybe_df(obj):
#     """Helper: if obj is a DataFrame return it, otherwise treat as filepath and pd.read_excel/csv."""
#     if isinstance(obj, pd.DataFrame):
#         return obj.copy()
#     if obj is None:
#         return pd.DataFrame()
#     # try excel first then csv
#     try:
#         return pd.read_excel(obj)
#     except Exception:
#         return pd.read_csv(obj)
# def run_round_1():
#     """Perform Round 1 seat allocation with proper PWD priority and category handling."""
#     conn = sqlite3.connect(DB_NAME)
#     cursor = conn.cursor()

#     try:
#         # Fetch candidates sorted by GATE score
#         # SELEECT App_no, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
#         cursor.execute("""
#             SELECT App_no, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
#             FROM candidates
#             WHERE MaxGATEScore_3yrs IS NOT NULL
#             ORDER BY MaxGATEScore_3yrs DESC
#         """)
#         candidates = cursor.fetchall()
#         print(f"Total candidates fetched: {len(candidates)}")

#         # Fetch seat matrix
#         cursor.execute("SELECT category, set_seats, seats_allocated FROM seat_matrix")
#         seat_matrix = {
#             cat.strip(): {"total": total or 0, "allocated": allocated or 0}
#             for cat, total, allocated in cursor.fetchall()
#         }
#         print("Seat Matrix Loaded:", seat_matrix)

#         # Extract common PWD quota (do NOT treat it as extra seat)
#         common_pwd_quota = seat_matrix.get("COMMON_PWD", {"total": 0})["total"]

#         # Create offers table if not exists
#         cursor.execute("""
#             CREATE TABLE IF NOT EXISTS offers (
#                 round_no INTEGER,
#                 App_no TEXT,
#                 Full_Name TEXT,
#                 category TEXT,
#                 MaxGATEScore_3yrs REAL,
#                 offer_status TEXT,
#                 PRIMARY KEY (round_no, App_no)
#             )
#         """)

#         offers_made = []
#         allocated_appnos = set()

#         # Split PWD and non-PWD candidates
#         pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() == "Yes"]
#         non_pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() != "Yes"]

#         # 🔹 Step 1: Allocate COMMON_PWD first (if available)
#         if common_pwd_quota > 0 and pwd_candidates:
#             top_pwd = pwd_candidates[0]  # Highest-ranked PWD candidate
#             app_no, name, base_cat, ews, gender, pwd, score = top_pwd

#             base_cat = base_cat.strip() if base_cat else "GEN"
#             gender = gender.strip().capitalize() if gender else "Male"
#             ews = ews.strip().capitalize() if ews else "No"

#             seat_key_parts = []
#             if ews == "Yes":
#                 seat_key_parts.append("EWS")
#             else:
#                 seat_key_parts.append(base_cat)

#             # If female, try Female seat first
#             possible_keys = []
#             if gender == "Female":
#                 possible_keys = [
#                     f"{seat_key_parts[0]}_Female",
#                     f"{seat_key_parts[0]}_FandM"
#                 ]
#             else:
#                 possible_keys = [f"{seat_key_parts[0]}_FandM"]

#             allocated = False
#             for key in possible_keys:
#                 if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
#                     seat_matrix[key]["allocated"] += 1
#                     offers_made.append((1, app_no, name, key, score, "Offered (Common PWD)"))
#                     allocated_appnos.add(app_no)
#                     allocated = True
#                     common_pwd_quota -= 1
#                     print(f"Allocating COMMON_PWD candidate {name} to {key}")
#                     break

#             if not allocated:
#                 print(f"COMMON_PWD candidate {name} could not be allocated (no seat available)")

#         # 🔹 Step 2: Allocate remaining PWD candidates (category-specific)
#         for app_no, name, base_cat, ews, gender, pwd, score in pwd_candidates:
#             if app_no in allocated_appnos:
#                 continue  # Already got COMMON_PWD seat

#             base_cat = base_cat.strip() if base_cat else "GEN"
#             gender = gender.strip().capitalize() if gender else "Male"
#             ews = ews.strip().capitalize() if ews else "No"

#             seat_key_parts = []
#             if ews == "Yes":
#                 seat_key_parts.append("EWS")
#             else:
#                 seat_key_parts.append(base_cat)

#             # Female can go to Female_PWD or FandM_PWD
#             possible_keys = []
#             if gender == "Female":
#                 possible_keys = [
#                     f"{seat_key_parts[0]}_Female_PWD",
#                     f"{seat_key_parts[0]}_FandM_PWD"
#                 ]
#             else:
#                 possible_keys = [f"{seat_key_parts[0]}_FandM_PWD"]

#             allocated = False
#             for key in possible_keys:
#                 if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
#                     seat_matrix[key]["allocated"] += 1
#                     offers_made.append((1, app_no, name, key, score, "Offered (PWD)"))
#                     allocated_appnos.add(app_no)
#                     allocated = True
#                     print(f"Allocating {name} (PWD) to {key}")
#                     break

#             if not allocated:
#                 print(f"No seat available for {name} (PWD)")

#         # 🔹 Step 3: Allocate Non-PWD candidates
#         for app_no, name, base_cat, ews, gender, pwd, score in non_pwd_candidates:
#             base_cat = base_cat.strip() if base_cat else "GEN"
#             gender = gender.strip().capitalize() if gender else "Male"
#             ews = ews.strip().capitalize() if ews else "No"

#             seat_key_parts = []
#             if ews == "Yes":
#                 seat_key_parts.append("EWS")
#             else:
#                 seat_key_parts.append(base_cat)

#             # Females: try Female → FandM; Males: FandM only
#             possible_keys = []
#             if gender == "Female":
#                 possible_keys = [
#                     f"{seat_key_parts[0]}_Female",
#                     f"{seat_key_parts[0]}_FandM"
#                 ]
#             else:
#                 possible_keys = [f"{seat_key_parts[0]}_FandM"]

#             allocated = False
#             for key in possible_keys:
#                 if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
#                     seat_matrix[key]["allocated"] += 1
#                     offers_made.append((1, app_no, name, key, score, "Offered"))
#                     allocated = True
#                     print(f"Allocating {name} to {key}")
#                     break

#             if not allocated:
#                 print(f"No seat available for {name}")

#         # 🔹 Step 4: Save results
#         cursor.executemany("""
#             INSERT OR IGNORE INTO offers (round_no, App_no, Full_Name, category, MaxGATEScore_3yrs, offer_status)
#             VALUES (?, ?, ?, ?, ?, ?)
#         """, offers_made)

#         # Update seat_matrix allocations in DB (except COMMON_PWD — it’s not a real seat)
#         for cat, data in seat_matrix.items():
#             if cat != "COMMON_PWD":
#                 cursor.execute(
#                     "UPDATE seat_matrix SET seats_allocated = ? WHERE category = ?",
#                     (data["allocated"], cat)
#                 )

#         conn.commit()
#         print("Seat matrix updated in DB")
#         QMessageBox.information(None, "Success", f"Round 1 allocation complete!\nTotal offers: {len(offers_made)}")

#     except Exception as e:
#         QMessageBox.critical(None, "Error", f"Error during round 1 allocation:\n{e}")
#     finally:
#         conn.close()

# # def download_offers(round_no=1):
# #     """Export offers for a given round to Excel."""
# #     conn = sqlite3.connect(DB_NAME)
# #     df = pd.read_sql_query(f"SELECT * FROM offers WHERE round_no = {round_no}", conn)
# #     conn.close()

# #     if df.empty:
# #         QMessageBox.warning(None, "No Offers", f"No offers found for Round {round_no}")
# #         return

# #     filename = f"Round{round_no}_Offers.xlsx"
# #     df.to_excel(filename, index=False)
# #     QMessageBox.information(None, "Download Complete", f"Offers saved as {filename}")
# def download_offers(round_no=1):
#     """Export offers for a given round to Excel with two sheets."""
#     conn = sqlite3.connect(DB_NAME)

#     # Fetch basic offers (Sheet 1)
#     df_offers = pd.read_sql_query(f"SELECT * FROM offers WHERE round_no = {round_no}", conn)

#     if df_offers.empty:
#         QMessageBox.warning(None, "No Offers", f"No offers found for Round {round_no}")
#         conn.close()
#         return

#     # Fetch detailed candidate info for offered candidates (Sheet 2)
#     app_nos = tuple(df_offers['App_no'].tolist())  # Tuple for SQL IN
#     if len(app_nos) == 1:
#         # SQL IN requires special handling for single element
#         app_nos = f"('{app_nos[0]}')"

#     query = f"""
#         SELECT o.round_no, c.*
#         FROM offers o
#         JOIN candidates c ON o.App_no = c.App_no
#         WHERE o.round_no = {round_no}
#         ORDER BY o.MaxGATEScore_3yrs DESC
#     """
#     df_detailed = pd.read_sql_query(query, conn)

#     conn.close()

#     # Save to Excel with multiple sheets
#     filename = f"Round{round_no}_Offers.xlsx"
#     with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
#         df_offers.to_excel(writer, sheet_name='Offers_Summary', index=False)
#         df_detailed.to_excel(writer, sheet_name='Offers_Detailed', index=False)

#     QMessageBox.information(None, "Download Complete", f"Offers saved as {filename}")

import sqlite3
import pandas as pd
from PySide6.QtWidgets import QMessageBox

DB_NAME = "mtech_offers.db"

def _read_maybe_df(obj):
    """Helper: if obj is a DataFrame return it, otherwise treat as filepath and pd.read_excel/csv."""
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    if obj is None:
        return pd.DataFrame()
    try:
        return pd.read_excel(obj)
    except Exception:
        return pd.read_csv(obj)

def run_round_1():
    """Perform Round 1 seat allocation with proper PWD priority and category handling."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    try:
        # Fetch candidates sorted by GATE score
        cursor.execute("""
            SELECT COAP, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
            FROM candidates
            WHERE MaxGATEScore_3yrs IS NOT NULL
            ORDER BY MaxGATEScore_3yrs DESC
        """)
        candidates = cursor.fetchall()
        print(f"Total candidates fetched: {len(candidates)}")

        # Fetch seat matrix
        cursor.execute("SELECT category, set_seats, seats_allocated FROM seat_matrix")
        seat_matrix = {
            cat.strip(): {"total": total or 0, "allocated": allocated or 0}
            for cat, total, allocated in cursor.fetchall()
        }
        print("Seat Matrix Loaded:", seat_matrix)

        # Extract common PWD quota (do NOT treat it as extra seat)
        common_pwd_quota = seat_matrix.get("COMMON_PWD", {"total": 0})["total"]

        # Create offers table if not exists
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS offers (
                round_no INTEGER,
                COAP TEXT,
                Full_Name TEXT,
                category TEXT,
                MaxGATEScore_3yrs REAL,
                offer_status TEXT,
                PRIMARY KEY (round_no, COAP)
            )
        """)

        offers_made = []
        allocated_coaps = set()

        # Split PWD and non-PWD candidates
        pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() == "Yes"]
        non_pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() != "Yes"]

        # 🔹 Step 1: Allocate COMMON_PWD first
        if common_pwd_quota > 0 and pwd_candidates:
            top_pwd = pwd_candidates[0]  # Highest-ranked PWD candidate
            coap, name, base_cat, ews, gender, pwd, score = top_pwd

            base_cat = base_cat.strip() if base_cat else "GEN"
            gender = gender.strip().capitalize() if gender else "Male"
            ews = ews.strip().capitalize() if ews else "No"

            seat_key_parts = ["EWS" if ews == "Yes" else base_cat]

            # Female preference
            possible_keys = [f"{seat_key_parts[0]}_Female", f"{seat_key_parts[0]}_FandM"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM"]

            allocated = False
            for key in possible_keys:
                if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
                    seat_matrix[key]["allocated"] += 1
                    offers_made.append((1, coap, name, key, score, "Offered (Common PWD)"))
                    allocated_coaps.add(coap)
                    allocated = True
                    common_pwd_quota -= 1
                    print(f"Allocating COMMON_PWD candidate {name} to {key}")
                    break
            if not allocated:
                print(f"COMMON_PWD candidate {name} could not be allocated (no seat available)")

        # 🔹 Step 2: Allocate remaining PWD candidates
        for coap, name, base_cat, ews, gender, pwd, score in pwd_candidates:
            if coap in allocated_coaps:
                continue

            base_cat = base_cat.strip() if base_cat else "GEN"
            gender = gender.strip().capitalize() if gender else "Male"
            ews = ews.strip().capitalize() if ews else "No"
            seat_key_parts = ["EWS" if ews=="Yes" else base_cat]

            possible_keys = [f"{seat_key_parts[0]}_Female_PWD", f"{seat_key_parts[0]}_FandM_PWD"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM_PWD"]

            allocated = False
            for key in possible_keys:
                if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
                    seat_matrix[key]["allocated"] += 1
                    offers_made.append((1, coap, name, key, score, "Offered (PWD)"))
                    allocated_coaps.add(coap)
                    allocated = True
                    print(f"Allocating {name} (PWD) to {key}")
                    break
            if not allocated:
                print(f"No seat available for {name} (PWD)")

        # 🔹 Step 3: Allocate Non-PWD candidates
        for coap, name, base_cat, ews, gender, pwd, score in non_pwd_candidates:
            base_cat = base_cat.strip() if base_cat else "GEN"
            gender = gender.strip().capitalize() if gender else "Male"
            ews = ews.strip().capitalize() if ews else "No"
            seat_key_parts = ["EWS" if ews=="Yes" else base_cat]

            possible_keys = [f"{seat_key_parts[0]}_Female", f"{seat_key_parts[0]}_FandM"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM"]

            allocated = False
            for key in possible_keys:
                if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
                    seat_matrix[key]["allocated"] += 1
                    offers_made.append((1, coap, name, key, score, "Offered"))
                    allocated = True
                    print(f"Allocating {name} to {key}")
                    break
            if not allocated:
                print(f"No seat available for {name}")

        # 🔹 Step 4: Save results
        cursor.executemany("""
            INSERT OR IGNORE INTO offers (round_no, COAP, Full_Name, category, MaxGATEScore_3yrs, offer_status)
            VALUES (?, ?, ?, ?, ?, ?)
        """, offers_made)

        # Update seat_matrix allocations in DB
        for cat, data in seat_matrix.items():
            if cat != "COMMON_PWD":
                cursor.execute("UPDATE seat_matrix SET seats_allocated = ? WHERE category = ?", (data["allocated"], cat))

        conn.commit()
        print("Seat matrix updated in DB")
        QMessageBox.information(None, "Success", f"Round 1 allocation complete!\nTotal offers: {len(offers_made)}")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during round 1 allocation:\n{e}")
    finally:
        conn.close()


def download_offers(round_no=1):
    """Export offers for a given round to Excel with two sheets using COAP numbers."""
    conn = sqlite3.connect(DB_NAME)

    # Sheet 1: Basic offers
    df_offers = pd.read_sql_query(f"""
        SELECT o.round_no, c.COAP, o.Full_Name, o.category, o.MaxGATEScore_3yrs, o.offer_status
        FROM offers o
        JOIN candidates c ON o.COAP = c.COAP
        WHERE o.round_no = {round_no}
    """, conn)

    if df_offers.empty:
        QMessageBox.warning(None, "No Offers", f"No offers found for Round {round_no}")
        conn.close()
        return

    # Sheet 2: Detailed offers
    query = f"""
        SELECT o.round_no, c.*
        FROM offers o
        JOIN candidates c ON o.COAP = c.COAP
        WHERE o.round_no = {round_no}
        ORDER BY o.MaxGATEScore_3yrs DESC
    """
    df_detailed = pd.read_sql_query(query, conn)
    conn.close()

    # Save to Excel with multiple sheets
    filename = f"Round{round_no}_Offers.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:  # openpyxl avoids extra dependency
        df_offers.to_excel(writer, sheet_name='Offers_Summary', index=False)
        df_detailed.to_excel(writer, sheet_name='Offers_Detailed', index=False)

    QMessageBox.information(None, "Download Complete", f"Offers saved as {filename}")
