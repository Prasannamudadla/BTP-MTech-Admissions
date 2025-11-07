# code with app no as main
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
# code with coap as main
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
#         cursor.execute("""
#             SELECT COAP, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
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
#                 COAP TEXT,
#                 Full_Name TEXT,
#                 category TEXT,
#                 MaxGATEScore_3yrs REAL,
#                 offer_status TEXT,
#                 PRIMARY KEY (round_no, COAP)
#             )
#         """)

#         offers_made = []
#         allocated_coaps = set()

#         # Split PWD and non-PWD candidates
#         pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() == "Yes"]
#         non_pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() != "Yes"]

#         # 🔹 Step 1: Allocate COMMON_PWD first
#         if common_pwd_quota > 0 and pwd_candidates:
#             top_pwd = pwd_candidates[0]  # Highest-ranked PWD candidate
#             coap, name, base_cat, ews, gender, pwd, score = top_pwd

#             base_cat = base_cat.strip() if base_cat else "GEN"
#             gender = gender.strip().capitalize() if gender else "Male"
#             ews = ews.strip().capitalize() if ews else "No"

#             seat_key_parts = ["EWS" if ews == "Yes" else base_cat]

#             # Female preference
#             possible_keys = [f"{seat_key_parts[0]}_Female", f"{seat_key_parts[0]}_FandM"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM"]

#             allocated = False
#             for key in possible_keys:
#                 if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
#                     seat_matrix[key]["allocated"] += 1
#                     offers_made.append((1, coap, name, key, score, "Offered (Common PWD)"))
#                     allocated_coaps.add(coap)
#                     allocated = True
#                     common_pwd_quota -= 1
#                     print(f"Allocating COMMON_PWD candidate {name} to {key}")
#                     break
#             if not allocated:
#                 print(f"COMMON_PWD candidate {name} could not be allocated (no seat available)")

#         # 🔹 Step 2: Allocate remaining PWD candidates
#         for coap, name, base_cat, ews, gender, pwd, score in pwd_candidates:
#             if coap in allocated_coaps:
#                 continue

#             base_cat = base_cat.strip() if base_cat else "GEN"
#             gender = gender.strip().capitalize() if gender else "Male"
#             ews = ews.strip().capitalize() if ews else "No"
#             seat_key_parts = ["EWS" if ews=="Yes" else base_cat]

#             possible_keys = [f"{seat_key_parts[0]}_Female_PWD", f"{seat_key_parts[0]}_FandM_PWD"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM_PWD"]

#             allocated = False
#             for key in possible_keys:
#                 if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
#                     seat_matrix[key]["allocated"] += 1
#                     offers_made.append((1, coap, name, key, score, "Offered (PWD)"))
#                     allocated_coaps.add(coap)
#                     allocated = True
#                     print(f"Allocating {name} (PWD) to {key}")
#                     break
#             if not allocated:
#                 print(f"No seat available for {name} (PWD)")

#         # 🔹 Step 3: Allocate Non-PWD candidates
#         for coap, name, base_cat, ews, gender, pwd, score in non_pwd_candidates:
#             base_cat = base_cat.strip() if base_cat else "GEN"
#             gender = gender.strip().capitalize() if gender else "Male"
#             ews = ews.strip().capitalize() if ews else "No"
#             seat_key_parts = ["EWS" if ews=="Yes" else base_cat]

#             possible_keys = [f"{seat_key_parts[0]}_Female", f"{seat_key_parts[0]}_FandM"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM"]

#             allocated = False
#             for key in possible_keys:
#                 if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
#                     seat_matrix[key]["allocated"] += 1
#                     offers_made.append((1, coap, name, key, score, "Offered"))
#                     allocated = True
#                     print(f"Allocating {name} to {key}")
#                     break
#             if not allocated:
#                 print(f"No seat available for {name}")

#         # 🔹 Step 4: Save results
#         cursor.executemany("""
#             INSERT OR IGNORE INTO offers (round_no, COAP, Full_Name, category, MaxGATEScore_3yrs, offer_status)
#             VALUES (?, ?, ?, ?, ?, ?)
#         """, offers_made)

#         # Update seat_matrix allocations in DB
#         for cat, data in seat_matrix.items():
#             if cat != "COMMON_PWD":
#                 cursor.execute("UPDATE seat_matrix SET seats_allocated = ? WHERE category = ?", (data["allocated"], cat))

#         conn.commit()
#         print("Seat matrix updated in DB")
#         QMessageBox.information(None, "Success", f"Round 1 allocation complete!\nTotal offers: {len(offers_made)}")

#     except Exception as e:
#         QMessageBox.critical(None, "Error", f"Error during round 1 allocation:\n{e}")
#     finally:
#         conn.close()


# def download_offers(round_no=1):
#     """Export offers for a given round to Excel with two sheets using COAP numbers."""
#     conn = sqlite3.connect(DB_NAME)

#     # Sheet 1: Basic offers
#     df_offers = pd.read_sql_query(f"""
#         SELECT o.round_no, c.COAP, o.Full_Name, o.category, o.MaxGATEScore_3yrs, o.offer_status
#         FROM offers o
#         JOIN candidates c ON o.COAP = c.COAP
#         WHERE o.round_no = {round_no}
#     """, conn)

#     if df_offers.empty:
#         QMessageBox.warning(None, "No Offers", f"No offers found for Round {round_no}")
#         conn.close()
#         return

#     # Sheet 2: Detailed offers
#     query = f"""
#         SELECT o.round_no, c.*
#         FROM offers o
#         JOIN candidates c ON o.COAP = c.COAP
#         WHERE o.round_no = {round_no}
#         ORDER BY o.MaxGATEScore_3yrs DESC
#     """
#     df_detailed = pd.read_sql_query(query, conn)
#     conn.close()

#     # Save to Excel with multiple sheets
#     filename = f"Round{round_no}_Offers.xlsx"
#     with pd.ExcelWriter(filename, engine='openpyxl') as writer:  # openpyxl avoids extra dependency
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
        # Tries to read the first sheet of an excel file by default
        return pd.read_excel(obj)
    except Exception:
        # Fallback to CSV
        return pd.read_csv(obj)

def _create_decision_tables(cursor, round_no):
    """Creates the necessary decision tables for a given round if they don't exist."""
    # Table 1: IIT Goa Candidate Decision Report (Mtech App No, Applicant Decision)
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS iit_goa_offers_round{round_no} (
            mtech_app_no TEXT,
            applicant_decision TEXT,
            PRIMARY KEY (mtech_app_no)
        )
    """)
    # Table 2: IIT Goa Offered But Accept and Freeze at Other Institutes (Mtech App No, Other Institution Decision)
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS accepted_other_institute_round{round_no} (
            mtech_app_no TEXT,
            other_institute_decision TEXT,
            PRIMARY KEY (mtech_app_no)
        )
    """)
    # Table 3: Consolidated Accept and Freeze Candidates Across All Institutes (COAP Reg Id, Applicant Decision)
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS consolidated_decisions_round{round_no} (
            coap_reg_id TEXT,
            applicant_decision TEXT,
            PRIMARY KEY (coap_reg_id)
        )
    """)

def upload_round_decisions(round_no, iit_goa_report, other_iit_report, consolidated_report):
    """
    Reads the three decision reports for a given round and saves them to the database.

    :param round_no: The round number (e.g., 1).
    :param iit_goa_report: Filepath or DataFrame for IIT Goa Candidate Decision Report.
    :param other_iit_report: Filepath or DataFrame for IIT Goa Offered But Accept and Freeze at Other Institutes.
    :param consolidated_report: Filepath or DataFrame for Consolidated Accept and Freeze Candidates Across All Institutes.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    try:
        _create_decision_tables(cursor, round_no)

        # 1. IIT Goa Candidate Decision Report
        df_goa = _read_maybe_df(iit_goa_report)
        df_goa = df_goa[["MTech Application No", "Applicant Decision"]].rename(
            columns={"MTech Application No": "mtech_app_no", "Applicant Decision": "applicant_decision"}
        )
        df_goa.to_sql(f'iit_goa_offers_round{round_no}', conn, if_exists='replace', index=False)
        
        # 2. IIT Goa Offered But Accept and Freeze at Other Institutes
        df_other = _read_maybe_df(other_iit_report)
        df_other = df_other[["MTech Application No", "Other Institution Decision"]].rename(
            columns={"MTech Application No": "mtech_app_no", "Other Institution Decision": "other_institute_decision"}
        )
        df_other.to_sql(f'accepted_other_institute_round{round_no}', conn, if_exists='replace', index=False)
        
        # 3. Consolidated Accept and Freeze Candidates Across All Institutes
        df_consolidated = _read_maybe_df(consolidated_report)
        df_consolidated = df_consolidated[["COAP Reg Id", "Applicant Decision"]].rename(
            columns={"COAP Reg Id": "coap_reg_id", "Applicant Decision": "applicant_decision"}
        )
        # Assuming COAP Reg Id is the COAP number in the candidates table for linking
        df_consolidated.to_sql(f'consolidated_decisions_round{round_no}', conn, if_exists='replace', index=False)

        conn.commit()
        QMessageBox.information(None, "Success", f"Decisions for Round {round_no} uploaded and saved successfully!")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during decision upload for Round {round_no}:\n{e}")
    finally:
        conn.close()

def _get_eligible_candidates_for_next_round(current_round):
    """
    Determines the list of COAP IDs for candidates eligible for the next round (round_no + 1).

    Logic:
    - Exclude candidates who 'Accept and Freeze' or 'Reject and Wait' an IIT Goa offer in the current round.
    - Exclude candidates who 'Accept and Freeze' an offer at any other institute (Consolidated report).
    - Candidates who 'Retain and Wait' or 'Reject and Continue' (implicitly) are eligible.
    - Candidates who didn't get an offer are eligible.
    - For Round N, decisions from Round 1 to Round N are considered.
    """
    conn = sqlite3.connect(DB_NAME)
    
    # 1. Identify COAP IDs who have Accepted and Frozen an offer at ANY institute (IIT Goa or Others)
    # These candidates are out of all subsequent rounds.
    frozen_coaps = set()
    
    for r in range(1, current_round + 1):
        # Decisions on IIT Goa Offer (Accept and Freeze)
        df_goa_offers = pd.read_sql_query(f"""
            SELECT o.COAP
            FROM offers o
            JOIN iit_goa_offers_round{r} d ON d.mtech_app_no = c.App_no
            JOIN candidates c ON o.COAP = c.COAP
            WHERE o.round_no = {r} AND d.applicant_decision = 'Accept and Freeze'
        """, conn)
        frozen_coaps.update(df_goa_offers['COAP'].tolist())

        # Consolidated Decisions (Accept and Freeze at any Institute)
        # Note: Consolidated report uses COAP Reg Id which is assumed to be the 'COAP' column in the 'candidates' table.
        df_consolidated = pd.read_sql_query(f"""
            SELECT coap_reg_id 
            FROM consolidated_decisions_round{r}
            WHERE applicant_decision = 'Accept and Freeze'
        """, conn)
        frozen_coaps.update(df_consolidated['coap_reg_id'].tolist())

    # 2. Identify COAP IDs who 'Reject and Wait' an IIT Goa offer in the *latest* round (and are thus out).
    # This is slightly more complex, as 'Reject and Wait' means they're out, but 'Retain and Wait' means they're in.
    # The simplest is to find all candidates who are NOT 'Retain and Wait' and are not 'Accept and Freeze'.
    
    # Let's consider all candidates who received an offer in the *latest* round (current_round) at IIT Goa.
    latest_goa_decisions_query = f"""
        SELECT o.COAP, d.applicant_decision
        FROM offers o
        JOIN candidates c ON o.COAP = c.COAP
        JOIN iit_goa_offers_round{current_round} d ON d.mtech_app_no = c.App_no
        WHERE o.round_no = {current_round}
    """
    df_latest_decisions = pd.read_sql_query(latest_goa_decisions_query, conn)
    
    # Candidates who explicitly 'Reject and Wait' in the latest round are also excluded
    rejected_and_out_coaps = df_latest_decisions[
        df_latest_decisions['applicant_decision'] == 'Reject and Wait'
    ]['COAP'].tolist()

    # 3. Identify ALL candidates who are out
    coaps_out = frozen_coaps.union(set(rejected_and_out_coaps))
    
    # 4. Fetch all candidates' COAP IDs from the candidates table who have a GATE score
    all_eligible_coaps_query = """
        SELECT COAP 
        FROM candidates 
        WHERE MaxGATEScore_3yrs IS NOT NULL
    """
    df_all_candidates = pd.read_sql_query(all_eligible_coaps_query, conn)
    all_coaps = set(df_all_candidates['COAP'].tolist())

    conn.close()

    # 5. Filter: Eligible for next round = All candidates - Candidates who are out
    eligible_coaps = list(all_coaps - coaps_out)
    
    print(f"Total candidates: {len(all_coaps)}, Excluded: {len(coaps_out)}, Eligible for Round {current_round + 1}: {len(eligible_coaps)}")

    return eligible_coaps

def run_round(round_no):
    """
    Perform seat allocation for a given round, respecting prior round decisions and exclusions.
    This function replaces run_round_1 and is generic for all rounds >= 1.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    try:
        eligible_coaps = None
        if round_no == 1:
            # For Round 1, all candidates with a GATE score are eligible
            eligible_coaps = [] # Will be populated by the main query below
        else:
            # For subsequent rounds, filter based on previous round decisions
            eligible_coaps = _get_eligible_candidates_for_next_round(round_no - 1)
            if not eligible_coaps:
                QMessageBox.warning(None, "Round Complete", f"No eligible candidates remain for Round {round_no}.")
                conn.close()
                return

        # Fetch candidates sorted by GATE score, applying eligibility filter
        coap_filter = ""
        if round_no > 1:
            # Create a string of COAP IDs for the SQL IN clause
            coap_list_str = ', '.join([f"'{c}'" for c in eligible_coaps])
            coap_filter = f"AND COAP IN ({coap_list_str})"
            
        # Select candidates who are *not* currently offered a seat in a previous round (Accept and Freeze candidates are already filtered out by _get_eligible_candidates_for_next_round)
        # We only consider those who are *not* yet allocated a confirmed (Accept and Freeze) seat.
        
        # For simplicity in this logic structure, we rely solely on the eligibility filter from the previous step
        cursor.execute(f"""
            SELECT COAP, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
            FROM candidates
            WHERE MaxGATEScore_3yrs IS NOT NULL
            {coap_filter}
            ORDER BY MaxGATEScore_3yrs DESC
        """)
        candidates = cursor.fetchall()
        print(f"Total candidates eligible for Round {round_no}: {len(candidates)}")

        # Fetch seat matrix (allocated seats should be the current confirmed seats)
        # Note: We assume seats_allocated in DB is updated only with 'Accept and Freeze' candidates' seats.
        # Since the provided logic updates seats_allocated immediately after offers, we'll *reset*
        # the allocation counts and recalculate confirmed seats for a more robust multi-round approach.
        
        # Recalculate confirmed seats before running the round
        confirmed_seats = _recalculate_confirmed_seats(round_no - 1, conn)
        seat_matrix = _get_seat_matrix_with_confirmed(conn, confirmed_seats)
        
        print(f"Seat Matrix Loaded (Confirmed Seats from R1 to R{round_no-1}):", {k: v['allocated'] for k, v in seat_matrix.items()})

        common_pwd_quota = seat_matrix.get("COMMON_PWD", {"total": 0})["total"]

        # The rest of the allocation logic remains the same (Steps 1, 2, 3) but uses the filtered candidates
        # and the updated seat_matrix with confirmed allocations.
        
        cursor.execute(f"""
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
        
        # Split PWD and non-PWD candidates from the *eligible* list
        pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() == "Yes"]
        non_pwd_candidates = [c for c in candidates if (c[5] or "").strip().capitalize() != "Yes"]

        # --- Allocation Steps (Same logic as run_round_1, but using the updated seat_matrix) ---

        # 🔹 Step 1: Allocate COMMON_PWD first
        temp_common_pwd_quota = common_pwd_quota # Use a temporary variable for allocation tracking
        if temp_common_pwd_quota > 0 and pwd_candidates:
            top_pwd = pwd_candidates[0]
            coap, name, base_cat, ews, gender, pwd, score = top_pwd

            base_cat = base_cat.strip() if base_cat else "GEN"
            gender = gender.strip().capitalize() if gender else "Male"
            ews = ews.strip().capitalize() if ews else "No"

            seat_key_parts = ["EWS" if ews == "Yes" else base_cat]
            possible_keys = [f"{seat_key_parts[0]}_Female", f"{seat_key_parts[0]}_FandM"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM"]

            allocated = False
            for key in possible_keys:
                # Allocation check: Check against total seats *minus* already allocated seats
                if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
                    seat_matrix[key]["allocated"] += 1
                    offers_made.append((round_no, coap, name, key, score, "Offered (Common PWD)"))
                    allocated_coaps.add(coap)
                    allocated = True
                    temp_common_pwd_quota -= 1
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
                    offers_made.append((round_no, coap, name, key, score, "Offered (PWD)"))
                    allocated_coaps.add(coap)
                    allocated = True
                    print(f"Allocating {name} (PWD) to {key}")
                    break
            if not allocated:
                print(f"No seat available for {name} (PWD)")

        # 🔹 Step 3: Allocate Non-PWD candidates
        for coap, name, base_cat, ews, gender, pwd, score in non_pwd_candidates:
            if coap in allocated_coaps:
                continue
                
            base_cat = base_cat.strip() if base_cat else "GEN"
            gender = gender.strip().capitalize() if gender else "Male"
            ews = ews.strip().capitalize() if ews else "No"
            seat_key_parts = ["EWS" if ews=="Yes" else base_cat]

            possible_keys = [f"{seat_key_parts[0]}_Female", f"{seat_key_parts[0]}_FandM"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM"]

            allocated = False
            for key in possible_keys:
                if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
                    seat_matrix[key]["allocated"] += 1
                    offers_made.append((round_no, coap, name, key, score, "Offered"))
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
        
        conn.commit()
        QMessageBox.information(None, "Success", f"Round {round_no} allocation complete!\nTotal offers: {len(offers_made)}")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during round {round_no} allocation:\n{e}")
    finally:
        conn.close()

def _recalculate_confirmed_seats(last_round, conn):
    """
    Recalculates the total confirmed seats (Accept and Freeze) up to the specified last_round.
    
    :param last_round: The last round to check (e.g., 1 for preparing for Round 2).
    :param conn: The active SQLite connection.
    :return: A dictionary of confirmed seat counts by category.
    """
    if last_round < 1:
        return {}
        
    confirmed_seats = {}
    
    for r in range(1, last_round + 1):
        # Find candidates who 'Accept and Freeze' their IIT Goa offer in this round
        query = f"""
            SELECT o.category, COUNT(o.COAP) as count
            FROM offers o
            JOIN candidates c ON o.COAP = c.COAP
            JOIN iit_goa_offers_round{r} d ON d.mtech_app_no = c.App_no
            WHERE o.round_no = {r} AND d.applicant_decision = 'Accept and Freeze'
            GROUP BY o.category
        """
        df_confirmed = pd.read_sql_query(query, conn)
        
        for index, row in df_confirmed.iterrows():
            cat = row['category']
            count = row['count']
            confirmed_seats[cat] = confirmed_seats.get(cat, 0) + count

    return confirmed_seats

def _get_seat_matrix_with_confirmed(conn, confirmed_seats):
    """Fetches the base seat matrix and updates the allocated count with confirmed seats."""
    cursor = conn.cursor()
    cursor.execute("SELECT category, set_seats, seats_allocated FROM seat_matrix")
    seat_matrix = {
        cat.strip(): {"total": total or 0, "allocated": confirmed_seats.get(cat.strip(), 0)}
        for cat, total, allocated in cursor.fetchall()
    }
    return seat_matrix

# --- The original download_offers function remains the same ---

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
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_offers.to_excel(writer, sheet_name='Offers_Summary', index=False)
        df_detailed.to_excel(writer, sheet_name='Offers_Detailed', index=False)

    QMessageBox.information(None, "Download Complete", f"Offers saved as {filename}")

# Note: The original `run_round_1` is replaced by the generic `run_round(1)`
# to allow for a unified, multi-round process.