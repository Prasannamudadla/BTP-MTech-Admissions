import sqlite3
import pandas as pd
from PySide6.QtWidgets import QMessageBox
import difflib

DB_NAME = "mtech_offers.db"

# ------------------------------------------------------
# Helper: Read DataFrame or Excel/CSV
# ------------------------------------------------------
def _read_maybe_df(obj):
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    if obj is None:
        return pd.DataFrame()
    try:
        return pd.read_excel(obj)
    except Exception:
        return pd.read_csv(obj)


# ------------------------------------------------------
# Create tables for each COAP round
# ------------------------------------------------------
def _create_decision_tables(cursor, round_no):
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS iit_goa_offers_round{round_no} (
            mtech_app_no TEXT,
            applicant_decision TEXT,
            PRIMARY KEY (mtech_app_no)
        )
    """)
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS accepted_other_institute_round{round_no} (
            mtech_app_no TEXT,
            other_institute_decision TEXT,
            PRIMARY KEY (mtech_app_no)
        )
    """)
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS consolidated_decisions_round{round_no} (
            coap_reg_id TEXT,
            applicant_decision TEXT,
            PRIMARY KEY (coap_reg_id)
        )
    """)


# ------------------------------------------------------
# Upload round decisions
# ------------------------------------------------------
def upload_round_decisions(round_no, goa_widget, other_widget, cons_widget):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    GOA_REQUIRED  = ["mtech_app_no", "applicant_decision"]
    OTHER_REQUIRED = ["mtech_app_no", "other_institute_decision"]
    CONS_REQUIRED = ["coap_reg_id", "applicant_decision"]

    try:
        _create_decision_tables(cursor, round_no)

        # --- GOA FILE ---
        df_goa = goa_widget.get_mapped_dataframe()
        df_goa = auto_match_columns(df_goa, GOA_REQUIRED)
        df_goa.to_sql(f'iit_goa_offers_round{round_no}', conn, 
                      if_exists='replace', index=False)

        # --- OTHER IIT FILE ---
        df_other = other_widget.get_mapped_dataframe()
        df_other = auto_match_columns(df_other, OTHER_REQUIRED)
        df_other.to_sql(f'accepted_other_institute_round{round_no}', conn, 
                        if_exists='replace', index=False)

        # --- CONSOLIDATED FILE ---
        df_cons = cons_widget.get_mapped_dataframe()
        df_cons = auto_match_columns(df_cons, CONS_REQUIRED)
        df_cons.to_sql(f'consolidated_decisions_round{round_no}', conn, 
                       if_exists='replace', index=False)

        conn.commit()
        QMessageBox.information(None, "Success",
                                f"Round {round_no} decisions uploaded successfully!")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during upload:\n{e}")

    finally:
        conn.close()

# def upload_round_decisions(round_no, goa_widget, other_widget, cons_widget):
#     conn = sqlite3.connect(DB_NAME)
#     cursor = conn.cursor()

#     try:
#         _create_decision_tables(cursor, round_no)

#         df_goa = goa_widget.get_mapped_dataframe()

#         df_goa = auto_match_columns(df_goa, [
#             "mtech_app_no", 
#             "applicant_decision"
#         ])

#         df_goa.to_sql(f'iit_goa_offers_round{round_no}', conn, if_exists='replace', index=False)

#         df_other = other_widget.get_mapped_dataframe()

#         df_other = auto_match_columns(df_other, [
#             "mtech_app_no",
#             "other_institute_decision"
#         ])

#         df_other.to_sql(f'accepted_other_institute_round{round_no}', conn, if_exists='replace', index=False)

#         df_cons = cons_widget.get_mapped_dataframe()

#         df_cons = auto_match_columns(df_cons, [
#             "coap_reg_id",
#             "applicant_decision"
#         ])

#         df_cons.to_sql(f'consolidated_decisions_round{round_no}', conn, if_exists='replace', index=False)

#         conn.commit()
#         QMessageBox.information(None, "Success", f"Round {round_no} decisions uploaded successfully!")

#     except Exception as e:
#         QMessageBox.critical(None, "Error", f"Error during upload:\n{e}")

#     finally:
#         conn.close()


# ------------------------------------------------------
# Determine candidates eligible for next round
# ------------------------------------------------------
def _get_eligible_candidates_for_next_round(current_round):
    conn = sqlite3.connect(DB_NAME)
    coaps_out = set()

    # All Accept & Freeze -> permanently out
    for r in range(1, current_round + 1):
        df_goa_frozen = pd.read_sql_query(f"""
            SELECT c.COAP
            FROM candidates c
            JOIN iit_goa_offers_round{r} d ON d.mtech_app_no = c.App_no
            WHERE d.applicant_decision = 'Accept and Freeze'
        """, conn)
        coaps_out.update(df_goa_frozen['COAP'].tolist())

        df_cons_frozen = pd.read_sql_query(f"""
            SELECT coap_reg_id
            FROM consolidated_decisions_round{r}
            WHERE applicant_decision = 'Accept and Freeze'
        """, conn)
        coaps_out.update(df_cons_frozen['coap_reg_id'].tolist())

    # Reject & Wait → removed from future rounds
    for r in range(1, current_round + 1):
        df_rejected = pd.read_sql_query(f"""
            SELECT c.COAP
            FROM candidates c
            JOIN iit_goa_offers_round{r} d ON d.mtech_app_no = c.App_no
            WHERE d.applicant_decision = 'Reject and Wait'
        """, conn)
        coaps_out.update(df_rejected['COAP'].tolist())

    # All candidates with GATE score
    df_all = pd.read_sql_query("""
        SELECT COAP
        FROM candidates
        WHERE MaxGATEScore_3yrs IS NOT NULL
    """, conn)

    conn.close()

    all_coaps = set(df_all['COAP'].tolist())
    eligible_coaps = list(all_coaps - coaps_out)

    return eligible_coaps


# ------------------------------------------------------
# Compute confirmed seats across previous rounds
# ------------------------------------------------------
def _recalculate_confirmed_seats(last_round, conn):
    if last_round < 1:
        return {}

    confirmed = {}

    for r in range(1, last_round + 1):
        df = pd.read_sql_query(f"""
            SELECT o.category, COUNT(o.COAP) as count
            FROM offers o
            JOIN candidates c ON o.COAP = c.COAP
            JOIN iit_goa_offers_round{r} d ON d.mtech_app_no = c.App_no
            WHERE o.round_no = {r} 
              AND d.applicant_decision = 'Accept and Freeze'
            GROUP BY o.category
        """, conn)

        for _, row in df.iterrows():
            confirmed[row['category']] = confirmed.get(row['category'], 0) + row['count']

    return confirmed


# ------------------------------------------------------
# Load seat matrix and apply confirmed seats
# ------------------------------------------------------
def _get_seat_matrix_with_confirmed(conn, confirmed_seats):
    cursor = conn.cursor()
    cursor.execute("SELECT category, set_seats, seats_allocated FROM seat_matrix")

    return {
        cat.strip(): {
            "total": total or 0,
            "allocated": confirmed_seats.get(cat.strip(), 0)
        }
        for cat, total, alloc in cursor.fetchall()
    }


# ------------------------------------------------------
# Retrieve candidates who selected "Retain and Wait"
# ------------------------------------------------------
def _get_retained_candidates(previous_round, conn):
    if previous_round < 1:
        return {}

    df = pd.read_sql_query(f"""
        SELECT o.COAP, o.category
        FROM offers o
        JOIN candidates c ON o.COAP = c.COAP
        JOIN iit_goa_offers_round{previous_round} d
            ON d.mtech_app_no = c.App_no
        WHERE o.round_no = {previous_round}
          AND d.applicant_decision = 'Retain and Wait'
    """, conn)

    return df.set_index("COAP")["category"].to_dict()

def _get_upgraded_candidates(previous_round, conn):
    if previous_round < 1:
        return {}
    
    # Logic to fetch candidates who made 'Retain and Wait' OR 'Accept and Wait'
    # decisions in the previous round, as both are eligible for upgrade/reallocation.
    df = pd.read_sql_query(f"""
        SELECT o.COAP, o.category
        FROM offers o
        JOIN candidates c ON o.COAP = c.COAP
        JOIN iit_goa_offers_round{previous_round} d 
            ON d.mtech_app_no = c.App_no
        WHERE o.round_no = {previous_round}
          AND d.applicant_decision IN ('Retain and Wait', 'Accept and Wait')
    """, conn)
    return df.set_index("COAP")["category"].to_dict()

def auto_match_columns(df, required_cols):
    """
    Automatically match uploaded DataFrame columns to required DB columns.

    - Exact match preferred
    - If not found, fuzzy match (similar/related)
    - Ignores extra columns
    - Raises error if no possible match is found
    """

    col_map = {}
    uploaded_cols = [c.lower().strip().replace(" ", "_") for c in df.columns]

    for req in required_cols:
        req_norm = req.lower()

        # --- 1. Exact match ---
        if req_norm in uploaded_cols:
            idx = uploaded_cols.index(req_norm)
            col_map[df.columns[idx]] = req
            continue

        # --- 2. Fuzzy match ---
        close = difflib.get_close_matches(req_norm, uploaded_cols, n=1, cutoff=0.55)

        if close:
            idx = uploaded_cols.index(close[0])
            col_map[df.columns[idx]] = req
            continue

        # --- 3. No match found → error ---
        raise ValueError(f"Required column '{req}' not found or matched in uploaded file.")

    # Apply rename
    df = df.rename(columns=col_map)

    # Keep only required columns
    df = df[required_cols]

    return df
# ------------------------------------------------------
# MAIN ROUND ALLOCATION (Fixed PWD Priority Logic - Internal Reservation)
# ------------------------------------------------------
def run_round(round_no):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    try:
        previous_round = round_no - 1

        # ======================================================
        # 1. Eligible candidates & Retained/Upgraded Map
        # (No change here)
        # ======================================================
        if round_no == 1:
            cursor.execute("""
                SELECT COAP
                FROM candidates 
                WHERE MaxGATEScore_3yrs IS NOT NULL
            """)
            eligible = [row[0] for row in cursor.fetchall()]
        else:
            eligible = _get_eligible_candidates_for_next_round(previous_round)

        if not eligible:
            QMessageBox.warning(None, "Round Complete", f"No eligible candidates for Round {round_no}.")
            conn.close()
            return

        upgraded_map = _get_upgraded_candidates(previous_round, conn)

        eligible_str = ", ".join([f"'{c}'" for c in eligible])

        df_cands = pd.read_sql_query(f"""
            SELECT COAP, App_no, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
            FROM candidates
            WHERE COAP IN ({eligible_str})
            ORDER BY MaxGATEScore_3yrs DESC
        """, conn)
        candidates = df_cands.values.tolist()

        # ======================================================
        # 2. Load seat matrix with confirmed seats & Setup 🛠️
        # (Modified to separate PWD quota from the seat matrix)
        # ======================================================
        confirmed = _recalculate_confirmed_seats(previous_round, conn)
        seat_matrix = _get_seat_matrix_with_confirmed(conn, confirmed)
        
        # --- NEW LOGIC: Extract COMMON_PWD and use it as a separate quota ---
        # This ensures COMMON_PWD is not counted as an extra seat.
        pwd_quota_entry = seat_matrix.pop("COMMON_PWD", {"total": 0, "allocated": 0})
        pwd_reservation_total = pwd_quota_entry["total"]
        pwd_reservation_allocated = pwd_quota_entry["allocated"] # Start from confirmed PWD seats
        # --------------------------------------------------------------------

        # Setup allocation lists
        offers_made = []
        allocated = set()

        def try_allocate(coap, name, score, seat_key, status):
            if seat_key not in seat_matrix:
                return False
            # Check if seat is available
            if seat_matrix[seat_key]["allocated"] < seat_matrix[seat_key]["total"]:
                seat_matrix[seat_key]["allocated"] += 1
                offers_made.append((round_no, coap, name, seat_key, score, status))
                allocated.add(coap)
                return True
            return False

        # ======================================================
        # 3. PWD PRE-ALLOCATION (Highest Scored PWD first) 🎯 🛠️
        # (Modified to allocate PWD candidates to their BASE category seat)
        # ======================================================
        # Filter and sort PWD candidates by score (already sorted, just filter)
        pwd_candidates = [
            (coap, app_no, name, base_cat, ews, gender, pwd_flag, score)
            for coap, app_no, name, base_cat, ews, gender, pwd_flag, score in candidates
            if (pwd_flag.strip().capitalize() if pwd_flag else "No") == "Yes"
        ]

        for coap, app_no, name, base_cat, ews, gender, pwd_flag, score in pwd_candidates:
            # Check the PWD quota
            if coap in allocated or pwd_reservation_allocated >= pwd_reservation_total:
                continue
            
            # 1. Determine the PWD candidate's BASE category seat key
            base_cat = base_cat.strip()
            gender_norm = gender.strip().capitalize() if gender else "Male"
            ews_norm = ews.strip().capitalize() if ews else "No"
            
            cat_prefix = "EWS" if ews_norm == "Yes" else base_cat
            
            # Priority: Try Female seat first, then FandM (assuming Female is internal to FandM)
            base_seat_key = f"{cat_prefix}_Female"
            
            # Fallback to FandM if Female seat key doesn't exist or is not applicable
            if base_seat_key not in seat_matrix or gender_norm != "Female":
                base_seat_key = f"{cat_prefix}_FandM"

            # 2. Allocate the PWD candidate to their BASE category seat, provided it has space.
            if try_allocate(coap, name, score, base_seat_key, "Offered (PWD Priority)"):
                # Crucial step: Increment the PWD allocated count
                pwd_reservation_allocated += 1
            
        # ======================================================
        # 4. MAIN ALLOCATION LOOP (Remaining Candidates/Seats)
        # (No change here, as PWD candidates are already handled)
        # ======================================================
        for coap, app_no, name, base_cat, ews, gender, pwd_flag, score in candidates:
            # Skip PWD candidates already allocated in Step 3
            if coap in allocated:
                continue

            # ... (rest of the Main Allocation Logic remains the same)
            base_cat = base_cat.strip()
            gender_norm = gender.strip().capitalize() if gender else "Male"
            ews_norm = ews.strip().capitalize() if ews else "No"
            
            # Set the offer status
            status = "Offered (Upgraded)" if coap in upgraded_map else "Offered"

            # --- Priority 1: General Seats ---
            general_keys = ["GEN_FandM"]
            if gender_norm == "Female":
                general_keys.insert(0, "GEN_Female")

            # --- Priority 2: Reserved/EWS Seats ---
            cat_prefix = "EWS" if ews_norm == "Yes" else base_cat
            reserved_keys = [f"{cat_prefix}_FandM"]
            if gender_norm == "Female":
                reserved_keys.insert(0, f"{cat_prefix}_Female")

            # --- Priority 3: Guaranteed Retained/Upgraded Seat ---
            upgraded_key = [upgraded_map[coap]] if coap in upgraded_map else []
            
            # Final Priority: GEN -> Own Category -> Retained
            priority = general_keys + reserved_keys + upgraded_key

            # Attempt Allocation
            for key in priority:
                if try_allocate(coap, name, score, key, status):
                    break # Candidate is allocated, move to the next candidate

        # ------------------------------------------------------
        # Save offers
        # ------------------------------------------------------
        # ... (Save offers logic remains the same)
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
        
        # The data in offers_made already reflects the base seat key allocated
        cursor.executemany("""
            INSERT OR REPLACE INTO offers 
            (round_no, COAP, Full_Name, category, MaxGATEScore_3yrs, offer_status)
            VALUES (?, ?, ?, ?, ?, ?)
        """, offers_made)

        conn.commit()
        QMessageBox.information(None, "Success", f"Round {round_no} allocation complete!\nTotal offers: {len(offers_made)}")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during round {round_no} allocation:\n{e}")

    finally:
        conn.close()
# second
# ------------------------------------------------------
# MAIN ROUND ALLOCATION (FIXED LOGIC)
# ------------------------------------------------------
# def run_round(round_no):
#     conn = sqlite3.connect(DB_NAME)
#     cursor = conn.cursor()

#     try:
#         previous_round = round_no - 1

#         # ======================================================
#         # 1. Eligible candidates
#         # ======================================================
#         if round_no == 1:
#             cursor.execute("""
#                 SELECT COAP
#                 FROM candidates 
#                 WHERE MaxGATEScore_3yrs IS NOT NULL
#             """)
#             eligible = [row[0] for row in cursor.fetchall()]
#         else:
#             eligible = _get_eligible_candidates_for_next_round(previous_round)

#         if not eligible:
#             QMessageBox.warning(None, "Round Complete", f"No eligible candidates for Round {round_no}.")
#             conn.close()
#             return

#         # Retained candidates
#         retained_map = _get_retained_candidates(previous_round, conn)

#         # Fetch sorted candidate data (by score DESC)
#         eligible_str = ", ".join([f"'{c}'" for c in eligible])

#         df_cands = pd.read_sql_query(f"""
#             SELECT COAP, Full_Name, Category, Ews, Gender, MaxGATEScore_3yrs
#             FROM candidates
#             WHERE COAP IN ({eligible_str})
#             ORDER BY MaxGATEScore_3yrs DESC
#         """, conn)
#         candidates = df_cands.values.tolist()

#         # ======================================================
#         # 2. Load seat matrix with confirmed seats
#         # ======================================================
#         confirmed = _recalculate_confirmed_seats(previous_round, conn)
#         seat_matrix = _get_seat_matrix_with_confirmed(conn, confirmed)

#         # ======================================================
#         # Prepare offers table
#         # ======================================================
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
#         allocated = set()

#         # ------------------------------------------------------
#         # Helper allocation
#         # ------------------------------------------------------
#         def try_allocate(coap, name, score, seat_key, status):
#             if seat_key not in seat_matrix:
#                 return False
#             # Check if seat is available
#             if seat_matrix[seat_key]["allocated"] < seat_matrix[seat_key]["total"]:
#                 seat_matrix[seat_key]["allocated"] += 1
#                 offers_made.append((round_no, coap, name, seat_key, score, status))
#                 allocated.add(coap)
#                 return True
#             return False

#         # ------------------------------------------------------
#         # 3. UNIFIED ALLOCATION LOOP (Ensures Merit Priority)
#         # ------------------------------------------------------
#         for coap, name, base_cat, ews, gender, score in candidates:
#             if coap in allocated:
#                 continue

#             base_cat = base_cat.strip()
#             gender_norm = gender.strip().capitalize() if gender else "Male"
#             ews_norm = ews.strip().capitalize() if ews else "No"
            
#             # 1. GENERAL Seats (Highest Priority for everyone)
#             general_keys = ["GEN_FandM"]
#             if gender_norm == "Female":
#                 general_keys.insert(0, "GEN_Female")

#             # 2. CANDIDATE'S OWN Category Seats (OBC/SC/ST/EWS)
#             # Fallback if the candidate fails to secure a General seat.
#             cat_prefix = "EWS" if ews_norm == "Yes" else base_cat
#             category_keys = [f"{cat_prefix}_FandM"]
#             if gender_norm == "Female":
#                 category_keys.insert(0, f"{cat_prefix}_Female")

#             # 3. RETAINED Seat (Guaranteed spot if retained previously)
#             retained_key = [retained_map[coap]] if coap in retained_map else []
            
#             # Final Priority: GEN -> Own Category -> Retained (if applicable)
#             # This order ensures the highest-scoring candidate is allocated the
#             # most general seat they qualify for first.
#             priority = general_keys + category_keys + retained_key
            
#             # Set the offer status
#             status = "Offered (Retained/Upgrade)" if coap in retained_map else "Offered"

#             # Attempt Allocation
#             for key in priority:
#                 if try_allocate(coap, name, score, key, status):
#                     break # Candidate is allocated, move to the next candidate

#         # ------------------------------------------------------
#         # Save offers
#         # ------------------------------------------------------
#         cursor.executemany("""
#             INSERT OR REPLACE INTO offers 
#             (round_no, COAP, Full_Name, category, MaxGATEScore_3yrs, offer_status)
#             VALUES (?, ?, ?, ?, ?, ?)
#         """, offers_made)

#         conn.commit()
#         QMessageBox.information(None, "Success", f"Round {round_no} allocation complete!\nTotal offers: {len(offers_made)}")

#     except Exception as e:
#         QMessageBox.critical(None, "Error", f"Error during round {round_no} allocation:\n{e}")

#     finally:
#         conn.close()

# first
# # ------------------------------------------------------
# # MAIN ROUND ALLOCATION
# # ------------------------------------------------------
# def run_round(round_no):
#     conn = sqlite3.connect(DB_NAME)
#     cursor = conn.cursor()

#     try:
#         previous_round = round_no - 1

#         # ======================================================
#         # 1. Eligible candidates
#         # ======================================================
#         if round_no == 1:
#             cursor.execute("""
#                 SELECT COAP
#                 FROM candidates 
#                 WHERE MaxGATEScore_3yrs IS NOT NULL
#             """)
#             eligible = [row[0] for row in cursor.fetchall()]
#         else:
#             eligible = _get_eligible_candidates_for_next_round(previous_round)

#         if not eligible:
#             QMessageBox.warning(None, "Round Complete", f"No eligible candidates for Round {round_no}.")
#             conn.close()
#             return

#         # Retained candidates
#         retained_map = _get_retained_candidates(previous_round, conn)

#         # Fetch sorted candidate data
#         eligible_str = ", ".join([f"'{c}'" for c in eligible])

#         df_cands = pd.read_sql_query(f"""
#             SELECT COAP, Full_Name, Category, Ews, Gender, MaxGATEScore_3yrs
#             FROM candidates
#             WHERE COAP IN ({eligible_str})
#             ORDER BY MaxGATEScore_3yrs DESC
#         """, conn)
#         candidates = df_cands.values.tolist()

#         # ======================================================
#         # 2. Load seat matrix with confirmed seats
#         # ======================================================
#         confirmed = _recalculate_confirmed_seats(previous_round, conn)
#         seat_matrix = _get_seat_matrix_with_confirmed(conn, confirmed)

#         # ======================================================
#         # Prepare offers table
#         # ======================================================
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
#         allocated = set()

#         # ------------------------------------------------------
#         # Helper allocation
#         # ------------------------------------------------------
#         def try_allocate(coap, name, score, seat_key, status):
#             if seat_key not in seat_matrix:
#                 return False
#             if seat_matrix[seat_key]["allocated"] < seat_matrix[seat_key]["total"]:
#                 seat_matrix[seat_key]["allocated"] += 1
#                 offers_made.append((round_no, coap, name, seat_key, score, status))
#                 allocated.add(coap)
#                 return True
#             return False

#         # ------------------------------------------------------
#         # 3. Allocate retained first
#         # ------------------------------------------------------
#         for coap, name, base_cat, ews, gender, score in candidates:
#             if coap not in retained_map:
#                 continue

#             retained_category = retained_map[coap]

#             # Try upgrade first
#             base_cat = base_cat.strip()
#             gender_norm = gender.strip().capitalize() if gender else "Male"
#             ews_norm = ews.strip().capitalize() if ews else "No"

#             # GENERAL seats first
#             general_keys = ["GEN_Female", "GEN_FandM"] if gender_norm == "Female" else ["GEN_FandM"]

#             # Category seats
#             cat_prefix = "EWS" if ews_norm == "Yes" else base_cat
#             category_keys = (
#                 [f"{cat_prefix}_Female", f"{cat_prefix}_FandM"]
#                 if gender_norm == "Female"
#                 else [f"{cat_prefix}_FandM"]
#             )

#             priority = general_keys + category_keys + [retained_category]

#             for key in priority:
#                 if try_allocate(coap, name, score, key, "Offered (Retained/Upgrade)"):
#                     break

#         # ------------------------------------------------------
#         # 4. Allocate remaining candidates (non-PWD only)
#         # ------------------------------------------------------
#         for coap, name, base_cat, ews, gender, score in candidates:
#             if coap in allocated:
#                 continue

#             base_cat = base_cat.strip()
#             gender_norm = gender.strip().capitalize() if gender else "Male"
#             ews_norm = ews.strip().capitalize() if ews else "No"

#             general_keys = ["GEN_Female", "GEN_FandM"] if gender_norm == "Female" else ["GEN_FandM"]

#             cat_prefix = "EWS" if ews_norm == "Yes" else base_cat
#             category_keys = (
#                 [f"{cat_prefix}_Female", f"{cat_prefix}_FandM"]
#                 if gender_norm == "Female"
#                 else [f"{cat_prefix}_FandM"]
#             )

#             for key in general_keys + category_keys:
#                 if try_allocate(coap, name, score, key, "Offered"):
#                     break

#         # ------------------------------------------------------
#         # Save offers
#         # ------------------------------------------------------
#         cursor.executemany("""
#             INSERT OR REPLACE INTO offers 
#             (round_no, COAP, Full_Name, category, MaxGATEScore_3yrs, offer_status)
#             VALUES (?, ?, ?, ?, ?, ?)
#         """, offers_made)

#         conn.commit()
#         QMessageBox.information(None, "Success", f"Round {round_no} allocation complete!\nTotal offers: {len(offers_made)}")

#     except Exception as e:
#         QMessageBox.critical(None, "Error", f"Error during round {round_no} allocation:\n{e}")

#     finally:
#         conn.close()


# ------------------------------------------------------
# Download offers for a round
# ------------------------------------------------------

# def download_offers(round_no):
#     """
#     Generates a comprehensive report for the given round (N).
#     The report includes:
#     1. New offers/upgrades made in the current round (N).
#     2. Seats held by candidates who 'Accepted and Freeze' in all previous rounds (< N).
#     """
#     conn = sqlite3.connect(DB_NAME)
    
#     # --- 1. Dynamically Build the Query to Find ALL Frozen Candidates ---
    
#     # This list will hold the SELECT statements for each previous round's 'Accept and Freeze' status
#     frozen_unions = []
    
#     # Iterate from Round 1 up to (but not including) the current round_no
#     for prev_round in range(1, round_no):
#         iit_goa_table = f"iit_goa_offers_round{prev_round}"
        
#         # 1. Select the COAP from the decision table where status is 'Accept and Freeze'
#         # 2. Add the round number where they froze the offer
#         frozen_unions.append(f"""
#             SELECT T1."COAP", '{prev_round}' AS accepted_round
#             FROM (SELECT DISTINCT T2.COAP FROM offers T2 WHERE T2.round_no = {prev_round}) T1 -- COAP in offers table
#             INNER JOIN {iit_goa_table} T3 
#                 ON T3.App_no = T1.COAP -- Assuming the decision table uses App_no/COAP for join
#             WHERE T3.Status = 'Accept and Freeze'
#         """)

#     # Combine the individual round SELECT statements using UNION ALL
#     all_frozen_coaps_cte = " UNION ALL ".join(frozen_unions)
    
#     # --- 2. Construct the Final Comprehensive SQL Query ---
    
#     sql_query = f"""
#         WITH AllFrozenCandidates AS (
#             -- This CTE combines all COAP IDs that have chosen 'Accept and Freeze' in rounds 1 to N-1
#             {all_frozen_coaps_cte}
#         ),
        
#         LastFrozenOffer AS (
#             -- This CTE finds the *specific* offer details (category, score) for the frozen candidate
#             SELECT
#                 T1.COAP,
#                 MAX(T1.round_no) AS max_round
#             FROM
#                 offers T1
#                 INNER JOIN AllFrozenCandidates T2 ON T1.COAP = T2.COAP
#             GROUP BY
#                 T1.COAP
#         )
        
#         -- PART 1: Previously Accepted/Frozen Candidates
#         SELECT
#             T3.COAP,
#             T3.Full_Name,
#             T3.MaxGATEScore_3yrs AS GATEScore,
#             T3.category AS category_offered, 
#             '{round_no}' AS round_no_report, -- Use current round number for display
#             'OFFER HELD (ACCEPTED)' AS offer_status
#         FROM
#             offers T3
#             INNER JOIN LastFrozenOffer LFO ON 
#                 T3.COAP = LFO.COAP AND T3.round_no = LFO.max_round
            
#         UNION ALL
        
#         -- PART 2: New Offers/Upgrades for the Current Round (round_no)
#         SELECT
#             COAP,
#             Full_Name,
#             MaxGATEScore_3yrs AS GATEScore,
#             category AS category_offered,
#             round_no AS round_no_report,
#             offer_status
#         FROM
#             offers
#         WHERE
#             round_no = {round_no}
            
#         ORDER BY
#             GATEScore DESC;
#     """
    
#     # ----------------------------------------------------

#     try:
#         # 3. Execute the query and load into a DataFrame
#         df_report = pd.read_sql_query(sql_query, conn)

#         # 4. Save the DataFrame to an Excel file
#         file_name = f"Round_{round_no}_Offers_Report.xlsx"
#         df_report.to_excel(file_name, index=False)
        
#         QMessageBox.information(None, "Download Complete", f"Offers for Round {round_no} downloaded to:\n{file_name}")

#     except Exception as e:
#         QMessageBox.critical(None, "Download Error", f"Could not generate offer report:\n{e}")

#     finally:
#         conn.close()
# def download_offers(round_no=1):
#     conn = sqlite3.connect(DB_NAME)

#     df_summary = pd.read_sql_query(f"""
#         SELECT o.round_no, c.COAP, o.Full_Name, o.category, o.MaxGATEScore_3yrs, o.offer_status
#         FROM offers o
#         JOIN candidates c ON o.COAP = c.COAP
#         WHERE o.round_no = {round_no}
#     """, conn)

#     if df_summary.empty:
#         QMessageBox.warning(None, "No Offers", f"No offers found for Round {round_no}")
#         conn.close()
#         return

#     df_details = pd.read_sql_query(f"""
#         SELECT o.round_no, c.*
#         FROM offers o
#         JOIN candidates c ON o.COAP = c.COAP
#         WHERE o.round_no = {round_no}
#         ORDER BY o.MaxGATEScore_3yrs DESC
#     """, conn)

#     conn.close()

#     filename = f"Round{round_no}_Offers.xlsx"

#     with pd.ExcelWriter(filename, engine='openpyxl') as writer:
#         df_summary.to_excel(writer, sheet_name="Offers_Summary", index=False)
#         df_details.to_excel(writer, sheet_name="Offers_Detailed", index=False)

#     QMessageBox.information(None, "Download Complete", f"Offers saved as {filename}")
# def download_offers(round_no):
#     """
#     Generates a comprehensive report for the given round (N).
#     The report includes:
#     1. New offers/upgrades made in the current round (N).
#     2. Seats held by candidates who 'Accept and Freeze' in all previous rounds (< N).
#     """
#     conn = sqlite3.connect(DB_NAME)
    
#     # --- 1. Dynamically Build the Query to Find ALL Frozen Candidates ---
    
#     frozen_unions = []
    
#     # Iterate from Round 1 up to (but not including) the current round_no
#     for prev_round in range(1, round_no):
#         iit_goa_table = f"iit_goa_offers_round{prev_round}"
        
#         # We target the 'iit_goa_offers_roundX' table for the 'Accept and Freeze' status
#         frozen_unions.append(f"""
#             SELECT T1.COAP, '{prev_round}' AS accepted_round
#             FROM (SELECT DISTINCT T2.COAP FROM offers T2 WHERE T2.round_no = {prev_round}) T1
#             INNER JOIN {iit_goa_table} T3 
#                 -- Use the column names identified from your PRAGMA query
#                 ON T3.mtech_app_no = T1.COAP 
#             WHERE 
#                 T3.applicant_decision = 'Accept and Freeze'
#         """)

#     # Handle the case where round_no is 1 (no previous rounds)
#     if not frozen_unions:
#         all_frozen_coaps_cte = "SELECT NULL AS COAP, NULL AS accepted_round WHERE 1 = 0" 
#     else:
#         all_frozen_coaps_cte = " UNION ALL ".join(frozen_unions)
    
#     # --- 2. Construct the Final Comprehensive SQL Query ---
    
#     sql_query = f"""
#         WITH AllFrozenCandidates AS (
#             {all_frozen_coaps_cte}
#         ),
        
#         LastFrozenOffer AS (
#             -- Find the specific offer details (category, score) for the most recent frozen offer
#             SELECT
#                 T1.COAP,
#                 MAX(T1.round_no) AS max_round
#             FROM
#                 offers T1
#                 INNER JOIN AllFrozenCandidates T2 ON T1.COAP = T2.COAP
#             GROUP BY
#                 T1.COAP
#         )
        
#         -- PART 1: Previously Accepted/Frozen Candidates (The seats that are held)
#         SELECT
#             T3.round_no AS round_no,
#             T3.COAP,
#             T3.Full_Name,
#             T3.category,
#             T3.MaxGATEScore_3yrs AS GATEScore,
#             -- Display status as 'OFFER HELD (ACCEPTED)' for reporting clarity
#             'OFFER HELD (ACCEPTED)' AS offer_status
#         FROM
#             offers T3
#             INNER JOIN LastFrozenOffer LFO ON 
#                 T3.COAP = LFO.COAP AND T3.round_no = LFO.max_round
            
#         UNION ALL
        
#         -- PART 2: New Offers/Upgrades for the Current Round (round_no)
#         SELECT
#             round_no,
#             COAP,
#             Full_Name,
#             category,
#             MaxGATEScore_3yrs AS GATEScore,
#             offer_status
#         FROM
#             offers
#         WHERE
#             round_no = {round_no}
            
#         ORDER BY
#             GATEScore DESC;
#     """
    
#     # ----------------------------------------------------

#     try:
#         # 3. Execute the query and load into a DataFrame
#         df_report = pd.read_sql_query(sql_query, conn)
        
#         # 4. Rename columns to match your desired Excel output format
#         df_report.rename(columns={
#             'round_no': 'round_no',  # This now contains the original offer round or current round
#             'GATEScore': 'GATEScore',
#             'category': 'category',
#             'offer_status': 'offer_status'
#         }, inplace=True)

#         # 5. Save the DataFrame to an Excel file
#         file_name = f"Round_{round_no}_Offers_Report.xlsx"
#         df_report.to_excel(file_name, index=False)
        
#         QMessageBox.information(None, "Download Complete", f"Offers for Round {round_no} downloaded to:\n{file_name}")

#     except Exception as e:
#         QMessageBox.critical(None, "Download Error", f"Could not generate offer report:\nExecution failed on sql\n{e}")

#     finally:
#         conn.close()
# small helper (place with other helpers)
def _safe_sql_df(conn, sql):
    """Run SQL and return DataFrame; re-raise exceptions to be handled by caller."""
    try:
        return pd.read_sql_query(sql, conn)
    except Exception:
        raise

# -------------------------
# Helper: collect frozen COAPs (Accept & Freeze)
# -------------------------
def _get_frozen_coaps(upto_round):
    """
    Return set of mtech_app_no/COAP strings that selected 'Accept and Freeze'
    in iit_goa_offers_round1..upto_round (defensive for column name variants).
    """
    conn = sqlite3.connect(DB_NAME)
    frozen = set()
    try:
        for r in range(1, upto_round + 1):
            table = f"iit_goa_offers_round{r}"
            # try both common column names
            for colname in ("applicant_decision", "Status"):
                try:
                    sql = f"SELECT mtech_app_no FROM {table} WHERE {colname} = 'Accept and Freeze'"
                    df = pd.read_sql_query(sql, conn)
                    if not df.empty:
                        for v in df.iloc[:, 0].dropna().astype(str).tolist():
                            frozen.add(v)
                    break
                except Exception:
                    continue
    finally:
        conn.close()
    return frozen


def _table_exists(conn, name):
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (name,))
    return cur.fetchone() is not None

def download_offers(round_no, out_filename=None):
    """
    Final exporter implementing:
    - Offers-List (with offerCat and IsPWD first columns)
    - Offer-24-format (detailed offered rows)
    - PWD-GEN, PWD-OBC, PWD-SC, PWD-ST, PWD-EWS
    - General-ALL (all offered), General-Female (all females)
    - EWS, OBC, OBC-Female, SC, SC-Female, ST, ST-Female
    Notes:
    - 'offered' and 'accepted' columns are 'Y' or ''.
    - IsPWD values are 'Yes' / 'No'.
    - offerCat is pulled from the offers table (if present); otherwise blank.
    - Adm_cat (if present) is removed from final outputs.
    """
    if out_filename is None:
        out_filename = f"Round_{round_no}_Offers_Report.xlsx"

    conn = sqlite3.connect(DB_NAME)
    try:
        # --- load base tables ---
        df_cand = pd.read_sql_query("SELECT * FROM candidates", conn)

        # normalize COAP column name
        if 'COAP' not in df_cand.columns and 'App_no' in df_cand.columns:
            df_cand = df_cand.rename(columns={'App_no': 'COAP'})

        # --- load offers for this round ---
        if _table_exists(conn, "offers"):
            df_off = pd.read_sql_query(f"SELECT * FROM offers WHERE round_no = {int(round_no)}", conn)
        else:
            df_off = pd.DataFrame(columns=["round_no", "COAP", "Full_Name", "category", "MaxGATEScore_3yrs", "offer_status"])

        # build map COAP -> offered category (offerCat) from df_off (if a column exists)
        offer_cat_col = None
        for cand in ('category', 'offer_category', 'OfferedCategory', 'Offered_Cat', 'OfferCat'):
            if cand in df_off.columns:
                offer_cat_col = cand
                break
        if offer_cat_col is None:
            # fallback: any column name containing 'cat'
            for c in df_off.columns:
                if 'cat' in c.lower():
                    offer_cat_col = c
                    break

        if offer_cat_col is not None and 'COAP' in df_off.columns:
            offer_cat_map = dict(zip(df_off['COAP'].astype(str), df_off[offer_cat_col].astype(str)))
        else:
            offer_cat_map = {}

        # --- accumulate decisions up to round_no (to compute 'accepted') ---
        decisions_frames = []
        for r in range(1, int(round_no) + 1):
            table = f"iit_goa_offers_round{r}"
            if _table_exists(conn, table):
                try:
                    dfd = pd.read_sql_query(f"SELECT mtech_app_no AS COAP, applicant_decision FROM {table}", conn)
                    dfd['round'] = r
                    decisions_frames.append(dfd)
                except Exception:
                    pass
            cons_table = f"consolidated_decisions_round{r}"
            if _table_exists(conn, cons_table):
                try:
                    dfc = pd.read_sql_query(f"SELECT coap_reg_id AS COAP, applicant_decision FROM {cons_table}", conn)
                    dfc['round'] = r
                    decisions_frames.append(dfc)
                except Exception:
                    pass

        if decisions_frames:
            df_decisions = pd.concat(decisions_frames, ignore_index=True)
        else:
            df_decisions = pd.DataFrame(columns=["COAP", "applicant_decision", "round"])

        accepted_coaps = set(df_decisions.loc[df_decisions['applicant_decision'] == 'Accept and Freeze', 'COAP'].astype(str).tolist())
        offered_coaps = set(df_off['COAP'].astype(str).tolist())

        # --- build master dataframe ---
        df_master = df_cand.copy()
        if 'COAP' not in df_master.columns and 'App_no' in df_master.columns:
            df_master = df_master.rename(columns={'App_no': 'COAP'})
        df_master['COAP'] = df_master['COAP'].astype(str)

        # offered / accepted flags
        df_master['offered'] = df_master['COAP'].apply(lambda x: 'Y' if x in offered_coaps else '')
        df_master['accepted'] = df_master['COAP'].apply(lambda x: 'Y' if x in accepted_coaps else '')

        # offerCat column
        df_master['offerCat'] = df_master['COAP'].apply(lambda x: offer_cat_map.get(x, '')).fillna('')

        # IsPWD column: detect existing PWD-like column names, otherwise fallback to 'Pwd'
        pwd_col = None
        for c in df_master.columns:
            if c.lower() in ('pwd', 'is_pwd', 'ispwd', 'physically_disabled'):
                pwd_col = c
                break
        if pwd_col is None and 'Pwd' in df_master.columns:
            pwd_col = 'Pwd'
        if pwd_col:
            df_master['IsPWD'] = df_master[pwd_col].fillna('').astype(str).apply(
                lambda s: 'Yes' if s.strip().upper() in ('YES', 'Y', '1', 'TRUE') else 'No'
            )
        else:
            df_master['IsPWD'] = 'No'  # default when no column present
        # --- canonicalize gate score column so all sheets have it ---
        gate_candidates = [
            'MaxGATEScore_3yrs', 'MaxGateS', 'GATE Score', 'GATE25',
            'GATE_Reg', 'GATE25RollN', 'GATE', 'GATE25Roll'
        ]

        # 1) Already exists in master?
        for gcol in gate_candidates:
            if gcol in df_master.columns:
                df_master['MaxGATEScore_3yrs'] = df_master[gcol]
                break

        # 2) Otherwise try from offers table
        if 'MaxGATEScore_3yrs' not in df_master.columns or df_master['MaxGATEScore_3yrs'].isnull().all():
            for gcol in gate_candidates:
                if gcol in df_off.columns:
                    offer_gate_map = dict(zip(df_off['COAP'].astype(str), df_off[gcol]))
                    df_master['MaxGATEScore_3yrs'] = df_master['COAP'].map(offer_gate_map)
                    break

        if 'MaxGATEScore_3yrs' not in df_master.columns:
            df_master['MaxGATEScore_3yrs'] = ''

        try:
            df_master['MaxGATEScore_3yrs'] = pd.to_numeric(df_master['MaxGATEScore_3yrs'], errors='coerce')
        except:
            pass


        # Normalized helper columns for filtering
        def up(x): return (x or '').strip().upper()
        df_master['CAT_NORM'] = df_master.get('Category', '').apply(up)
        df_master['GENDER_NORM'] = df_master.get('Gender', '').apply(up)
        df_master['EWS_NORM'] = df_master.get('Ews', '').apply(up)
        df_master['PWD_NORM'] = df_master.get(pwd_col if pwd_col else 'Pwd', '').apply(up)

        # Remove Adm_cat (if present)
        for drop_col in ('Adm_cat', 'adm_cat', 'AdmCat'):
            if drop_col in df_master.columns:
                df_master = df_master.drop(columns=[drop_col])

        # Prepare ordering: offered/IsPWD/offerCat/accepted first (Offers-List special)
        all_cols = df_master.columns.tolist()
        front_default = ['offered', 'accepted']
        front_offer_list = ['offerCat', 'IsPWD', 'offered', 'accepted']
        # ensure only existing columns are kept
        front_default = [c for c in front_default if c in all_cols]
        front_offer_list = [c for c in front_offer_list if c in all_cols]

        # Sort offered rows by score (if available)
        score_col = 'MaxGATEScore_3yrs' if 'MaxGATEScore_3yrs' in df_master.columns else None
        if score_col:
            df_offered_sorted = df_master[df_master['offered'] == 'Y'].sort_values(by=[score_col], ascending=False)
        else:
            df_offered_sorted = df_master[df_master['offered'] == 'Y']
        # -----------------------------
        # Build Offer-24-format (detailed sheet from offered rows)
        # -----------------------------
        # merge offered rows with master to pick candidate details
        _merge_left = df_offered_sorted.reset_index(drop=True)
        merged = _merge_left.merge(df_master, on='COAP', how='left', suffixes=('', '_cand')) if 'COAP' in _merge_left.columns else _merge_left.copy()

        def col_or_blank(df, col):
            return df[col] if col in df.columns else pd.Series([''] * len(df), index=df.index)

        # pick best gate column available
        gate_candidates = ['MaxGATEScore_3yrs', 'GATE Score', 'GATE25', 'GATE_Reg', 'MaxGateS', 'GATE25RollN', 'GATE']
        gate_col = next((c for c in gate_candidates if c in merged.columns), None)

        df_offer24 = pd.DataFrame(index=merged.index)
        # Application Seq No - try common names
        df_offer24['Application Seq No'] = col_or_blank(merged, 'Si_NO').where(col_or_blank(merged, 'Si_NO') != '', col_or_blank(merged, 'AppNo')).where(col_or_blank(merged, 'AppNo') != '', col_or_blank(merged, 'App_no'))
        df_offer24['AppStatus'] = 'Pending'
        df_offer24['Remarks'] = ''
        df_offer24['App Date'] = pd.Timestamp.now().strftime('%d/%b/%Y')

        df_offer24['GATE Reg'] = merged[gate_col] if gate_col in merged.columns else ''
        df_offer24['Mtech Application Number'] = col_or_blank(merged, 'App_no').where(col_or_blank(merged, 'App_no') != '', col_or_blank(merged, 'Mtech Application Number')).where(col_or_blank(merged, 'Mtech Application Number') != '', col_or_blank(merged, 'AppNo'))
        df_offer24['GATE Score'] = merged[gate_col] if gate_col in merged.columns else ''

        # Candidate name heuristics
        name_col = 'Full_Name' if 'Full_Name' in merged.columns else ('FullName' if 'FullName' in merged.columns else ( 'Name' if 'Name' in merged.columns else None))
        df_offer24['Candidate Name'] = merged[name_col] if name_col and name_col in merged.columns else col_or_blank(merged, 'FullName')

        # Program/Category/Institute columns (best-effort)
        df_offer24['Offered Program'] = merged.get('Offered Program', merged.get('branch', ''))
        df_offer24['Offered Program Code'] = merged.get('Offered Program Code', merged.get('branch', ''))
        df_offer24['Offered Category'] = merged.get('category', merged.get('offerCat', merged.get('offer_category', '')))
        df_offer24['Round No'] = int(round_no)
        df_offer24['Institute Name'] = merged.get('Institute Name', 'IIT Goa') if 'Institute Name' in merged.columns else 'IIT Goa'
        df_offer24['Institute ID'] = merged.get('Institute ID', '') if 'Institute ID' in merged.columns else ''
        df_offer24['Institute Type'] = merged.get('Institute Type', 'IIT') if 'Institute Type' in merged.columns else 'IIT'
        df_offer24['Form status'] = ''

        # final column order (only keep columns that exist)
        cols_order = [
            'Application Seq No','AppStatus','Remarks','App Date','GATE Reg','Mtech Application Number',
            'GATE Score','Candidate Name','Offered Program','Offered Program Code','Offered Category',
            'Round No','Institute Name','Institute ID','Institute Type','Form status'
        ]
        df_offer24 = df_offer24[[c for c in cols_order if c in df_offer24.columns]]


        df_non_offered = df_master[df_master['offered'] != 'Y']

        # Helper: sheet-safe name
        def sheet_name_safe(name):
            return name if len(name) <= 31 else name[:31]

        # Build sheet factories in requested exact order
        sheet_factories = [
            ('Offers-List', lambda: df_offered_sorted),
            ('Offer-24-format', lambda: df_offer24.copy()),

            ('PWD-GEN', lambda: df_master[(df_master['PWD_NORM'] == 'YES') & (df_master['CAT_NORM'] == 'GEN')]),
            ('PWD-OBC', lambda: df_master[(df_master['PWD_NORM'] == 'YES') & (df_master['CAT_NORM'] == 'OBC')]),
            ('PWD-SC', lambda: df_master[(df_master['PWD_NORM'] == 'YES') & (df_master['CAT_NORM'] == 'SC')]),
            ('PWD-ST', lambda: df_master[(df_master['PWD_NORM'] == 'YES') & (df_master['CAT_NORM'] == 'ST')]),
            ('PWD-EWS', lambda: df_master[(df_master['PWD_NORM'] == 'YES') & ((df_master['EWS_NORM'] == 'YES') | (df_master['CAT_NORM'] == 'EWS'))]),

            ('General-ALL', lambda: pd.concat([
    df_master[df_master['offered'] == 'Y'].sort_values(by=[score_col], ascending=False) if score_col else df_master[df_master['offered'] == 'Y'],
    df_master[df_master['offered'] != 'Y']
], ignore_index=True)),

            ('General-Female', lambda: df_master[df_master['GENDER_NORM'] == 'FEMALE']),

            ('EWS', lambda: df_master[(df_master['EWS_NORM'] == 'YES') | (df_master['CAT_NORM'] == 'EWS')]),

            ('OBC', lambda: df_master[df_master['CAT_NORM'] == 'OBC']),
            ('OBC-Female', lambda: df_master[(df_master['CAT_NORM'] == 'OBC') & (df_master['GENDER_NORM'] == 'FEMALE')]),

            ('SC', lambda: df_master[df_master['CAT_NORM'] == 'SC']),
            ('SC-Female', lambda: df_master[(df_master['CAT_NORM'] == 'SC') & (df_master['GENDER_NORM'] == 'FEMALE')]),

            ('ST', lambda: df_master[df_master['CAT_NORM'] == 'ST']),
            ('ST-Female', lambda: df_master[(df_master['CAT_NORM'] == 'ST') & (df_master['GENDER_NORM'] == 'FEMALE')])
        ]

        # Write sheets
        with pd.ExcelWriter(out_filename, engine='openpyxl') as writer:
            for sheet_name, factory in sheet_factories:
                try:
                    df_sheet = factory().copy()
                    # ensure DataFrame exists with columns even if empty
                    if df_sheet.empty:
                        df_sheet = pd.DataFrame(columns=all_cols)

                    # Put offered rows first in mixed sheets
                    if 'offered' in df_sheet.columns:
                        top = df_sheet[df_sheet['offered'] == 'Y']
                        bottom = df_sheet[df_sheet['offered'] != 'Y']
                        # sort top by score if available
                        if score_col and not top.empty:
                            top = top.sort_values(by=[score_col], ascending=False)
                        df_sheet = pd.concat([top, bottom], ignore_index=True)

                    # column ordering per sheet
                    if sheet_name == 'Offers-List':
                        front = front_offer_list
                    else:
                        front = front_default
                    # build final column list, preserving order and avoiding duplicates
                    # Ensure MaxGATEScore_3yrs appears immediately after Full_Name when both exist
                    if 'Full_Name' in df_sheet.columns and 'MaxGATEScore_3yrs' in df_sheet.columns:
                        # create ordered list that injects MaxGATEScore_3yrs right after Full_Name
                        reordered = []
                        seen = set()
                        for c in (front + all_cols):
                            if c not in df_sheet.columns or c in seen:
                                continue
                            # when we hit Full_Name, append it then MaxGATEScore_3yrs
                            if c == 'Full_Name':
                                reordered.append('Full_Name')
                                seen.add('Full_Name')
                                if 'MaxGATEScore_3yrs' in df_sheet.columns and 'MaxGATEScore_3yrs' not in seen:
                                    reordered.append('MaxGATEScore_3yrs')
                                    seen.add('MaxGATEScore_3yrs')
                            else:
                                if c != 'MaxGATEScore_3yrs':  # skip gate col here — will be added after Full_Name
                                    reordered.append(c)
                                    seen.add(c)
                        # If Full_Name wasn't found in front+all_cols loop, fall back to simple ordering
                        if not reordered:
                            final_cols = [c for c in (front + all_cols) if c in df_sheet.columns]
                        else:
                            final_cols = reordered
                    else:
                        # default behavior when either column missing
                        final_cols = [c for c in (front + all_cols) if c in df_sheet.columns]

                    # safety: if somehow empty, use df_sheet.columns
                    if not final_cols:
                        final_cols = list(df_sheet.columns)

                    # write to excel
                    df_sheet[final_cols].to_excel(writer, sheet_name=sheet_name_safe(sheet_name), index=False)

                except Exception:
                    # continue writing other sheets even if one fails
                    continue

        QMessageBox.information(None, "Download Complete", f"Offers for Round {round_no} exported to:\n{out_filename}")

    except Exception as e:
        QMessageBox.critical(None, "Download Error", f"Could not generate offer report:\n{e}")

    finally:
        conn.close()