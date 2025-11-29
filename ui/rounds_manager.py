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
def download_offers(round_no):
    """
    Generates a comprehensive report for the given round (N).
    The report includes:
    1. New offers/upgrades made in the current round (N).
    2. Seats held by candidates who 'Accepted and Freeze' in all previous rounds (< N).
    """
    conn = sqlite3.connect(DB_NAME)
    
    # --- 1. Dynamically Build the Query to Find ALL Frozen Candidates ---
    
    # This list will hold the SELECT statements for each previous round's 'Accept and Freeze' status
    frozen_unions = []
    
    # Iterate from Round 1 up to (but not including) the current round_no
    for prev_round in range(1, round_no):
        iit_goa_table = f"iit_goa_offers_round{prev_round}"
        
        # 1. Select the COAP from the decision table where status is 'Accept and Freeze'
        # 2. Add the round number where they froze the offer
        frozen_unions.append(f"""
            SELECT T1."COAP", '{prev_round}' AS accepted_round
            FROM (SELECT DISTINCT T2.COAP FROM offers T2 WHERE T2.round_no = {prev_round}) T1 -- COAP in offers table
            INNER JOIN {iit_goa_table} T3 
                ON T3.App_no = T1.COAP -- Assuming the decision table uses App_no/COAP for join
            WHERE T3.Status = 'Accept and Freeze'
        """)

    # Combine the individual round SELECT statements using UNION ALL
    all_frozen_coaps_cte = " UNION ALL ".join(frozen_unions)
    
    # --- 2. Construct the Final Comprehensive SQL Query ---
    
    sql_query = f"""
        WITH AllFrozenCandidates AS (
            -- This CTE combines all COAP IDs that have chosen 'Accept and Freeze' in rounds 1 to N-1
            {all_frozen_coaps_cte}
        ),
        
        LastFrozenOffer AS (
            -- This CTE finds the *specific* offer details (category, score) for the frozen candidate
            SELECT
                T1.COAP,
                MAX(T1.round_no) AS max_round
            FROM
                offers T1
                INNER JOIN AllFrozenCandidates T2 ON T1.COAP = T2.COAP
            GROUP BY
                T1.COAP
        )
        
        -- PART 1: Previously Accepted/Frozen Candidates
        SELECT
            T3.COAP,
            T3.Full_Name,
            T3.MaxGATEScore_3yrs AS GATEScore,
            T3.category AS category_offered, 
            '{round_no}' AS round_no_report, -- Use current round number for display
            'OFFER HELD (ACCEPTED)' AS offer_status
        FROM
            offers T3
            INNER JOIN LastFrozenOffer LFO ON 
                T3.COAP = LFO.COAP AND T3.round_no = LFO.max_round
            
        UNION ALL
        
        -- PART 2: New Offers/Upgrades for the Current Round (round_no)
        SELECT
            COAP,
            Full_Name,
            MaxGATEScore_3yrs AS GATEScore,
            category AS category_offered,
            round_no AS round_no_report,
            offer_status
        FROM
            offers
        WHERE
            round_no = {round_no}
            
        ORDER BY
            GATEScore DESC;
    """
    
    # ----------------------------------------------------

    try:
        # 3. Execute the query and load into a DataFrame
        df_report = pd.read_sql_query(sql_query, conn)

        # 4. Save the DataFrame to an Excel file
        file_name = f"Round_{round_no}_Offers_Report.xlsx"
        df_report.to_excel(file_name, index=False)
        
        QMessageBox.information(None, "Download Complete", f"Offers for Round {round_no} downloaded to:\n{file_name}")

    except Exception as e:
        QMessageBox.critical(None, "Download Error", f"Could not generate offer report:\n{e}")

    finally:
        conn.close()
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
def download_offers(round_no):
    """
    Generates a comprehensive report for the given round (N).
    The report includes:
    1. New offers/upgrades made in the current round (N).
    2. Seats held by candidates who 'Accept and Freeze' in all previous rounds (< N).
    """
    conn = sqlite3.connect(DB_NAME)
    
    # --- 1. Dynamically Build the Query to Find ALL Frozen Candidates ---
    
    frozen_unions = []
    
    # Iterate from Round 1 up to (but not including) the current round_no
    for prev_round in range(1, round_no):
        iit_goa_table = f"iit_goa_offers_round{prev_round}"
        
        # We target the 'iit_goa_offers_roundX' table for the 'Accept and Freeze' status
        frozen_unions.append(f"""
            SELECT T1.COAP, '{prev_round}' AS accepted_round
            FROM (SELECT DISTINCT T2.COAP FROM offers T2 WHERE T2.round_no = {prev_round}) T1
            INNER JOIN {iit_goa_table} T3 
                -- Use the column names identified from your PRAGMA query
                ON T3.mtech_app_no = T1.COAP 
            WHERE 
                T3.applicant_decision = 'Accept and Freeze'
        """)

    # Handle the case where round_no is 1 (no previous rounds)
    if not frozen_unions:
        all_frozen_coaps_cte = "SELECT NULL AS COAP, NULL AS accepted_round WHERE 1 = 0" 
    else:
        all_frozen_coaps_cte = " UNION ALL ".join(frozen_unions)
    
    # --- 2. Construct the Final Comprehensive SQL Query ---
    
    sql_query = f"""
        WITH AllFrozenCandidates AS (
            {all_frozen_coaps_cte}
        ),
        
        LastFrozenOffer AS (
            -- Find the specific offer details (category, score) for the most recent frozen offer
            SELECT
                T1.COAP,
                MAX(T1.round_no) AS max_round
            FROM
                offers T1
                INNER JOIN AllFrozenCandidates T2 ON T1.COAP = T2.COAP
            GROUP BY
                T1.COAP
        )
        
        -- PART 1: Previously Accepted/Frozen Candidates (The seats that are held)
        SELECT
            T3.round_no AS round_no,
            T3.COAP,
            T3.Full_Name,
            T3.category,
            T3.MaxGATEScore_3yrs AS GATEScore,
            -- Display status as 'OFFER HELD (ACCEPTED)' for reporting clarity
            'OFFER HELD (ACCEPTED)' AS offer_status
        FROM
            offers T3
            INNER JOIN LastFrozenOffer LFO ON 
                T3.COAP = LFO.COAP AND T3.round_no = LFO.max_round
            
        UNION ALL
        
        -- PART 2: New Offers/Upgrades for the Current Round (round_no)
        SELECT
            round_no,
            COAP,
            Full_Name,
            category,
            MaxGATEScore_3yrs AS GATEScore,
            offer_status
        FROM
            offers
        WHERE
            round_no = {round_no}
            
        ORDER BY
            GATEScore DESC;
    """
    
    # ----------------------------------------------------

    try:
        # 3. Execute the query and load into a DataFrame
        df_report = pd.read_sql_query(sql_query, conn)
        
        # 4. Rename columns to match your desired Excel output format
        df_report.rename(columns={
            'round_no': 'round_no',  # This now contains the original offer round or current round
            'GATEScore': 'GATEScore',
            'category': 'category',
            'offer_status': 'offer_status'
        }, inplace=True)

        # 5. Save the DataFrame to an Excel file
        file_name = f"Round_{round_no}_Offers_Report.xlsx"
        df_report.to_excel(file_name, index=False)
        
        QMessageBox.information(None, "Download Complete", f"Offers for Round {round_no} downloaded to:\n{file_name}")

    except Exception as e:
        QMessageBox.critical(None, "Download Error", f"Could not generate offer report:\nExecution failed on sql\n{e}")

    finally:
        conn.close()