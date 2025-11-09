import sqlite3
import pandas as pd
from PySide6.QtWidgets import QMessageBox

DB_NAME = "mtech_offers.db"

# --- Helper Functions ---

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

# --- Main Logic Functions ---

def upload_round_decisions(round_no, iit_goa_report, other_iit_report, consolidated_report):
    """
    Reads the three decision reports for a given round and saves them to the database.
    
    NOTE: Column names are standardized here for consistency with the DB schema.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    try:
        _create_decision_tables(cursor, round_no)

        # 1. IIT Goa Candidate Decision Report (Requires "MTech Application No", "Applicant Decision")
        df_goa = _read_maybe_df(iit_goa_report)
        df_goa = df_goa[["MTech Application No", "Applicant Decision"]].rename(
            columns={"MTech Application No": "mtech_app_no", "Applicant Decision": "applicant_decision"}
        )
        df_goa.to_sql(f'iit_goa_offers_round{round_no}', conn, if_exists='replace', index=False)
        
        # 2. IIT Goa Offered But Accept and Freeze at Other Institutes (Requires "MTech Application No", "Other Institution Decision")
        df_other = _read_maybe_df(other_iit_report)
        df_other = df_other[["MTech Application No", "Other Institution Decision"]].rename(
            columns={"MTech Application No": "mtech_app_no", "Other Institution Decision": "other_institute_decision"}
        )
        df_other.to_sql(f'accepted_other_institute_round{round_no}', conn, if_exists='replace', index=False)
        
        # 3. Consolidated Accept and Freeze Candidates Across All Institutes (Requires "COAP Reg Id", "Applicant Decision")
        df_consolidated = _read_maybe_df(consolidated_report)
        df_consolidated = df_consolidated[["COAP Reg Id", "Applicant Decision"]].rename(
            columns={"COAP Reg Id": "coap_reg_id", "Applicant Decision": "applicant_decision"}
        )
        # Assuming COAP Reg Id is the COAP number in the candidates table for linking
        df_consolidated.to_sql(f'consolidated_decisions_round{round_no}', conn, if_exists='replace', index=False)

        conn.commit()
        # The message is correct if round_no is passed as the previous round (N-1)
        QMessageBox.information(None, "Success", f"Decisions for Round {round_no} uploaded and saved successfully!")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during decision upload for Round {round_no}:\n{e}")
    finally:
        conn.close()

def _get_eligible_candidates_for_next_round(current_round):
    """
    CORRECTED LOGIC: Determines the COAP IDs eligible for the next round (current_round + 1).
    """
    conn = sqlite3.connect(DB_NAME)
    coaps_out = set()
    
    # 1. Gather ALL 'Accept and Freeze' candidates from all previous rounds (Permanently out)
    for r in range(1, current_round + 1):
        # A. IIT Goa Offer (Accept and Freeze) - FIXED: Uses c.App_no for join
        df_goa_frozen = pd.read_sql_query(f"""
            SELECT c.COAP
            FROM candidates c
            JOIN iit_goa_offers_round{r} d ON d.mtech_app_no = c.App_no
            WHERE d.applicant_decision = 'Accept and Freeze'
        """, conn)
        coaps_out.update(df_goa_frozen['COAP'].tolist())

        # B. Consolidated Decisions (Accept and Freeze at any Institute)
        df_consolidated_frozen = pd.read_sql_query(f"""
            SELECT coap_reg_id 
            FROM consolidated_decisions_round{r}
            WHERE applicant_decision = 'Accept and Freeze'
        """, conn)
        coaps_out.update(df_consolidated_frozen['coap_reg_id'].tolist())

    # 2. Gather 'Reject and Wait' candidates from the LATEST round (current_round)
    # FIXED: Uses c.App_no for join
    # df_rejected_out = pd.read_sql_query(f"""
    #     SELECT c.COAP
    #     FROM candidates c
    #     JOIN iit_goa_offers_round{current_round} d ON d.mtech_app_no = c.App_no
    #     WHERE d.applicant_decision = 'Reject and Wait'
    # """, conn)
    # coaps_out.update(df_rejected_out['COAP'].tolist())
    for r in range(1, current_round + 1):
        df_rejected_out = pd.read_sql_query(f"""
            SELECT c.COAP
            FROM candidates c
            JOIN iit_goa_offers_round{r} d ON d.mtech_app_no = c.App_no
            WHERE d.applicant_decision = 'Reject and Wait'
        """, conn)
        coaps_out.update(df_rejected_out['COAP'].tolist())
    # 3. Fetch all candidates with a GATE score
    df_all_candidates = pd.read_sql_query("""
        SELECT COAP 
        FROM candidates 
        WHERE MaxGATEScore_3yrs IS NOT NULL
    """, conn)
    all_coaps = set(df_all_candidates['COAP'].tolist())

    conn.close()

    # 4. Filter: Eligible for next round = All candidates - Candidates who are out
    eligible_coaps = list(all_coaps - coaps_out)
    
    return eligible_coaps

# --- Confirmed Seat Recalculation (Fixed SQL column 'App_no') ---

def _recalculate_confirmed_seats(last_round, conn):
    """Recalculates the total confirmed seats (Accept and Freeze) up to the specified last_round."""
    if last_round < 1:
        return {}
        
    confirmed_seats = {}
    
    for r in range(1, last_round + 1):
        # Find candidates who 'Accept and Freeze' their IIT Goa offer in this round
        # FIXED: Uses c.App_no for join
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

# --- Get Retained Candidates (New Function for Logic Fix) ---

def _get_retained_candidates(previous_round, conn):
    """
    Fetches COAP IDs and the SEAT category offered to candidates 
    who chose 'Retain and Wait' in the previous round.
    """
    if previous_round < 1:
        return {}
        
    query = f"""
        SELECT o.COAP, o.category
        FROM offers o
        JOIN candidates c ON o.COAP = c.COAP
        JOIN iit_goa_offers_round{previous_round} d ON d.mtech_app_no = c.App_no
        WHERE o.round_no = {previous_round} 
          AND d.applicant_decision = 'Retain and Wait'
    """
    df_retained = pd.read_sql_query(query, conn)
    
    # Return as a dictionary: {COAP: category}
    return df_retained.set_index('COAP')['category'].to_dict()


# --- Main Allocation Logic (Fixed Retain and Wait) ---

def run_round(round_no):
    """
    Perform seat allocation for a given round, respecting prior round decisions and exclusions.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    try:
        eligible_coaps = None
        previous_round = round_no - 1
        
        # 1. Determine eligible COAPs
        if round_no == 1:
            # For Round 1, all candidates with a GATE score are eligible
            cursor.execute("""
                SELECT COAP FROM candidates WHERE MaxGATEScore_3yrs IS NOT NULL
            """)
            eligible_coaps = [row[0] for row in cursor.fetchall()]
        else:
            # For subsequent rounds, filter based on previous round decisions
            eligible_coaps = _get_eligible_candidates_for_next_round(previous_round)
            if not eligible_coaps:
                QMessageBox.warning(None, "Round Complete", f"No eligible candidates remain for Round {round_no}.")
                conn.close()
                return

        # 2. Get list of retained candidates from previous round (for allocation priority)
        retained_coaps_map = _get_retained_candidates(previous_round, conn)
        
        # 3. Fetch all eligible candidates data, sorted by GATE score
        coap_list_str = ', '.join([f"'{c}'" for c in eligible_coaps])
        
        cursor.execute(f"""
            SELECT COAP, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
            FROM candidates
            WHERE COAP IN ({coap_list_str})
            ORDER BY MaxGATEScore_3yrs DESC
        """)
        candidates = cursor.fetchall()
        print(f"Total candidates eligible for Round {round_no}: {len(candidates)}")

        # 4. Recalculate confirmed seats and load seat matrix
        confirmed_seats = _recalculate_confirmed_seats(previous_round, conn)
        seat_matrix = _get_seat_matrix_with_confirmed(conn, confirmed_seats)
        
        print(f"Seat Matrix Loaded (Confirmed Seats from R1 to R{previous_round}):", {k: v['allocated'] for k, v in seat_matrix.items()})

        common_pwd_quota = seat_matrix.get("COMMON_PWD", {"total": 0})["total"]

        # Prepare for offer making
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

        
        # 5. Allocation Loop (Modified to prioritize Retain and Wait)
        
        # --- Helper for checking and recording seat allocation ---
        def try_allocate_seat(coap, name, score, key, status):
            nonlocal seat_matrix
            # Ensure we don't allocate a seat that is already confirmed by someone else
            if key in seat_matrix and seat_matrix[key]["allocated"] < seat_matrix[key]["total"]:
                seat_matrix[key]["allocated"] += 1
                offers_made.append((round_no, coap, name, key, score, status))
                allocated_coaps.add(coap)
                print(f"Allocating {name} to {key} ({status})")
                return True
            return False

        # --- Sub-step 5.1: Allocate Retained Candidates First (FIX: Retention Only) ---
        retained_candidates_data = [c for c in candidates if c[0] in retained_coaps_map]
        
        for coap, name, base_cat, ews, gender, pwd, score in retained_candidates_data:
            
            if coap in allocated_coaps:
                continue # Skip if allocated via a special rule before this loop

            # The candidate already has a category they were offered in the previous round
            retained_category = retained_coaps_map[coap]

            allocated = False
            
            # 1. No Upgrade Check: Directly re-offer the retained seat
            if try_allocate_seat(coap, name, score, retained_category, "Offered (Retained)"):
                allocated = True
            
            # NOTE: If try_allocate_seat returns False, it means the seat was somehow filled
            # by an 'Accept and Freeze' candidate between rounds. This should be extremely rare
            # if the confirmed seat recalculation is accurate, but the logic prevents double-offering.


        # --- Sub-step 5.2: Allocate COMMON_PWD (Remaining candidates only) ---
        temp_common_pwd_quota = common_pwd_quota
        
        # 🔹 Step 1: Allocate COMMON_PWD first
        if temp_common_pwd_quota > 0 and pwd_candidates:
             # Find top PWD candidate not yet allocated (and not retained)
             top_pwd = next((c for c in pwd_candidates if c[0] not in allocated_coaps), None)
             
             if top_pwd:
                coap, name, base_cat, ews, gender, pwd, score = top_pwd

                base_cat = base_cat.strip() if base_cat else "GEN"
                gender = gender.strip().capitalize() if gender else "Male"
                ews = ews.strip().capitalize() if ews else "No"

                seat_key_parts = ["EWS" if ews == "Yes" else base_cat]
                possible_keys = [f"{seat_key_parts[0]}_Female", f"{seat_key_parts[0]}_FandM"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM"]

                # Ensure candidate is not already handled as retained
                if coap not in allocated_coaps:
                    for key in possible_keys:
                        if try_allocate_seat(coap, name, score, key, "Offered (Common PWD)"):
                            temp_common_pwd_quota -= 1
                            break
        
        # 🔹 Step 2: Allocate remaining PWD candidates
        for coap, name, base_cat, ews, gender, pwd, score in pwd_candidates:
            if coap in allocated_coaps:
                continue

            base_cat = base_cat.strip() if base_cat else "GEN"
            gender = gender.strip().capitalize() if gender else "Male"
            ews = ews.strip().capitalize() if ews else "No"
            seat_key_parts = ["EWS" if ews=="Yes" else base_cat]

            possible_keys = [f"{seat_key_parts[0]}_Female_PWD", f"{seat_key_parts[0]}_FandM_PWD"] if gender=="Female" else [f"{seat_key_parts[0]}_FandM_PWD"]

            for key in possible_keys:
                if try_allocate_seat(coap, name, score, key, "Offered (PWD)"):
                    break
            else:
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

            for key in possible_keys:
                if try_allocate_seat(coap, name, score, key, "Offered"):
                    break
            else:
                print(f"No seat available for {name}")
        # 🔹 Step 4: Save results
        cursor.executemany("""
            INSERT OR REPLACE INTO offers (round_no, COAP, Full_Name, category, MaxGATEScore_3yrs, offer_status)
            VALUES (?, ?, ?, ?, ?, ?)
        """, offers_made)
        
        conn.commit()
        QMessageBox.information(None, "Success", f"Round {round_no} allocation complete!\nTotal offers: {len(offers_made)}")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during round {round_no} allocation:\n{e}")
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
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_offers.to_excel(writer, sheet_name='Offers_Summary', index=False)
        df_detailed.to_excel(writer, sheet_name='Offers_Detailed', index=False)

    QMessageBox.information(None, "Download Complete", f"Offers saved as {filename}")

# Note: The original `run_round_1` is replaced by the generic `run_round(1)`
# to allow for a unified, multi-round process.

