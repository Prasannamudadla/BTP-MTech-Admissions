# ui/update_dialog.py
import sqlite3
from pathlib import Path
from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QGridLayout, QVBoxLayout, QLabel, QPushButton, QScrollArea
)

# Map UI labels -> DB columns (None => calculated field)
FIELD_MAP = {
    "FullName":              "Full_Name",
    "ApplicationNumber":     "App_no",
    "COAP":                  "COAP",
    "Email":                 "Email",
    "MaxGateScore":          "MaxGATEScore_3yrs",
    "Gender":                "Gender",
    "Category":              "Category",
    "EWS":                   "Ews",
    "PWD":                   "Pwd",
    "SSCper":                "SSC_per",
    "HSSCper":               "HSSC_per",
    "DegreeCGPA8thSem":      "Degree_CGPA_8th",
    
    # Offer/Decision Fields (Calculated/Pulled from offers/decision tables)
    "Offered":               None, 
    "Accepted":              None, # Y, N, R, E
    "OfferCat":              None,
    "isOfferPwd":            None, # Y/N
    "OfferedRound":          None,
    "RetainRound":           None,
    "RejectOrAcceptRound":   None,
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
            
            # For fields mapped to DB columns in 'candidates' table
            if colname:
                v = data.get(colname)
                value = "NULL" if (v is None or v == "") else str(v)
            # For calculated/offer fields (where colname is None)
            elif label in data:
                 v = data.get(label)
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
        """Fetch entire row for this COAP ID, plus the latest offer/decision status."""
        try:
            conn = self._connect()
            cursor = conn.cursor()

            # 1. Get candidate's core data
            cursor.execute("""
                SELECT * FROM candidates
                WHERE COAP = ?
                LIMIT 1
            """, (self.coap_id,))
            candidate_row = cursor.fetchone()
            
            if not candidate_row:
                conn.close()
                return {}
            
            data = dict(candidate_row)
            app_no = data["App_no"] 
            
            # Initialize offer/decision fields to default NULL
            data["Offered"] = "NULL"
            data["Accepted"] = "NULL"
            data["OfferCat"] = "NULL"
            data["isOfferPwd"] = "NULL"
            data["OfferedRound"] = "NULL"
            data["RetainRound"] = "NULL"
            data["RejectOrAcceptRound"] = "NULL"


            # 2. Get the latest offer details (OfferedRound, OfferCat, isOfferPwd)
            cursor.execute("""
                SELECT round_no, category, offer_status
                FROM offers
                WHERE COAP = ?
                ORDER BY round_no DESC
            """, (self.coap_id,))
            all_offers = cursor.fetchall()

            max_round_with_offer = max(offer["round_no"] for offer in all_offers) if all_offers else 0
            
            if all_offers:
                latest_offer_row = dict(all_offers[0])
                
                # Logic to find the ORIGINAL (non-retained) offer round
                original_offer_round = latest_offer_row["round_no"] # Default to latest entry
                
                # Find the highest round that is NOT clearly a retained status
                for offer in all_offers:
                    offer_status = offer["offer_status"]
                    if "RETAINED" not in offer_status.upper():
                        original_offer_round = offer["round_no"]
                        latest_offer_row = dict(offer)
                        break # Found the original offer round
                
                data["Offered"] = "Y"
                data["OfferCat"] = latest_offer_row["category"]
                data["OfferedRound"] = original_offer_round
                
                # Determine isOfferPwd
                offer_status = latest_offer_row["offer_status"]
                is_pwd_offer = False
                if "_PWD" in data["OfferCat"].upper() or "PWD" in offer_status.upper():
                    is_pwd_offer = True
                    
                data["isOfferPwd"] = "Yes" if is_pwd_offer else "No"


            # 3. Determine the latest decision made by the candidate (Y, N, R) and the Retain Round
            
            latest_decision_found = False
            
            if max_round_with_offer > 0:
                # Loop backwards from the highest round to find the latest decision
                for r in range(max_round_with_offer, 0, -1):
                    decision_table = f"iit_goa_offers_round{r}"
                    
                    cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{decision_table}'")
                    if cursor.fetchone():
                        cursor.execute(f"""
                            SELECT applicant_decision
                            FROM {decision_table}
                            WHERE mtech_app_no = ?
                            LIMIT 1
                        """, (app_no,))
                        decision_row = cursor.fetchone()

                        if decision_row:
                            decision = decision_row["applicant_decision"]
                            
                            # A. Record the ABSOLUTE LATEST FINAL decision (Accept/Reject)
                            if not latest_decision_found:
                                if decision == 'Accept and Freeze':
                                    data["Accepted"] = "Y"
                                    data["RejectOrAcceptRound"] = r
                                    latest_decision_found = True
                                    
                                elif decision == 'Reject and Wait':
                                    data["Accepted"] = "N"
                                    data["RejectOrAcceptRound"] = r
                                    latest_decision_found = True
                            
                            # B. Record the LATEST RETAIN decision (regardless of later final decision)
                            if decision == 'Retain and Wait' and data["RetainRound"] == "NULL":
                                # This is the first (latest) retain decision we encounter in the backward loop
                                data["RetainRound"] = r
                                # If the latest decision (A) was not Y/N, then R is the current status
                                if not latest_decision_found:
                                    data["Accepted"] = "R"
                                
                                # If we found both the latest final decision AND the RetainRound, we can break.
                                if latest_decision_found:
                                    break
            
            # If we went through all rounds and found no decision, accepted remains "NULL".
            # If a Retain was found but no final decision, Accepted is "R".
            

            # 4. Check for 'Accepted elsewhere' status (E) - This takes absolute precedence
            if max_round_with_offer > 0:
                for r in range(1, max_round_with_offer + 1): 
                    consolidated_table = f"consolidated_decisions_round{r}"
                    cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{consolidated_table}'")
                    if cursor.fetchone():
                        # CORRECTED SYNTAX: using consolidated_table variable instead of f-string brace error
                        cursor.execute(f"""
                            SELECT coap_reg_id
                            FROM {consolidated_table} 
                            WHERE coap_reg_id = ? AND applicant_decision = 'Accept and Freeze'
                            LIMIT 1
                        """, (self.coap_id,))
                        if cursor.fetchone():
                            # Candidate accepted elsewhere (E)
                            data["Accepted"] = "E"
                            data["RejectOrAcceptRound"] = r 
                            data["RetainRound"] = "NULL" 
                            data["Offered"] = "N"
                            data["OfferCat"] = "NULL"
                            data["OfferedRound"] = "NULL"
                            data["isOfferPwd"] = "NULL"
                            break 
            
            conn.close()
            return data
        except Exception as e:
            print(f"Error loading record for {self.coap_id}: {e}")
            return {}
    # def _load_record(self) -> dict:
    #     """Fetch entire row for this COAP ID, plus the latest offer/decision status."""
    #     try:
    #         conn = self._connect()
    #         cursor = conn.cursor()

    #         # 1. Get candidate's core data
    #         cursor.execute("""
    #             SELECT * FROM candidates
    #             WHERE COAP = ?
    #             LIMIT 1
    #         """, (self.coap_id,))
    #         candidate_row = cursor.fetchone()
            
    #         if not candidate_row:
    #             conn.close()
    #             return {}
            
    #         data = dict(candidate_row)
    #         app_no = data["App_no"] 
            
    #         # Initialize offer/decision fields to default NULL
    #         data["Offered"] = "NULL"
    #         data["Accepted"] = "NULL"
    #         data["OfferCat"] = "NULL"
    #         data["isOfferPwd"] = "NULL"
    #         data["OfferedRound"] = "NULL"
    #         data["RetainRound"] = "NULL"
    #         data["RejectOrAcceptRound"] = "NULL"


    #         # 2. Get the latest offer details (round, category, status)
    #         # Find the ORIGINAL offer round (the highest round that is NOT 'Retained')
    #         cursor.execute("""
    #             SELECT round_no, category, offer_status
    #             FROM offers
    #             WHERE COAP = ?
    #             ORDER BY round_no DESC
    #         """, (self.coap_id,))
    #         all_offers = cursor.fetchall()

    #         latest_offer_row = None
    #         original_offer_round = None
            
    #         # Identify the original offer round and the very latest entry
    #         if all_offers:
    #             latest_offer_row = dict(all_offers[0]) # The latest entry in the offers table
                
    #             # Iterate through offers to find the highest round that is NOT (Retained) or (Upgrade)
    #             # This logic is tricky, let's simplify: the OfferedRound should be the round of the first *valid* offer,
    #             # unless they were upgraded. For simplicity, we stick to the round listed if it's the *initial* offer.
                
    #             # Check for the highest round where the status is just 'Offered' or 'Offered (Common PWD)'
    #             for offer in all_offers:
    #                 offer_status = offer["offer_status"]
    #                 if offer_status in ('Offered', 'Offered (Common PWD)', 'Offered (PWD)'):
    #                     original_offer_round = offer["round_no"]
    #                     latest_offer_row = dict(offer)
    #                     break
                
    #             # Fallback: if all offers are (Retained) or (Upgrade), use the latest entry
    #             if original_offer_round is None:
    #                 original_offer_round = latest_offer_row["round_no"]

                
    #             data["Offered"] = "Y"
    #             data["OfferCat"] = latest_offer_row["category"]
                
    #             # CRITICAL FIX: Use the round identified by the simplified logic
    #             # This ensures COAP22011204 shows Round 2, if Round 3 was Retained/Upgraded.
    #             data["OfferedRound"] = original_offer_round
                
    #             # Determine isOfferPwd using the latest offer's category/status
    #             offer_status = latest_offer_row["offer_status"]
    #             is_pwd_offer = False
    #             if "_PWD" in data["OfferCat"].upper() or "PWD" in offer_status.upper():
    #                 is_pwd_offer = True
                    
    #             data["isOfferPwd"] = "Y" if is_pwd_offer else "N"


    #         # 3. Determine the latest decision made by the candidate (Y, N, R)
    #         max_round_with_offer = max(offer["round_no"] for offer in all_offers) if all_offers else 0
            
    #         if max_round_with_offer > 0:
    #             # Loop backwards from the highest round to find the latest decision
    #             for r in range(max_round_with_offer, 0, -1):
    #                 decision_table = f"iit_goa_offers_round{r}"
                    
    #                 cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{decision_table}'")
    #                 if cursor.fetchone():
    #                     cursor.execute(f"""
    #                         SELECT applicant_decision
    #                         FROM {decision_table}
    #                         WHERE mtech_app_no = ?
    #                         LIMIT 1
    #                     """, (app_no,))
    #                     decision_row = cursor.fetchone()

    #                     if decision_row:
    #                         decision = decision_row["applicant_decision"]
                            
    #                         if decision == 'Accept and Freeze':
    #                             data["Accepted"] = "Y"
    #                             data["RejectOrAcceptRound"] = r
    #                             # Stop searching, this is the final decision
    #                             break
    #                         elif decision == 'Reject and Wait':
    #                             data["Accepted"] = "N"
    #                             data["RejectOrAcceptRound"] = r
    #                             # Stop searching, this is the final decision
    #                             break
                            
    #                         # Decision for 'Retain and Wait' (Must be the latest Retain)
    #                         elif decision == 'Retain and Wait':
    #                             data["Accepted"] = "R"
    #                             # This is the round where they decided to retain
    #                             data["RetainRound"] = r
    #                             # Continue to check earlier rounds in case they accepted/rejected/froze later
    #                             # But if we reach here, it means in this round 'r', they made a retain decision
    #                             # which implies any final accept/reject must be in a later round (r+1, r+2, etc.)
    #                             # Since we are iterating backwards, if we found a Retain, we should break IF 
    #                             # no subsequent (Accept/Reject) decision was found. Since Accept/Reject breaks above,
    #                             # finding Retain here means it's the latest non-final decision. 
    #                             # We must reset Accept/Reject fields here if they were previously set by a later round.
    #                             data["RejectOrAcceptRound"] = "NULL" # Reset if a later one wasn't found
                                
    #                             # We break here because any prior decisions are superseded by this latest Retain
    #                             break 


    #         # 4. Check for 'Accepted elsewhere' status (E) - This takes absolute precedence
    #         if max_round_with_offer > 0:
    #             for r in range(1, max_round_with_offer + 1): 
    #                 consolidated_table = f"consolidated_decisions_round{r}"
    #                 cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{consolidated_table}'")
    #                 if cursor.fetchone():
    #                     cursor.execute(f"""
    #                         SELECT coap_reg_id
    #                         FROM {consolidated_table}
    #                         WHERE coap_reg_id = ? AND applicant_decision = 'Accept and Freeze'
    #                         LIMIT 1
    #                     """, (self.coap_id,))
    #                     if cursor.fetchone():
    #                         # Candidate accepted elsewhere (E)
    #                         data["Accepted"] = "E"
    #                         data["RejectOrAcceptRound"] = r 
    #                         data["RetainRound"] = "NULL" 
    #                         data["Offered"] = "N" # Candidate is out
    #                         data["OfferCat"] = "NULL"
    #                         data["OfferedRound"] = "NULL"
    #                         data["isOfferPwd"] = "NULL"
    #                         break 
            
    #         conn.close()
    #         return data
    #     except Exception as e:
    #         print(f"Error loading record for {self.coap_id}: {e}")
    #         return {}

    def _make_card(self, label: str, value: str) -> QWidget:
        """
        Creates a card-like UI element with:
        - label on top
        - value displayed inside a rounded rectangle
        """

        card = QWidget()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # Title label
        title = QLabel(label)
        title.setAlignment(Qt.AlignLeft)
        title.setStyleSheet("""
            font-size: 13px;
            font-weight: 600;
            color: #333;
        """)

        # Value label
        display_value = value if value not in (None, "", "NULL") else "NULL"
        color = "#111" if display_value != "NULL" else "#999"

        val = QLabel(display_value)
        val.setAlignment(Qt.AlignLeft)
        val.setWordWrap(True)
        val.setStyleSheet(f"""
            background: #fafafa;
            border: 1px solid #ddd;
            border-radius: 6px;
            padding: 8px;
            font-size: 12px;
            color: {color};
        """)

        layout.addWidget(title)
        layout.addWidget(val)

        return card

    