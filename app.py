import streamlit as st
import pandas as pd
from docx import Document
import jellyfish
import re
import io
from datetime import datetime

# ==========================================
# CONSTANTS & CONFIGURATION
# ==========================================
ACTION_WORDS = {"OVERHAUL", "RENEW", "REPLACE", "CHANGE", "PULLED OUT", "DISMANTLED", "RECONDITIONED"}
INSPECTION_WORDS = {"CHECK", "INSPECT", "INSP", "CLEAN", "TEST", "MEASURE"}
PMS_LOCKED_COLS = ["Component Name", "Last Overhaul Date", "Total Running Hours"]

st.set_page_config(page_title="Temporal Pipeline", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# EXTRACTION LAYER (Supports Multiple Word Docs)
# ==========================================
def spatial_lock_pms(excel_bytes):
    try:
        df = pd.read_excel(excel_bytes, engine='openpyxl')
        df.columns = df.columns.str.strip()
        missing = [col for col in PMS_LOCKED_COLS if col not in df.columns]
        if missing: return None, f"Missing columns: {missing}"
        return df[PMS_LOCKED_COLS].dropna(subset=["Component Name"]), None
    except Exception as e:
        return None, f"Excel parsing error: {str(e)}"

def extract_multiple_tec19(uploaded_files):
    """Processes an unlimited array of uploaded Word documents."""
    all_extracted_data = []
    
    for file in uploaded_files:
        if file.name.endswith('.doc'):
            return None, f"Legacy format detected in {file.name}. Please Save As .docx."
            
        try:
            doc = Document(io.BytesIO(file.getvalue()))
            for table in doc.tables:
                for row in table.rows:
                    row_text = " ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                    dates = re.findall(r'\b\d{2}-\d{2}-\d{2}\b', row_text)
                    
                    if dates:
                        job_desc = re.sub(r'\b\d{2}-\d{2}-\d{2}\b', '', row_text).strip()
                        completed_date = dates[-1] # Target final date
                        if len(job_desc) > 5:
                            all_extracted_data.append({"Text": job_desc, "Date": completed_date, "Source": file.name})
        except Exception as e:
            return None, f"Error parsing {file.name}: {str(e)}"
            
    return all_extracted_data, None

# ==========================================
# NLP & MATH LOGIC
# ==========================================
def apply_weighted_shield(text: str) -> bool:
    upper_text = text.upper()
    if any(action in upper_text for action in ACTION_WORDS): return True
    if any(insp in upper_text for insp in INSPECTION_WORDS): return False
    return False

def generate_phonetic_hash(text: str) -> set:
    clean = re.sub(r'[^A-Z0-9\s]', '', str(text).upper())
    return set(jellyfish.metaphone(word) for word in clean.split() if len(word) > 2)

def parse_date(date_str):
    """Safely converts DD-MM-YY to a workable datetime object."""
    try:
        return datetime.strptime(date_str, "%d-%m-%y")
    except:
        return None

def run_audit(pms_df, tec_data, audit_date):
    results = {"Syncs": [], "Ghosts": [], "Unlogged_Hours": [], "Quarantine": []}
    
    # Pre-Hash Valid TEC-19 Entries
    filtered_tec = [log for log in tec_data if apply_weighted_shield(log["Text"])]
    for log in filtered_tec:
        log["Hashes"] = generate_phonetic_hash(log["Text"])
        log["ParsedDate"] = parse_date(log["Date"])

    for _, row in pms_df.iterrows():
        pms_comp = str(row['Component Name'])
        pms_date_str = str(row['Last Overhaul Date'])
        
        # Safe extraction of running hours
        try:
            pms_hours = float(row['Total Running Hours'])
        except:
            pms_hours = 0.0

        pms_hashes = generate_phonetic_hash(pms_comp)
        if not pms_hashes: continue
            
        match_found = False
        
        for log in filtered_tec:
            intersection = pms_hashes.intersection(log["Hashes"])
            if not pms_hashes: continue
            hash_score = len(intersection) / len(pms_hashes)
            
            # If Semantic Match is Strong (Over 40% Token Intersection)
            if hash_score > 0.4: 
                match_found = True
                
                # MATHEMATICAL TRAP: Check Running Hours
                if log["ParsedDate"]:
                    days_since_overhaul = (audit_date - log["ParsedDate"]).days
                    max_possible_hours = max(days_since_overhaul * 24, 0)
                    
                    if pms_hours > max_possible_hours:
                        results["Unlogged_Hours"].append({
                            "Component": pms_comp,
                            "TEC Date": log["Date"],
                            "PMS Hours Claim": pms_hours,
                            "Max Physics Allowed": max_possible_hours,
                            "Status": "CRITICAL: Hours not reset after overhaul."
                        })
                        break # Trap triggered, route to Unlogged

                # Temporal Check
                if log["Date"] == pms_date_str:
                    results["Syncs"].append({"Component": pms_comp, "Date": pms_date_str, "Proof": log["Text"]})
                else:
                    results["Quarantine"].append({
                        "Component": pms_comp,
                        "Conflict": f"Date Mismatch. PMS: {pms_date_str} vs TEC: {log['Date']}.",
                    })
                break
                
        if not match_found:
            results["Ghosts"].append({"Component": pms_comp, "Claimed Date": pms_date_str})

    return results

# ==========================================
# FRONTEND UI
# ==========================================
def main():
    if "audit_results" not in st.session_state: st.session_state.audit_results = None

    with st.sidebar:
        st.title("⚓ Temporal Zero-Trust Pipeline")
        st.divider()
        audit_date = st.date_input("Select Date of Audit (For Hour Math)")
        pms_file = st.file_uploader("1. Master PMS (Excel)", type=['xlsx'])
        
        # Upgraded to accept multiple files
        tec_files = st.file_uploader("2. TEC-19 Logs (Word)", type=['docx'], accept_multiple_files=True) 
        
        if st.button("▶ Run Audit Engine", type="primary"):
            if not pms_file or not tec_files:
                st.error("Both datasets required.")
                st.stop()

            with st.spinner("Locking coordinates and calculating temporal limits..."):
                pms_df, pms_error = spatial_lock_pms(pms_file.getvalue())
                tec_data, tec_error = extract_multiple_tec19(tec_files)

                if pms_error: st.error(pms_error); st.stop()
                if tec_error: st.error(tec_error); st.stop()

                # Execute Math
                st.session_state.audit_results = run_audit(pms_df, tec_data, datetime.combine(audit_date, datetime.min.time()))
                st.success("Audit Complete.")

    st.title("Reconciliation Dashboard")
    
    if st.session_state.audit_results:
        res = st.session_state.audit_results
        t1, t2, t3, t4 = st.tabs(["📊 Overview", "🚨 Overhaul Hour Traps", "👻 Ghost Overhauls", "☣️ Quarantine / Syncs"])

        with t1:
            st.metric("Total Components Analyzed", len(res["Syncs"]) + len(res["Ghosts"]) + len(res["Unlogged_Hours"]) + len(res["Quarantine"]))
            
        with t2:
            st.error("Components were overhauled, but the PMS hours physically exceed the time elapsed. Counter was not reset.")
            st.dataframe(pd.DataFrame(res["Unlogged_Hours"]), use_container_width=True)

        with t3:
            st.warning("Expected with partial data. PMS claims overhaul, but no TEC-19 proof uploaded yet.")
            st.dataframe(pd.DataFrame(res["Ghosts"]), use_container_width=True)

        with t4:
            st.write("**Quarantine Bay (Conflicts):**")
            st.dataframe(pd.DataFrame(res["Quarantine"]), use_container_width=True)
            st.write("**Verified Syncs:**")
            st.dataframe(pd.DataFrame(res["Syncs"]), use_container_width=True)
    else:
        st.info("Awaiting Uplink.")

if __name__ == "__main__":
    main()
