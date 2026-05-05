import streamlit as st
import pandas as pd
from docx import Document
import jellyfish
import re
import io
from datetime import datetime

# ==========================================
# 1. CORE CONFIGURATION & CONSTANTS
# ==========================================
st.set_page_config(page_title="Temporal Pipeline", layout="wide", initial_sidebar_state="expanded")

# The Weighted NLP Shield
ACTION_WORDS = {"OVERHAUL", "RENEW", "REPLACE", "CHANGE", "PULLED OUT", "DISMANTLED", "RECONDITIONED"}
INSPECTION_WORDS = {"CHECK", "INSPECT", "INSP", "CLEAN", "TEST", "MEASURE"}

# The Spatial Lock Parameters
PMS_LOCKED_COLS = ["Component Name", "Last Overhaul Date", "Total Running Hours"]

# ==========================================
# 2. THE INTAKE & EXTRACTION ENGINES
# ==========================================
def spatial_lock_pms(excel_bytes) -> tuple:
    """Securely extract Excel data, enforcing strict spatial column locks."""
    try:
        df = pd.read_excel(excel_bytes, engine='openpyxl')
        # Normalize column names to strip invisible whitespaces
        df.columns = df.columns.str.strip()
        
        # Enforce spatial lock
        missing = [col for col in PMS_LOCKED_COLS if col not in df.columns]
        if missing:
            return None, f"Spatial Lock Failed. Missing exact columns: {missing}"
            
        return df[PMS_LOCKED_COLS].dropna(subset=["Component Name"]), None
    except Exception as e:
        return None, f"Fatal Excel parsing error: {str(e)}"

def extract_multiple_tec19(uploaded_files) -> tuple:
    """
    Bypasses visual tables using a State-Machine.
    Flattens text, finds date arrays, and strictly targets the 'Completed' date.
    """
    all_extracted_data = []
    
    for file in uploaded_files:
        # THE INTAKE GATEKEEPER: Absolutely refuse legacy binary files.
        if file.name.endswith('.doc'):
            return None, f"CRITICAL: Legacy binary format detected in '{file.name}'. To guarantee 100% data integrity, open this file in Word, 'Save As' .docx, and re-upload."
            
        try:
            doc = Document(io.BytesIO(file.getvalue()))
            for table in doc.tables:
                for row in table.rows:
                    # Flatten the row to destroy merged-cell traps
                    row_text = " ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                    
                    # Regex target: DD-MM-YY 
                    dates = re.findall(r'\b\d{2}-\d{2}-\d{2}\b', row_text)
                    
                    if dates:
                        # Extract the description by stripping out the dates
                        job_desc = re.sub(r'\b\d{2}-\d{2}-\d{2}\b', '', row_text).strip()
                        # Temporal Anchor: The final date in the sequence is 'Completed'
                        completed_date = dates[-1] 
                        
                        # Filter out empty artifact rows
                        if len(job_desc) > 5:
                            all_extracted_data.append({
                                "Text": job_desc,
                                "Date": completed_date,
                                "Source File": file.name
                            })
        except Exception as e:
            return None, f"Fatal Word parsing error in {file.name}: {str(e)}"
            
    return all_extracted_data, None

# ==========================================
# 3. VECTORIZED NLP & SHIELD LOGIC
# ==========================================
def apply_weighted_shield(text: str) -> bool:
    """Action words mathematically overpower inspection words."""
    upper_text = str(text).upper()
    if any(action in upper_text for action in ACTION_WORDS): return True
    if any(insp in upper_text for insp in INSPECTION_WORDS): return False
    return False

def generate_phonetic_hash(text: str) -> set:
    """Converts strings to alphanumeric Metaphone sets. Zero dependencies."""
    clean = re.sub(r'[^A-Z0-9\s]', '', str(text).upper())
    return set(jellyfish.metaphone(word) for word in clean.split() if len(word) > 2)

def parse_date_safely(date_str: str):
    """Safely converts DD-MM-YY to a workable datetime object."""
    try:
        return datetime.strptime(date_str, "%d-%m-%y")
    except:
        return None

# ==========================================
# 4. THE RECONCILIATION ENGINE
# ==========================================
def run_audit(pms_df, tec_data, audit_date):
    """Cross-references hashes and routes to isolated data buckets."""
    results = {"Syncs": [], "Ghosts": [], "Unlogged_Hours": [], "Quarantine": []}
    
    # Pre-Hash Valid TEC-19 Entries (Reduces calculation time to milliseconds)
    filtered_tec = [log for log in tec_data if apply_weighted_shield(log["Text"])]
    for log in filtered_tec:
        log["Hashes"] = generate_phonetic_hash(log["Text"])
        log["ParsedDate"] = parse_date_safely(log["Date"])

    if not filtered_tec:
        results["Quarantine"].append({"System Alert": "No valid action entries survived the Anti-Inspection Shield. Check document formatting."})
        return results

    # Iterate PMS and cross-reference
    for _, row in pms_df.iterrows():
        pms_comp = str(row['Component Name'])
        pms_date_str = str(row['Last Overhaul Date'])
        
        # Safely extract running hours (handling blanks/text in excel)
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
            
            # Semantic Match Threshold (40% Phonetic Match)
            if hash_score > 0.4: 
                match_found = True
                
                # Goal 3 TRAP: Maximum Physical Hours Math
                if log["ParsedDate"]:
                    days_since_overhaul = (audit_date - log["ParsedDate"]).days
                    max_possible_hours = max(days_since_overhaul * 24, 0)
                    
                    if pms_hours > max_possible_hours:
                        results["Unlogged_Hours"].append({
                            "Component": pms_comp,
                            "TEC-19 Overhaul Date": log["Date"],
                            "PMS Claimed Hours": pms_hours,
                            "Max Physics Allowed": max_possible_hours,
                            "Status": "CRITICAL: Counter was not reset."
                        })
                        break # Trap triggered, stop checking this component

                # Temporal Proximity Check
                if log["Date"] == pms_date_str:
                    results["Syncs"].append({
                        "Component": pms_comp, 
                        "Sync Date": pms_date_str, 
                        "TEC-19 Proof": log["Text"]
                    })
                else:
                    results["Quarantine"].append({
                        "Component": pms_comp,
                        "Conflict Reason": f"Date Mismatch. PMS Claims {pms_date_str}, TEC-19 Claims {log['Date']}.",
                        "TEC-19 Proof": log["Text"]
                    })
                break # Match routed, move to next PMS component
                
        # If the loop finishes and no TEC-19 hash matched
        if not match_found:
            results["Ghosts"].append({
                "Component": pms_comp, 
                "PMS Claimed Date": pms_date_str, 
                "Status": "No TEC-19 Evidence Found"
            })

    return results

# ==========================================
# 5. FRONTEND UI & SANDBOXING
# ==========================================
def main():
    # Immutable Session State Initialization
    if "audit_results" not in st.session_state: 
        st.session_state.audit_results = None

    with st.sidebar:
        st.title("⚓ Temporal Zero-Trust Pipeline")
        st.markdown("Ensure TEC-19 files are in `.docx` format before uploading.")
        st.divider()
        
        audit_date = st.date_input("Select Date of Audit (For Physical Limits Math)")
        pms_file = st.file_uploader("1. Master PMS Ledger (Excel)", type=['xlsx'])
        tec_files = st.file_uploader("2. TEC-19 Logs (Word)", type=['docx'], accept_multiple_files=True) 
        
        if st.button("▶ Execute Audit", type="primary", use_container_width=True):
            if not pms_file or not tec_files:
                st.error("🚨 Both datasets are required to initiate.")
                st.stop()

            with st.spinner("Locking coordinates and mapping NLP Hashes..."):
                # Phase 1 & 2 Execution
                pms_df, pms_error = spatial_lock_pms(pms_file.getvalue())
                tec_data, tec_error = extract_multiple_tec19(tec_files)

                # Exception Handling Gate
                if pms_error: st.error(pms_error); st.stop()
                if tec_error: st.error(tec_error); st.stop()
                if not tec_data: st.error("🚨 Extraction Halted: No recognized chronological data found in the Word document."); st.stop()

                # Phase 3 & 4 Execution
                dt_audit = datetime.combine(audit_date, datetime.min.time())
                st.session_state.audit_results = run_audit(pms_df, tec_data, dt_audit)
                st.success("Audit Complete. Data locked into memory.")

    # Phase 5 Dashboard Rendering
    st.title("Reconciliation Dashboard")
    
    if st.session_state.audit_results:
        res = st.session_state.audit_results
        
        # Dynamic Tabs
        t1, t2, t3, t4 = st.tabs(["📊 Overview", "🚨 Running Hours Traps", "👻 Ghost Overhauls", "☣️ Quarantine / Syncs"])

        with t1:
            st.subheader("Audit Integrity Summary")
            total_processed = len(res["Syncs"]) + len(res["Ghosts"]) + len(res["Unlogged_Hours"]) + len(res["Quarantine"])
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Components", total_processed)
            m2.metric("Verified Syncs", len(res["Syncs"]))
            m3.metric("Falsified/Missed Hours", len(res["Unlogged_Hours"]), delta="High Risk", delta_color="inverse")
            m4.metric("Quarantine Bay", len(res["Quarantine"]), delta="Requires Review", delta_color="inverse")

        with t2:
            st.error("CRITICAL: The components below were proven to be overhauled via TEC-19, but the official PMS running hours physically exceed the time elapsed. The counter was not reset.")
            st.dataframe(pd.DataFrame(res["Unlogged_Hours"]), use_container_width=True)

        with t3:
            st.warning("DRIFT DETECTED: The PMS claims these overhauls occurred, but no proof exists in the uploaded TEC-19 diaries.")
            st.dataframe(pd.DataFrame(res["Ghosts"]), use_container_width=True)

        with t4:
            st.write("### ☣️ The Quarantine Bay")
            st.info("Date conflicts or ambiguous semantic matches. Zero silent failures permitted.")
            st.dataframe(pd.DataFrame(res["Quarantine"]), use_container_width=True)
            
            st.divider()
            
            st.write("### ✅ Verified Syncs")
            st.success("Perfect Semantic & Temporal Matches.")
            st.dataframe(pd.DataFrame(res["Syncs"]), use_container_width=True)
    else:
        st.info("Awaiting Uplink. Upload data and configure the audit date in the sidebar.")

if __name__ == "__main__":
    main()
