import streamlit as st
import pandas as pd
import jellyfish
import re
import io
import os
import tempfile
import subprocess
import shutil
from datetime import datetime
from docx import Document

# ==========================================
# 1. CORE CONFIGURATION
# ==========================================
st.set_page_config(page_title="Temporal Pipeline", layout="wide", initial_sidebar_state="expanded")

ACTION_WORDS = {"OVERHAUL", "RENEW", "REPLACE", "CHANGE", "PULLED OUT", "DISMANTLED", "RECONDITIONED"}
INSPECTION_WORDS = {"CHECK", "INSPECT", "INSP", "CLEAN", "TEST", "MEASURE"}
PMS_LOCKED_COLS = ["Component Name", "Last Overhaul Date", "Total Running Hours"]

# ==========================================
# 2. DUAL-FORMAT INTAKE ENGINE
# ==========================================
def extract_legacy_doc(file_bytes) -> str:
    """Uses Linux OS-level 'antiword' to shatter 1997-2003 binary files."""
    if not shutil.which('antiword'):
        raise Exception("CRITICAL: 'antiword' is not installed on this server. You must add a 'packages.txt' file to your GitHub root containing the word 'antiword'.")

    # Sandbox the binary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.doc') as temp_file:
        temp_file.write(file_bytes)
        temp_path = temp_file.name

    try:
        # Execute binary extraction
        result = subprocess.run(['antiword', temp_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        text = result.stdout.decode('utf-8', errors='ignore')
        
        if result.returncode != 0:
            raise Exception(result.stderr.decode('utf-8', errors='ignore'))
            
        return text
    finally:
        os.remove(temp_path) # Terminate sandbox

def extract_modern_docx(file_bytes) -> str:
    """Extracts modern XML-based Word files."""
    doc = Document(io.BytesIO(file_bytes))
    raw_text = []
    for table in doc.tables:
        for row in table.rows:
            raw_text.append(" ".join([cell.text.strip() for cell in row.cells if cell.text.strip()]))
    for para in doc.paragraphs:
        if para.text.strip(): raw_text.append(para.text.strip())
    return "\n".join(raw_text)

def process_tec19_files(uploaded_files) -> tuple:
    """Routes files to correct extractor and applies State-Machine logic."""
    all_extracted_data = []
    
    for file in uploaded_files:
        try:
            # Route based on extension
            if file.name.lower().endswith('.doc'):
                raw_text = extract_legacy_doc(file.getvalue())
            else:
                raw_text = extract_modern_docx(file.getvalue())
            
            # State-Machine Date Regex Parser
            for line in raw_text.split('\n'):
                line = line.strip()
                dates = re.findall(r'\b\d{2}-\d{2}-\d{2}\b', line)
                
                if dates:
                    job_desc = re.sub(r'\b\d{2}-\d{2}-\d{2}\b', '', line).strip()
                    completed_date = dates[-1] # Chronological Anchor
                    
                    if len(job_desc) > 5:
                        all_extracted_data.append({
                            "Text": job_desc,
                            "Date": completed_date,
                            "Source": file.name
                        })
        except Exception as e:
            return None, f"Fatal parsing error in {file.name}: {str(e)}"
            
    return all_extracted_data, None

def spatial_lock_pms(excel_bytes) -> tuple:
    """Strict Header Enforcement for Excel."""
    try:
        df = pd.read_excel(excel_bytes, engine='openpyxl')
        df.columns = df.columns.str.strip()
        missing = [col for col in PMS_LOCKED_COLS if col not in df.columns]
        if missing: return None, f"Missing exact columns: {missing}"
        return df[PMS_LOCKED_COLS].dropna(subset=["Component Name"]), None
    except Exception as e:
        return None, f"Excel parsing error: {str(e)}"

# ==========================================
# 3. NLP & RECONCILIATION
# ==========================================
def apply_weighted_shield(text: str) -> bool:
    upper_text = str(text).upper()
    if any(action in upper_text for action in ACTION_WORDS): return True
    if any(insp in upper_text for insp in INSPECTION_WORDS): return False
    return False

def generate_phonetic_hash(text: str) -> set:
    clean = re.sub(r'[^A-Z0-9\s]', '', str(text).upper())
    return set(jellyfish.metaphone(word) for word in clean.split() if len(word) > 2)

def parse_date_safely(date_str: str):
    try:
        return datetime.strptime(date_str, "%d-%m-%y")
    except:
        return None

def run_audit(pms_df, tec_data, audit_date):
    """The Core Zero-Trust Engine."""
    results = {"Syncs": [], "Ghosts": [], "Unlogged_Hours": [], "Quarantine": []}
    
    # Pre-hash TEC entries
    filtered_tec = [log for log in tec_data if apply_weighted_shield(log["Text"])]
    for log in filtered_tec:
        log["Hashes"] = generate_phonetic_hash(log["Text"])
        log["ParsedDate"] = parse_date_safely(log["Date"])

    if not filtered_tec:
        results["Quarantine"].append({"System Alert": "No valid action entries survived."})
        return results

    # Scan PMS Ledger
    for _, row in pms_df.iterrows():
        pms_comp = str(row['Component Name'])
        pms_date_str = str(row['Last Overhaul Date'])
        
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
            
            # Semantic Match (40% Phonetic Confidence)
            if hash_score > 0.4: 
                match_found = True
                
                # Math Trap: Physical Hours Allowed
                if log["ParsedDate"]:
                    days_since_overhaul = (audit_date - log["ParsedDate"]).days
                    max_possible_hours = max(days_since_overhaul * 24, 0)
                    
                    if pms_hours > max_possible_hours:
                        results["Unlogged_Hours"].append({
                            "Component": pms_comp,
                            "TEC Date": log["Date"],
                            "PMS Hours": pms_hours,
                            "Max Allowed": max_possible_hours,
                            "Status": "CRITICAL: Counter not reset."
                        })
                        break

                # Temporal Check
                if log["Date"] == pms_date_str:
                    results["Syncs"].append({"Component": pms_comp, "Sync Date": pms_date_str, "TEC Proof": log["Text"]})
                else:
                    results["Quarantine"].append({
                        "Component": pms_comp,
                        "Conflict": f"Date Mismatch. PMS Claims {pms_date_str}, TEC Claims {log['Date']}.",
                        "TEC Proof": log["Text"]
                    })
                break 
                
        if not match_found:
            results["Ghosts"].append({"Component": pms_comp, "PMS Date": pms_date_str, "Status": "No TEC Evidence"})

    return results

# ==========================================
# 4. FRONTEND UI & SANDBOXING
# ==========================================
def main():
    if "audit_results" not in st.session_state: st.session_state.audit_results = None

    with st.sidebar:
        st.title("⚓ Temporal Zero-Trust Pipeline")
        st.divider()
        audit_date = st.date_input("Select Date of Audit (For Physical Limits Math)")
        
        pms_file = st.file_uploader("1. Master PMS Ledger (Excel)", type=['xlsx', 'xls'])
        tec_files = st.file_uploader("2. TEC-19 Logs (Word)", type=['doc', 'docx'], accept_multiple_files=True) 
        
        if st.button("▶ Execute Audit", type="primary", use_container_width=True):
            if not pms_file or not tec_files:
                st.error("🚨 Both datasets are required.")
                st.stop()

            with st.spinner("Cracking binary files and mapping NLP Hashes..."):
                pms_df, pms_error = spatial_lock_pms(pms_file.getvalue())
                tec_data, tec_error = process_tec19_files(tec_files)

                if pms_error: st.error(pms_error); st.stop()
                if tec_error: st.error(tec_error); st.stop()
                if not tec_data: st.error("🚨 Extraction Halted: No chronological data found."); st.stop()

                dt_audit = datetime.combine(audit_date, datetime.min.time())
                st.session_state.audit_results = run_audit(pms_df, tec_data, dt_audit)
                st.success("Audit Complete.")

    st.title("Reconciliation Dashboard")
    
    if st.session_state.audit_results:
        res = st.session_state.audit_results
        t1, t2, t3, t4 = st.tabs(["📊 Overview", "🚨 Running Hours Traps", "👻 Ghost Overhauls", "☣️ Quarantine / Syncs"])

        with t1:
            total_processed = len(res["Syncs"]) + len(res["Ghosts"]) + len(res["Unlogged_Hours"]) + len(res["Quarantine"])
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Components", total_processed)
            m2.metric("Verified Syncs", len(res["Syncs"]))
            m3.metric("Falsified Hours", len(res["Unlogged_Hours"]), delta="High Risk", delta_color="inverse")
            m4.metric("Quarantine Bay", len(res["Quarantine"]), delta="Requires Review", delta_color="inverse")

        with t2:
            st.dataframe(pd.DataFrame(res["Unlogged_Hours"]), use_container_width=True)

        with t3:
            st.dataframe(pd.DataFrame(res["Ghosts"]), use_container_width=True)

        with t4:
            st.write("### ☣️ The Quarantine Bay")
            st.dataframe(pd.DataFrame(res["Quarantine"]), use_container_width=True)
            st.write("### ✅ Verified Syncs")
            st.dataframe(pd.DataFrame(res["Syncs"]), use_container_width=True)
    else:
        st.info("Awaiting Uplink. Ensure packages.txt is deployed to Streamlit.")

if __name__ == "__main__":
    main()
    
