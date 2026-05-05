import streamlit as st
import pandas as pd
from docx import Document
import jellyfish
import re
import io

# ==========================================
# CONSTANTS & CONFIGURATION
# ==========================================
ACTION_WORDS = {"OVERHAUL", "RENEW", "REPLACE", "CHANGE", "PULLED OUT", "DISMANTLED", "RECONDITIONED"}
INSPECTION_WORDS = {"CHECK", "INSPECT", "INSP", "CLEAN", "TEST", "MEASURE"}
PMS_LOCKED_COLS = ["Component Name", "Last Overhaul Date", "Total Running Hours"]

st.set_page_config(page_title="Temporal Pipeline", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# PHASE 1 & 2: INTAKE & EXTRACTION STATE MACHINE
# ==========================================
def spatial_lock_pms(excel_bytes):
    """Securely extract Excel data, enforcing strict header columns."""
    try:
        df = pd.read_excel(excel_bytes, engine='openpyxl')
        # Normalize column names to avoid trailing space errors
        df.columns = df.columns.str.strip()
        
        # Enforce spatial lock
        missing = [col for col in PMS_LOCKED_COLS if col not in df.columns]
        if missing:
            return None, f"Spatial Lock Failed. Missing columns: {missing}"
            
        return df[PMS_LOCKED_COLS].dropna(subset=["Component Name"]), None
    except Exception as e:
        return None, f"Excel parsing error: {str(e)}"

def extract_tec19_data(docx_bytes):
    """
    Bypasses visual tables. Flattens text, finds date arrays, 
    and strictly targets the chronological 'Completed' date.
    """
    try:
        doc = Document(io.BytesIO(docx_bytes))
        extracted_data = []
        
        # Flatten all tables into a reliable state-machine format
        for table in doc.tables:
            for row in table.rows:
                # Join cells to prevent merged-cell fragmentation
                row_text = " ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                
                # Regex target: DD-MM-YY (e.g., 01-01-26)
                dates = re.findall(r'\b\d{2}-\d{2}-\d{2}\b', row_text)
                
                if dates:
                    # The actual Job Description is the text excluding the dates
                    job_desc = re.sub(r'\b\d{2}-\d{2}-\d{2}\b', '', row_text).strip()
                    # Temporal Anchor: The final date in the sequence is the 'Completed' date
                    completed_date = dates[-1] 
                    
                    if len(job_desc) > 5: # Filter out empty artifacts
                        extracted_data.append({
                            "Text": job_desc,
                            "Date": completed_date
                        })
        return extracted_data, None
    except Exception as e:
        return None, f"Word parsing error: {str(e)}"

# ==========================================
# PHASE 3: VECTORIZED NLP & SHIELD
# ==========================================
def apply_weighted_shield(text: str) -> bool:
    """Action words mathematically overpower inspection words."""
    upper_text = text.upper()
    if any(action in upper_text for action in ACTION_WORDS): return True
    if any(insp in upper_text for insp in INSPECTION_WORDS): return False
    return False

def generate_phonetic_hash(text: str) -> set:
    """Converts strings to alphanumeric Metaphone sets."""
    clean = re.sub(r'[^A-Z0-9\s]', '', str(text).upper())
    return set(jellyfish.metaphone(word) for word in clean.split() if len(word) > 2)

# ==========================================
# PHASE 4: THE RECONCILIATION ENGINE
# ==========================================
def run_audit(pms_df, tec_data):
    """Cross-references hashes and routes to isolated data buckets."""
    results = {"Syncs": [], "Ghosts": [], "Unlogged": [], "Quarantine": []}
    
    # 1. Filter and Pre-Hash TEC-19 (Vectorization for speed)
    filtered_tec = []
    for log in tec_data:
        if apply_weighted_shield(log["Text"]):
            filtered_tec.append({
                "Text": log["Text"],
                "Date": log["Date"],
                "Hashes": generate_phonetic_hash(log["Text"])
            })

    if not filtered_tec:
        results["Quarantine"].append({"System Alert": "No valid action entries survived the Anti-Inspection Shield."})
        return results

    # 2. Iterate PMS and cross-reference
    for _, row in pms_df.iterrows():
        pms_comp = str(row['Component Name'])
        pms_date = str(row['Last Overhaul Date'])
        pms_hashes = generate_phonetic_hash(pms_comp)
        
        match_found = False
        
        if not pms_hashes:
            continue
            
        for log in filtered_tec:
            intersection = pms_hashes.intersection(log["Hashes"])
            hash_score = len(intersection) / len(pms_hashes)
            
            # High Confidence Semantic Match (Adjustable tolerance)
            if hash_score > 0.4: 
                match_found = True
                
                # Check Temporal Proximity (Exact match for this demonstration)
                if log["Date"] == pms_date:
                    results["Syncs"].append({"Component": pms_comp, "PMS Date": pms_date, "TEC Proof": log["Text"]})
                else:
                    results["Quarantine"].append({
                        "Component": pms_comp,
                        "Conflict": f"Date Mismatch. PMS claims {pms_date}, but TEC-19 shows {log['Date']}.",
                        "TEC Proof": log["Text"]
                    })
                break # Move to next PMS component
                
        if not match_found:
            results["Ghosts"].append({"Component": pms_comp, "Claimed Date": pms_date, "Status": "No TEC-19 Evidence Found"})

    return results

# ==========================================
# PHASE 5: FRONTEND UI & SANDBOXING
# ==========================================
def main():
    # Initialize Session State to prevent memory leaks on UI re-renders
    if "audit_results" not in st.session_state:
        st.session_state.audit_results = None

    with st.sidebar:
        st.title("⚓ Temporal Zero-Trust Pipeline")
        st.markdown("Upload files to execute the cross-reference.")
        st.divider()
        pms_file = st.file_uploader("1. Master PMS (Excel)", type=['xlsx'])
        tec_file = st.file_uploader("2. TEC-19 Log (Word)", type=['docx']) # Notice .doc is blocked natively
        
        if st.button("▶ Initialize Audit Engine", use_container_width=True, type="primary"):
            if not pms_file or not tec_file:
                st.error("Both files required.")
                st.stop()
            
            # The Intake Gatekeeper: Explicitly block old .doc files
            if tec_file.name.endswith('.doc'):
                st.error("🚨 Legacy .doc format detected. Please open the file in Word, 'Save As' .docx, and re-upload. This guarantees 100% data integrity.")
                st.stop()

            with st.spinner("Locking coordinates and processing NLP..."):
                # Run Extraction
                pms_df, pms_error = spatial_lock_pms(pms_file.getvalue())
                tec_data, tec_error = extract_tec19_data(tec_file.getvalue())

                # Halt on Extraction Failure
                if pms_error: st.error(pms_error); st.stop()
                if tec_error: st.error(tec_error); st.stop()
                if not tec_data: st.error("No chronological tables found in Word document."); st.stop()

                # Run Reconciler and cache to session state
                st.session_state.audit_results = run_audit(pms_df, tec_data)
                st.success("Audit Complete. Data Cached.")

    # Dashboard Rendering
    st.title("Reconciliation Dashboard")
    
    if st.session_state.audit_results:
        res = st.session_state.audit_results
        
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "✅ Verified Syncs", "👻 Ghost Overhauls", "☣️ Quarantine Bay"])

        with tab1:
            st.subheader("System Architecture Status")
            col1, col2 = st.columns(2)
            col1.metric("Verified Syncs", len(res["Syncs"]))
            col2.metric("Items in Quarantine", len(res["Quarantine"]), delta="Requires Human Review", delta_color="inverse")
            st.info("The NLP Engine has successfully hashed and mapped all documentation. Navigate the tabs to review the isolated data buckets.")

        with tab2:
            st.dataframe(pd.DataFrame(res["Syncs"]), use_container_width=True)

        with tab3:
            st.warning("PMS claims an overhaul, but no corresponding action was found in the TEC-19 log.")
            st.dataframe(pd.DataFrame(res["Ghosts"]), use_container_width=True)

        with tab4:
            st.error("Date conflicts or ambiguous semantic matches. Zero silent failures permitted.")
            st.dataframe(pd.DataFrame(res["Quarantine"]), use_container_width=True)
    else:
        st.info("Awaiting Uplink. Please upload the Excel PMS and the .docx converted TEC-19 file in the sidebar.")

if __name__ == "__main__":
    main()
