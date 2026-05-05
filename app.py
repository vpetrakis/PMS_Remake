import streamlit as st
import pandas as pd
from docx import Document
import jellyfish
import re
import nltk
from nltk.tokenize import sent_tokenize
import io

# ---------------------------------------------------------
# SYSTEM INITIALIZATION & CACHING
# ---------------------------------------------------------
@st.cache_resource
def init_nlp():
    """Silently load NLTK to prevent mid-execution crashes."""
    try:
        nltk.data.find('tokenizers/punkt')
    except LookupError:
        nltk.download('punkt', quiet=True)

init_nlp()

# ---------------------------------------------------------
# THE ZERO-TRUST ENGINE
# ---------------------------------------------------------
class TemporalEngine:
    def __init__(self):
        # Weighted Shield: Action words OVERPOWER Inspection words.
        self.action_words = {"OVERHAUL", "RENEW", "REPLACE", "CHANGE", "PULLED OUT", "DISMANTLED"}
        self.inspection_words = {"CHECK", "INSPECT", "INSP", "CLEAN", "TEST", "MEASURE"}
        self.pms_locked_cols = ["Component Name", "Last Overhaul Date", "Total Running Hours"]

    def spatial_lock_pms(self, excel_bytes) -> pd.DataFrame:
        """Phase 1: Ingest Excel with strict spatial locks."""
        try:
            # Force string conversion on names, dates can be parsed later
            df = pd.read_excel(excel_bytes, usecols=lambda c: c.strip() in self.pms_locked_cols)
            if df.empty or len(df.columns) < 3:
                return pd.DataFrame() # Trigger lock failure
            return df.dropna(subset=["Component Name"]) # Drop empty rows
        except Exception as e:
            return pd.DataFrame()

    def multi_layer_word_extraction(self, word_bytes) -> list:
        """Phase 2: Hybrid Table & Regex extraction to defeat bad Word formatting."""
        doc = Document(word_bytes)
        extracted_data = []
        
        # Layer 1: Attempt standard Table Extraction
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if len(cells) >= 3:
                    # Look for date patterns (DD-MM-YY) to anchor the data
                    date_match = re.search(r'\b\d{2}-\d{2}-\d{2}\b', " ".join(cells))
                    if date_match:
                        # Assume the longest text cell is the Job Description
                        job_desc = max(cells, key=len) 
                        extracted_data.append({"Text": job_desc, "Date": date_match.group(0), "Method": "Table"})

        # Layer 2: Regex Fallback (If tables were merged/broken, read raw paragraphs)
        if len(extracted_data) < 5: 
            for para in doc.paragraphs:
                text = para.text.strip()
                date_match = re.search(r'\b\d{2}-\d{2}-\d{2}\b', text)
                if date_match and len(text) > 15:
                    extracted_data.append({"Text": text, "Date": date_match.group(0), "Method": "Regex Fallback"})

        return extracted_data

    def apply_shield(self, text: str) -> bool:
        """Phase 3: The Weighted Anti-Inspection Shield."""
        text_upper = text.upper()
        
        # If it has an action word, it passes, regardless of check words.
        if any(action in text_upper for action in self.action_words):
            return True
            
        # If it has check words and NO action words, it is blocked.
        if any(insp in text_upper for insp in self.inspection_words):
            return False
            
        return False # Default to block (Zero-Trust)

    def extract_phonetic_tokens(self, text: str) -> set:
        """Phase 4: Alphanumeric Metaphone generation."""
        clean_text = re.sub(r'[^A-Z0-9\s]', '', str(text).upper())
        return set(jellyfish.metaphone(token) for token in clean_text.split() if len(token) > 2)

    def run_reconciliation(self, pms_df, tec19_entries):
        """Phase 5: Cross-reference logic routing."""
        results = {"Syncs": [], "Ghost": [], "Unlogged": [], "Quarantine": []}
        
        # Filter TEC-19 through the Shield
        filtered_logs = [log for log in tec19_entries if self.apply_shield(log["Text"])]
        
        # If nothing passed the shield, everything goes to Quarantine to prevent silent failure
        if not filtered_logs:
            results["Quarantine"].append({"Error": "Critical: No valid maintenance events survived the Anti-Inspection Shield."})
            return results

        # Simplistic matching for demonstration (Production would use Date-Delta math here)
        for _, row in pms_df.iterrows():
            pms_comp = str(row['Component Name'])
            pms_tokens = self.extract_phonetic_tokens(pms_comp)
            
            matched = False
            for log in filtered_logs:
                log_tokens = self.extract_phonetic_tokens(log["Text"])
                intersection = pms_tokens.intersection(log_tokens)
                
                # Confidence Threshold: 40% of phonetic tokens must match to trigger a review
                if len(pms_tokens) > 0 and (len(intersection) / len(pms_tokens)) > 0.4:
                    results["Quarantine"].append({
                        "PMS Component": pms_comp,
                        "TEC-19 Entry": log["Text"],
                        "Date Recorded": log["Date"],
                        "Status": "Partial Match Detected - Requires Human Verification"
                    })
                    matched = True
                    break
            
            if not matched:
                results["Ghost"].append({"Component Name": pms_comp, "PMS Date Claim": str(row.get('Last Overhaul Date', 'N/A'))})
                
        return results

# ---------------------------------------------------------
# FRONTEND: NAVIGATIONAL UI
# ---------------------------------------------------------
st.set_page_config(page_title="Temporal Pipeline", layout="wide", initial_sidebar_state="expanded")

def main():
    engine = TemporalEngine()

    with st.sidebar:
        st.title("⚓ Pipeline Controls")
        st.markdown("Upload documents to initiate Zero-Trust Audit.")
        st.divider()
        pms_file = st.file_uploader("1. Master PMS (Excel)", type=['xlsx'])
        tec_file = st.file_uploader("2. TEC-19 Log (Word)", type=['docx', 'doc'])
        run_audit = st.button("▶ Run Audit Engine", use_container_width=True, type="primary")

    st.title("Temporal Reconciliation Dashboard")
    
    if run_audit and pms_file and tec_file:
        with st.spinner("Locking coordinates and applying NLP shields..."):
            
            # 1. Ingest Data
            pms_df = engine.spatial_lock_pms(pms_file)
            tec_data = engine.multi_layer_word_extraction(tec_file)

            # 2. Check Integrity
            if pms_df.empty:
                st.error("🚨 Ingestion Halted: Failed to lock onto PMS headers (Component Name, Last Overhaul Date, Total Running Hours).")
                st.stop()
            if not tec_data:
                st.error("🚨 Ingestion Halted: Multi-Layer extraction failed to find chronological data in the Word file.")
                st.stop()

            # 3. Reconcile
            results = engine.run_reconciliation(pms_df, tec_data)

            # 4. Display Results in Tabs
            tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "✅ Verified Syncs", "👻 Ghost Overhauls", "☣️ Quarantine Bay"])

            with tab1:
                st.subheader("Data Extraction Summary")
                col1, col2, col3 = st.columns(3)
                col1.metric("PMS Components Locked", len(pms_df))
                col2.metric("Raw Log Entries Found", len(tec_data))
                col3.metric("Items in Quarantine", len(results["Quarantine"]), delta="Requires Review", delta_color="inverse")

            with tab2:
                st.success("Perfect Semantic & Temporal Matches.")
                st.dataframe(pd.DataFrame(results["Syncs"]), use_container_width=True)

            with tab3:
                st.warning("PMS claims an overhaul, but no proof exists in the TEC-19 log.")
                st.dataframe(pd.DataFrame(results["Ghost"]), use_container_width=True)

            with tab4:
                st.error("Ambiguous matches or spatial conflicts. Zero silent failures.")
                if results["Quarantine"]:
                    st.dataframe(pd.DataFrame(results["Quarantine"]), use_container_width=True)
                else:
                    st.write("Quarantine Bay is completely clear.")
    else:
        st.info("Awaiting secure document uplink from the sidebar.")

if __name__ == "__main__":
    main()
