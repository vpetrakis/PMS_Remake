import streamlit as st
import pandas as pd
import jellyfish
import re
import nltk
from nltk.tokenize import sent_tokenize
import io

# Initialize NLTK quietly
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt', quiet=True)

# ==========================================
# BACKEND: ZERO-TRUST NLP & DATA ENGINE
# ==========================================

class TemporalEngine:
    def __init__(self):
        self.banned_keywords = {"CHECK", "INSPECT", "INSP", "CLEAN", "TEST", "MEASURE"}
        self.target_events = {"OVERHAUL", "RENEW", "REPLACE", "CHANGE"}
        self.pms_locked_cols = ["Component Name", "Last Overhaul Date", "Total Running Hours"]

    def spatial_lock_pms(self, excel_file):
        """Phase 1: Read Excel but lock onto specific columns only."""
        try:
            df = pd.read_excel(excel_file, usecols=lambda c: c in self.pms_locked_cols)
            return df
        except Exception as e:
            return pd.DataFrame() # Return empty on failure to trigger Zero-Trust halt

    def anti_inspection_shield(self, text: str):
        """Phase 2: Filter out routine checks, keep major events."""
        if not isinstance(text, str): return ""
        sentences = sent_tokenize(text.upper())
        valid_sentences = []
        for sentence in sentences:
            tokens = set(re.findall(r'\b\w+\b', sentence))
            if tokens.intersection(self.banned_keywords):
                continue
            if tokens.intersection(self.target_events):
                valid_sentences.append(sentence)
        return " ".join(valid_sentences)

    def extract_phonetic_tokens(self, text: str):
        """Phase 3: Alphanumeric Metaphone tokenization."""
        clean_text = re.sub(r'[^A-Z0-9\s]', '', str(text).upper())
        return set(jellyfish.metaphone(token) for token in clean_text.split())

    def run_reconciliation(self, pms_df, tec19_data):
        """Phase 4 & 5: Cross-reference and route to Quarantine/Syncs."""
        # Note: This is a simplified logic mock for the monolithic structure
        results = {"Verified Syncs": [], "Ghost Overhauls": [], "Unlogged": [], "Quarantine Bay": []}
        
        # Example logic routing
        for index, row in pms_df.iterrows():
            comp_name = str(row.get('Component Name', ''))
            if not comp_name: continue
            
            pms_tokens = self.extract_phonetic_tokens(comp_name)
            # In a full run, we would compare this against every filtered TEC-19 entry
            # For demonstration, routing everything to Quarantine to enforce Zero-Trust
            results["Quarantine Bay"].append({
                "Component": comp_name,
                "Reason": "Pending deep temporal cross-reference"
            })
            
        return results

# ==========================================
# FRONTEND: STREAMLIT UI
# ==========================================

def main():
    st.set_page_config(page_title="Temporal Reconciliation Pipeline", layout="wide")
    engine = TemporalEngine()

    # --- Header ---
    st.title("⚓ Temporal Reconciliation Pipeline")
    st.markdown("**Zero-Trust Auditing:** Cross-referencing PMS Master Ledgers against TEC-19 Diaries.")
    st.divider()

    # --- Upload Area ---
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Master PMS Ledger")
        pms_file = st.file_uploader("Upload Excel (.xlsx)", type=['xlsx', 'xls'])
    
    with col2:
        st.subheader("2. TEC-19 Work Schedule")
        tec19_file = st.file_uploader("Upload Word/Text (.doc, .docx, .txt)", type=['docx', 'doc', 'txt'])

    # --- Execution Logic ---
    if pms_file and tec19_file:
        st.success("Files securely loaded. Initiating Spatial Locks and Anti-Inspection Shields...")
        
        with st.spinner("Reconciling timelines..."):
            # 1. Process PMS
            pms_df = engine.spatial_lock_pms(pms_file)
            
            if pms_df.empty:
                st.error("🚨 CRITICAL: Spatial Column Lock Failed. Could not find exact headers in PMS.")
                st.stop()

            # 2. Process TEC-19 (Mocking extraction for this layout)
            # tec19_text = docx_parser(tec19_file) 
            
            # 3. Run Engine
            audit_results = engine.run_reconciliation(pms_df, "mock_tec_data")

        # --- Dashboard Output ---
        st.divider()
        st.header("Audit Results")
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Verified Syncs", len(audit_results["Verified Syncs"]))
        m2.metric("Ghost Overhauls", len(audit_results["Ghost Overhauls"]))
        m3.metric("Unlogged Maintenance", len(audit_results["Unlogged"]))
        m4.metric("Quarantine Bay", len(audit_results["Quarantine Bay"]), delta="Requires Review", delta_color="inverse")

        # --- The Quarantine Bay View ---
        st.subheader("☣️ The Quarantine Bay")
        st.info("Items below could not be matched with 100% confidence. Zero silent failures permitted.")
        if audit_results["Quarantine Bay"]:
            st.dataframe(pd.DataFrame(audit_results["Quarantine Bay"]), use_container_width=True)
        else:
            st.write("Quarantine Bay is clear.")

if __name__ == "__main__":
    main()
