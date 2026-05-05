import streamlit as st
import pandas as pd
import jellyfish
import os
import tempfile
import subprocess
import shutil
import json
from datetime import datetime
from docx import Document
import google.generativeai as genai

# ==========================================
# 1. CORE CONFIGURATION
# ==========================================
st.set_page_config(page_title="Temporal Pipeline | AI Edition", layout="wide")

PMS_LOCKED_COLS = ["Component Name", "Last Overhaul Date", "Total Running Hours"]

# ==========================================
# 2. FILE EXTRACTION (Getting the raw text)
# ==========================================
def extract_legacy_doc(file_bytes) -> str:
    """Uses OS-level 'antiword' to shatter 1997-2003 binary files into raw text."""
    if not shutil.which('antiword'):
        raise Exception("CRITICAL: 'antiword' is not installed. Check packages.txt.")

    with tempfile.NamedTemporaryFile(delete=False, suffix='.doc') as temp_file:
        temp_file.write(file_bytes)
        temp_path = temp_file.name

    try:
        result = subprocess.run(['antiword', temp_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        text = result.stdout.decode('utf-8', errors='ignore')
        if result.returncode != 0:
            raise Exception(result.stderr.decode('utf-8', errors='ignore'))
        return text
    finally:
        os.remove(temp_path)

def extract_modern_docx(file_bytes) -> str:
    """Extracts text from modern XML Word files."""
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])

# ==========================================
# 3. THE LLM STRUCTURING ENGINE (Pathway B)
# ==========================================
def parse_text_with_llm(raw_text: str, api_key: str) -> list:
    """Uses AI to intelligently extract overhauls and format them as JSON."""
    genai.configure(api_key=api_key)
    # Using Gemini 1.5 Flash for high-speed, high-accuracy text extraction
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = f"""
    You are an expert maritime auditor. Analyze the following daily logbook text.
    Your objective is to find ONLY major maintenance events (OVERHAUL, RENEW, REPLACE, CHANGE).
    Strictly IGNORE routine checks (CHECK, INSPECT, CLEAN, TEST, MEASURE).
    
    Output the data as a strict, valid JSON array of objects with exactly two keys: "Component" and "Date".
    Format the Date as "DD-MM-YY" (find the 'Completed' date if multiple exist).
    If no major events are found, return an empty array [].
    
    Logbook Text:
    {raw_text}
    """
    
    try:
        response = model.generate_content(prompt)
        # Strip markdown formatting to parse JSON safely
        json_string = response.text.replace("```json", "").replace("
```", "").strip()
        return json.loads(json_string)
    except Exception as e:
        st.error(f"LLM Extraction Failed: {str(e)}")
        return []

def spatial_lock_pms(excel_bytes) -> tuple:
    try:
        df = pd.read_excel(excel_bytes, engine='openpyxl')
        df.columns = df.columns.str.strip()
        missing = [col for col in PMS_LOCKED_COLS if col not in df.columns]
        if missing: return None, f"Missing columns: {missing}"
        return df[PMS_LOCKED_COLS].dropna(subset=["Component Name"]), None
    except Exception as e:
        return None, f"Excel error: {str(e)}"

# ==========================================
# 4. VECTORIZED NLP & ZERO-TRUST MATH
# ==========================================
def generate_phonetic_hash(text: str) -> set:
    clean = ''.join(e for e in str(text).upper() if e.isalnum() or e.isspace())
    return set(jellyfish.metaphone(word) for word in clean.split() if len(word) > 2)

def parse_date_safely(date_str: str):
    try:
        # Handles various 2-digit year formats
        return datetime.strptime(date_str, "%d-%m-%y")
    except:
        return None

def run_audit(pms_df, tec_json_data, audit_date):
    """Cross-references the AI's JSON against the Excel Math."""
    results = {"Syncs": [], "Ghosts": [], "Unlogged_Hours": [], "Quarantine": []}
    
    # Pre-hash TEC entries
    for log in tec_json_data:
        log["Hashes"] = generate_phonetic_hash(log.get("Component", ""))
        log["ParsedDate"] = parse_date_safely(log.get("Date", ""))

    if not tec_json_data:
        results["Quarantine"].append({"System Alert": "The AI found no major overhauls in the uploaded logbooks."})
        return results

    # Scan PMS Ledger
    for _, row in pms_df.iterrows():
        pms_comp = str(row['Component Name'])
        pms_date_str = str(row['Last Overhaul Date'])
        
        try: pms_hours = float(row['Total Running Hours'])
        except: pms_hours = 0.0

        pms_hashes = generate_phonetic_hash(pms_comp)
        if not pms_hashes: continue
            
        match_found = False
        
        for log in tec_json_data:
            intersection = pms_hashes.intersection(log["Hashes"])
            if not pms_hashes: continue
            hash_score = len(intersection) / len(pms_hashes)
            
            # Semantic Match (40% Phonetic Confidence)
            if hash_score > 0.4: 
                match_found = True
                
                # Goal 3 Math Trap: Physical Hours Allowed
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
                    results["Syncs"].append({"Component": pms_comp, "Sync Date": pms_date_str, "TEC Proof": log["Component"]})
                else:
                    results["Quarantine"].append({
                        "Component": pms_comp,
                        "Conflict": f"Date Mismatch. PMS Claims {pms_date_str}, TEC Claims {log['Date']}.",
                        "TEC Proof": log["Component"]
                    })
                break 
                
        if not match_found:
            results["Ghosts"].append({"Component": pms_comp, "PMS Date": pms_date_str, "Status": "No TEC Evidence"})

    return results

# ==========================================
# 5. FRONTEND UI
# ==========================================
def main():
    if "audit_results" not in st.session_state: st.session_state.audit_results = None

    with st.sidebar:
        st.title("⚓ AI Temporal Pipeline")
        api_key = st.text_input("Enter Gemini API Key", type="password", help="Required to parse unstructured .doc files.")
        st.divider()
        audit_date = st.date_input("Select Date of Audit (For Physical Limits Math)")
        
        pms_file = st.file_uploader("1. Master PMS Ledger (Excel)", type=['xlsx', 'xls'])
        tec_files = st.file_uploader("2. TEC-19 Logs (Word)", type=['doc', 'docx'], accept_multiple_files=True) 
        
        if st.button("▶ Execute AI Audit", type="primary", use_container_width=True):
            if not api_key:
                st.error("🚨 API Key required for AI text structuring.")
                st.stop()
            if not pms_file or not tec_files:
                st.error("🚨 Both datasets are required.")
                st.stop()

            with st.spinner("Shattering binaries, structuring text with AI, and auditing timelines..."):
                # Phase 1: Spatial Lock
                pms_df, pms_error = spatial_lock_pms(pms_file.getvalue())
                if pms_error: st.error(pms_error); st.stop()

                # Phase 2: Extract & Structure via LLM
                all_tec_json = []
                for file in tec_files:
                    if file.name.lower().endswith('.doc'):
                        raw_text = extract_legacy_doc(file.getvalue())
                    else:
                        raw_text = extract_modern_docx(file.getvalue())
                    
                    structured_data = parse_text_with_llm(raw_text, api_key)
                    all_tec_json.extend(structured_data)

                if not all_tec_json: 
                    st.error("🚨 Extraction Halted: The AI could not find any major overhauls in the provided documents.")
                    st.stop()

                # Phase 3 & 4: Execute Zero-Trust Math
                dt_audit = datetime.combine(audit_date, datetime.min.time())
                st.session_state.audit_results = run_audit(pms_df, all_tec_json, dt_audit)
                st.success("Audit Complete.")

    # Dashboard Rendering
    st.title("Reconciliation Dashboard")
    
    if st.session_state.audit_results:
        res = st.session_state.audit_results
        t1, t2, t3, t4 = st.tabs(["📊 Overview", "🚨 Running Hours Traps", "👻 Ghost Overhauls", "☣️ Quarantine / Syncs"])

        with t1:
            total_processed = len(res["Syncs"]) + len(res["Ghosts"]) + len(res["Unlogged_Hours"]) + len(res["Quarantine"])
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Components Analyzed", total_processed)
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
        st.info("Awaiting Uplink. Enter your API key and upload files to begin.")

if __name__ == "__main__":
    main()
