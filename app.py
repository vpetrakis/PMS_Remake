import io
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime

import jellyfish
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="Temporal Pipeline Corrected", layout="wide")

ACTION_WORDS = {
    "OVERHAUL", "RENEW", "REPLACE", "CHANGE", "PULLED OUT", "DISMANTLED",
    "RECONDITIONED", "PULLED", "OPENED", "MOUNTED BACK", "REPLACED"
}
INSPECTION_WORDS = {"CHECK", "INSPECT", "INSP", "CLEAN", "TEST", "MEASURE", "VISUAL INSPECTION"}
PMS_LOCKED_COLS = ["Component Name", "Last Overhaul Date", "Total Running Hours"]
MONTHS = 'JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER'


def normalize_text(text: str) -> str:
    text = str(text or "").upper()
    replacements = {
        "M.E.": "MAIN ENGINE",
        "ME ": "MAIN ENGINE ",
        "DG ": "DIESEL GENERATOR ",
        "L.O.": "LUBE OIL",
        "F.O.": "FUEL OIL",
        "EXH.": "EXHAUST",
        "COMP.": "COMPRESSOR",
        "AIR COND.": "AIR CONDITION",
        "AIR CONDITIONED": "AIR CONDITION",
        "TC": "TURBOCHARGER",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r'[^A-Z0-9\s\-/\.]+', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def parse_date_safely(date_str: str):
    raw = str(date_str or '').strip().upper().replace('.', '')
    patterns = [
        "%d-%m-%y", "%d/%m/%y", "%d-%m-%Y", "%d/%m/%Y",
        "%d %b %y", "%d %B %y", "%d %b %Y", "%d %B %Y"
    ]
    for fmt in patterns:
        try:
            return datetime.strptime(raw, fmt)
        except Exception:
            pass
    return None


def extract_legacy_doc(file_bytes) -> tuple[str, str | None]:
    antiword = shutil.which('antiword')
    if not antiword:
        return "", "antiword not installed in this environment; convert .doc to .docx or install antiword via OS packages"
    with tempfile.NamedTemporaryFile(delete=False, suffix='.doc') as temp_file:
        temp_file.write(file_bytes)
        temp_path = temp_file.name
    try:
        result = subprocess.run([antiword, temp_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
        text = result.stdout.decode('utf-8', errors='ignore')
        if result.returncode != 0:
            return text, result.stderr.decode('utf-8', errors='ignore') or 'antiword failed'
        return text, None
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def extract_modern_docx(file_bytes) -> tuple[str, str | None]:
    try:
        doc = Document(io.BytesIO(file_bytes))
        parts = []
        for table in doc.tables:
            for row in table.rows:
                vals = [c.text.strip() for c in row.cells if c.text.strip()]
                if vals:
                    parts.append(' | '.join(vals))
        for para in doc.paragraphs:
            if para.text.strip():
                parts.append(para.text.strip())
        return "\n".join(parts), None
    except Exception as e:
        return "", f"DOCX extraction failed: {e}"


def extract_word_text(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    if uploaded_file.name.lower().endswith('.doc'):
        return extract_legacy_doc(file_bytes)
    return extract_modern_docx(file_bytes)


def split_tec_segments(raw_text: str) -> list[str]:
    text = re.sub(r'\s+', ' ', raw_text)
    text = re.sub(rf'({MONTHS}(?:\s+\d{{4}})?)', r'\n\1', text)
    text = re.sub(r'(WEEK\s+\d+)', r'\n\1', text)
    text = re.sub(r'(?<!\d)(\d{1,3})(?=[A-Z])', r'\n\1', text)
    return [seg.strip() for seg in text.split('\n') if seg.strip()]


def process_tec19_files(uploaded_files):
    all_extracted_data = []
    quarantine = []
    raw_dump = []

    for file in uploaded_files:
        raw_text, err = extract_word_text(file)
        raw_dump.append({"Source": file.name, "Raw Text": raw_text[:15000], "Error": err or ""})

        if err and not raw_text:
            quarantine.append({"Source": file.name, "Reason": err, "Raw Segment": ""})
            continue

        month = None
        week = None
        for seg in split_tec_segments(raw_text):
            if re.match(rf'^{MONTHS}', seg):
                month = seg
                continue
            m_week = re.match(r'^WEEK\s+(\d+)', seg)
            if m_week:
                week = int(m_week.group(1))
                continue

            dates = re.findall(r'\b\d{2}-\d{2}-\d{2}\b', seg)
            text_wo_dates = re.sub(r'\b\d{2}-\d{2}-\d{2}\b', ' ', seg)
            text_wo_dates = re.sub(r'\bCE\b', ' ', text_wo_dates, flags=re.I)
            text_wo_dates = re.sub(r'\s+', ' ', text_wo_dates).strip()
            norm = normalize_text(text_wo_dates)

            if len(norm) < 8 or len(norm.split()) < 2:
                quarantine.append({"Source": file.name, "Reason": "too_short_or_noisy", "Raw Segment": seg})
                continue

            if dates:
                completed_date = dates[-1]
                all_extracted_data.append({
                    "Text": text_wo_dates,
                    "Normalized": norm,
                    "Date": completed_date,
                    "ParsedDate": parse_date_safely(completed_date),
                    "Source": file.name,
                    "Source Month": month,
                    "Week": week,
                    "Segment Type": "SCHEDULED_OR_DATED"
                })
            else:
                all_extracted_data.append({
                    "Text": text_wo_dates,
                    "Normalized": norm,
                    "Date": "",
                    "ParsedDate": None,
                    "Source": file.name,
                    "Source Month": month,
                    "Week": week,
                    "Segment Type": "FREE_TEXT_EXECUTION"
                })

    return all_extracted_data, quarantine, raw_dump


def spatial_lock_pms(excel_bytes) -> tuple[pd.DataFrame | None, str | None]:
    try:
        try:
            df_raw = pd.read_excel(io.BytesIO(excel_bytes), engine='openpyxl', header=None)
        except Exception:
            df_raw = pd.read_excel(io.BytesIO(excel_bytes), header=None)

        header_idx = 0
        for idx, row in df_raw.head(30).iterrows():
            row_text = " ".join(str(val).upper() for val in row.values)
            if any(k in row_text for k in ["COMPONENT", "DESCRIPTION", "OVERHAUL", "DATE", "HOUR", "RUN"]):
                header_idx = idx
                break

        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), engine='openpyxl', header=header_idx)
        except Exception:
            df = pd.read_excel(io.BytesIO(excel_bytes), header=header_idx)

        df.columns = df.columns.astype(str).str.strip().str.upper()

        target_cols = {}
        for col in df.columns:
            c = str(col).upper()
            if "COMPONENT" in c or "DESCRIPTION" in c or "EQUIPMENT" in c or c == "NAME":
                target_cols.setdefault("Component Name", col)
            elif ("OVERHAUL" in c or "LAST" in c or "DATE" in c) and "HOUR" not in c:
                target_cols.setdefault("Last Overhaul Date", col)
            elif "HOUR" in c or "HRS" in c or "RUN" in c:
                target_cols.setdefault("Total Running Hours", col)

        if len(target_cols) < 3:
            return None, f"Could not map required columns. Found headers: {list(df.columns)}"

        df = df.rename(columns={v: k for k, v in target_cols.items()})
        out = df[PMS_LOCKED_COLS].copy()
        out["Component Name"] = out["Component Name"].astype(str).str.strip()
        out = out[out["Component Name"].ne("")]
        out["Normalized"] = out["Component Name"].map(normalize_text)
        out["ParsedDate"] = out["Last Overhaul Date"].map(parse_date_safely)
        out["Total Running Hours"] = pd.to_numeric(out["Total Running Hours"], errors='coerce').fillna(0.0)
        return out.dropna(subset=["Component Name"]), None
    except Exception as e:
        return None, f"Fatal Excel extraction error: {str(e)}"


def apply_weighted_shield(text: str) -> bool:
    upper_text = normalize_text(text)
    if any(action in upper_text for action in ACTION_WORDS):
        return True
    if any(insp in upper_text for insp in INSPECTION_WORDS):
        return False
    return False


def generate_phonetic_hash(text: str) -> set:
    clean = normalize_text(text)
    words = [w for w in clean.split() if len(w) > 2 and not w.isdigit()]
    return {jellyfish.metaphone(word) for word in words if jellyfish.metaphone(word)}


def run_audit(pms_df, tec_data, audit_date):
    results = {
        "Syncs": [], "Ghosts": [], "Unlogged_Hours": [],
        "Quarantine": [], "TEC_Filtered_Out": []
    }

    filtered_tec = []
    for log in tec_data:
        if apply_weighted_shield(log["Text"]):
            item = log.copy()
            item["Hashes"] = generate_phonetic_hash(item["Normalized"])
            filtered_tec.append(item)
        else:
            results["TEC_Filtered_Out"].append({
                "Source": log["Source"],
                "Text": log["Text"],
                "Reason": "inspection_or_non_action"
            })

    if not filtered_tec:
        results["Quarantine"].append({"System Alert": "No valid action entries survived the Anti-Inspection Shield."})
        return results

    for _, row in pms_df.iterrows():
        pms_comp = str(row['Component Name'])
        pms_date_str = str(row['Last Overhaul Date'])
        pms_date = row.get('ParsedDate', None)
        pms_hours = float(row['Total Running Hours']) if pd.notna(row['Total Running Hours']) else 0.0
        pms_hashes = generate_phonetic_hash(row['Normalized'])

        if not pms_hashes:
            results["Quarantine"].append({
                "Component": pms_comp,
                "Conflict": "No usable phonetic tokens generated for PMS component"
            })
            continue

        best_match = None
        best_score = 0.0
        for log in filtered_tec:
            if not log.get("Hashes"):
                continue
            intersection = pms_hashes.intersection(log["Hashes"])
            score = len(intersection) / m
