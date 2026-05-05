import io
import re
import difflib
from dataclasses import dataclass
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title='POSEIDON PMS Auditor', layout='wide', initial_sidebar_state='expanded')

VESSEL_CYLINDER_COUNTS = {
    'MV ALEXIS': 5,
    'MV BAY': 6,
    'MV CHRISTIANNA': 5,
    'MV CHRISTINA V': 5,
    'MV COURAGE': 6,
    'MV DIGNITY': 7,
    'MV FALCON': 6,
    'MV GEORGIA T': 7,
    'MV GLORY': 5,
    'MV HILL': 6,
    'MV MARATHON': 5,
    'MV PIONEER': 6,
    'MV SEA': 5,
    'MV SKY': 5,
    'MV STEFANOS T': 7,
}

ALIASES = {
    'M.E.': 'MAIN ENGINE',
    'ME': 'MAIN ENGINE',
    'DG': 'DIESEL GENERATOR',
    'AIR COND.': 'AIR CONDITION',
    'A/C': 'AIR CONDITION',
    'AIR CONDITIONED': 'AIR CONDITION',
    'REFRIGERATING': 'REFRIGERATION',
    'COMP.': 'COMPRESSOR',
    'L.O.': 'LUBE OIL',
    'F.O.': 'FUEL OIL',
    'FW': 'FRESH WATER',
    'EXH.': 'EXHAUST',
    'GS': 'GENERAL SERVICE',
    'TC': 'TURBOCHARGER',
}

FAMILIES = {
    'FUEL VALVE': 'VALVE',
    'EXHAUST VALVE': 'VALVE',
    'STARTING VALVE': 'VALVE',
    'SAFETY VALVE': 'VALVE',
    'FUEL PUMP': 'PUMP',
    'COMPRESSOR': 'COMPRESSOR',
    'AIR COOLER': 'COOLER',
    'COOLER': 'COOLER',
    'PUMP': 'PUMP',
    'FILTER': 'FILTER',
    'MOTOR': 'MOTOR',
    'FAN': 'FAN',
    'PISTON': 'PISTON',
    'LINER': 'LINER',
    'CYLINDER COVER': 'CYLINDER_COVER',
    'BEARING': 'BEARING',
    'TURBOCHARGER': 'TURBOCHARGER',
}

RESET_ACTIONS = {'OVERHAUL', 'REPLACE', 'RENEW', 'CHANGE', 'PULLED OUT', 'PULL OUT', 'RECONDITIONED'}
MONTHS = 'JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER'


@dataclass
class ParseResult:
    records: pd.DataFrame
    quarantine: pd.DataFrame
    errors: list


def normalize_text(text: str) -> str:
    t = (text or '').upper().strip()
    for k, v in sorted(ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        t = re.sub(rf'\b{re.escape(k)}\b', v, t)
    t = re.sub(r'[^A-Z0-9\s]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def extract_unit_no(text: str):
    m = re.search(r'\b(?:NO|UNIT)\s*(\d+)\b', text)
    return int(m.group(1)) if m else None


def extract_cylinder_no(text: str):
    m = re.search(r'\bCYL(?:INDER)?\s*(?:NO)?\s*(\d+)\b', text)
    return int(m.group(1)) if m else None


def classify_family(text: str) -> str:
    for token, fam in FAMILIES.items():
        if token in text:
            return fam
    return 'OTHER'


def detect_action(text: str) -> str:
    actions = ['OVERHAUL', 'REPLACE', 'RENEW', 'CHANGE', 'PULLED OUT', 'PULL OUT', 'RECONDITIONED', 'CHECK', 'TEST', 'CLEAN', 'INSPECTION']
    for a in actions:
        if a in text:
            return a
    return 'UNSPECIFIED'


def classify_scope(text: str, cylinder_no, unit_no) -> str:
    if cylinder_no is not None:
        return 'CYLINDER'
    if unit_no is not None:
        return 'UNIT'
    if 'MAIN ENGINE' in text or 'DIESEL GENERATOR' in text:
        return 'ENGINE_LEVEL'
    return 'SYSTEM'


def extract_docx_text(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    parts = []
    for table in doc.tables:
        for row in table.rows:
            vals = [c.text.strip() for c in row.cells if c.text.strip()]
            if vals:
                parts.append(' '.join(vals))
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    return '\n'.join(parts)


def parse_tec_text(raw_text: str, source_file: str, vessel_name: str):
    text = re.sub(r'\s+', ' ', raw_text)
    text = re.sub(rf'({MONTHS}(?:\s+\d{{4}})?)', r'\n\1', text)
    text = re.sub(r'(WEEK\s+\d+)', r'\n\1', text)
    text = re.sub(r'(?<!\d)(\d{1,3})(?=[A-Z])', r'\n\1', text)

    month = None
    week = None
    rows = []
    quarantine = []

    for seg in [x.strip() for x in text.split('\n') if x.strip()]:
        if re.match(rf'^{MONTHS}', seg):
            month = seg
            continue
        week_m = re.match(r'^WEEK\s+(\d+)', seg)
        if week_m:
            week = int(week_m.group(1))
            continue
        dates = re.findall(r'\d{2}-\d{2}-\d{2}', seg)
        clean = re.sub(r'\d{2}-\d{2}-\d{2}', ' ', seg)
        clean = re.sub(r'\bCE\b', ' ', clean, flags=re.I)
        clean = re.sub(r'\s+', ' ', clean).strip()
        if len(clean) < 6:
            quarantine.append({'source_file': source_file, 'raw_text': seg, 'reason': 'too_short'})
            continue
        norm = normalize_text(clean)
        cyl = extract_cylinder_no(norm)
        unit = extract_unit_no(norm)
        rows.append({
            'vessel_name': vessel_name,
            'source_file': source_file,
            'source_month': month,
            'week_no': week,
            'raw_text': clean,
            'normalized_text': norm,
            'deadline_date': dates[0] if len(dates) >= 1 else None,
            'issued_date': dates[1] if len(dates) >= 2 else None,
            'action_type': detect_action(norm),
            'equipment_family': classify_family(norm),
            'unit_no': unit,
            'cylinder_no': cyl,
            'scope_type': classify_scope(norm, cyl, unit),
            'row_kind': 'SCHEDULED' if dates else 'EXECUTED',
        })
    return pd.DataFrame(rows), pd.DataFrame(quarantine)


def parse_running_hours_docx(raw_text: str, vessel_name: str):
    txt = re.sub(r'\s+', ' ', raw_text.upper())
    rows = []
    quarantine = []

    generic_patterns = [
        ('COMPRESSOR', r'(AIR COND\.? COMPRESSOR NO\.?\s*\d|REFRIGERATION COMPRESSOR NO\.?\s*\d|STARTING MAIN AIR COMPRESSOR NO\.?\s*\d|SERVICE AIR COMPRESSOR|EMERGENCY AIR COMPRESSOR NO\.?\s*\d)\s*(\d{1,2}\s+[A-Z]+\.?\s+\d{2,4})\s*(\d{1,6})'),
        ('TURBOCHARGER', r'(GENERAL OH|BALANCING OF ROTOR SHAFT|AIR COOLER CLEANING)\s*(\d{1,2}\s+[A-Z]+\.?\s+\d{2,4})\s*(\d{1,6})'),
        ('AUXILIARY', r'(FURNACE INSPECTION|BURNER ATOMIZER|FORCED DRAFT FAN|WASHING THE TUBES)\s*(\d{1,2}\s+[A-Z]+\.?\s+\d{2,4})'),
    ]

    for section, pattern in generic_patterns:
        for m in re.finditer(pattern, txt):
            comp = m.group(1).strip()
            last_date = m.group(2).strip() if m.lastindex and m.lastindex >= 2 else None
            hrs = m.group(3).strip() if m.lastindex and m.lastindex >= 3 else None
            norm = normalize_text(comp)
            cyl = extract_cylinder_no(norm)
            unit = extract_unit_no(norm)
            rows.append({
                'vessel_name': vessel_name,
                'component_raw': comp,
                'component_normalized': norm,
                'equipment_family': classify_family(norm),
                'unit_no': unit,
                'cylinder_no': cyl,
                'scope_type': classify_scope(norm, cyl, unit),
                'section_name': section,
                'last_oh_date_raw': last_date,
                'hours_since_oh_raw': hrs,
                'periodicity_hours': None,
            })

    me_items = ['CYLINDER COVER', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 'CYLINDER LINER', 'EXHAUST VALVE', 'STARTING VALVE', 'SAFETY VALVE', 'FUEL VALVES', 'FUEL PUMP', 'PLUNGER AND BARREL RENEWAL', 'FUEL PUMP SUCTION VALVE', 'FUEL PUMP PUNCTURE VALVE', 'CROSSHEAD BEARINGS', 'BOTTOM END BEARINGS', 'MAIN BEARINGS']
    for item in me_items:
        if item in txt:
            norm = normalize_text(item)
            rows.append({
                'vessel_name': vessel_name,
                'component_raw': item,
                'component_normalized': norm,
                'equipment_family': classify_family(norm),
                'unit_no': None,
                'cylinder_no': None,
                'scope_type': classify_scope(norm, None, None),
                'section_name': 'MAIN_ENGINE',
                'last_oh_date_raw': None,
                'hours_since_oh_raw': None,
                'periodicity_hours': None,
            })

    if not rows:
        quarantine.append({'reason': 'no_running_hours_patterns_found', 'raw_text': raw_text[:1000]})

    df = pd.DataFrame(rows).drop_duplicates()
    return df, pd.DataFrame(quarantine)


def parse_tec_files(uploaded_files, vessel_name: str) -> ParseResult:
    frames, quarantines, errors = [], [], []
    for f in uploaded_files:
        try:
            if not f.name.lower().endswith('.docx'):
                raise ValueError('Only .docx TEC files are supported in the 2-file version')
            raw = extract_docx_text(f.getvalue())
            df, q = parse_tec_text(raw, source_file=f.name, vessel_name=vessel_name)
            frames.append(df)
            quarantines.append(q)
        except Exception as e:
            errors.append(f'{f.name}: {e}')
    return ParseResult(
        records=pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(),
        quarantine=pd.concat(quarantines, ignore_index=True) if quarantines else pd.DataFrame(),
        errors=errors,
    )


def parse_pms_file(uploaded_file, vessel_name: str) -> ParseResult:
    errors = []
    quarantine = pd.DataFrame()
    try:
        if uploaded_file.name.lower().endswith('.docx'):
            raw = extract_docx_text(uploaded_file.getvalue())
            records, quarantine = parse_running_hours_docx(raw, vessel_name=vessel_name)
        elif uploaded_file.name.lower().endswith(('.xlsx', '.xls')):
            records = pd.read_excel(uploaded_file, engine='openpyxl')
            records.columns = [str(c).strip() for c in records.columns]
            if 'Component Name' not in records.columns:
                cands = [c for c in records.columns if 'component' in c.lower() or 'description' in c.lower()]
                if cands:
                    records = records.rename(columns={cands[0]: 'Component Name'})
            if 'Component Name' not in records.columns:
                raise ValueError('Component Name column missing in PMS spreadsheet')
            records['component_raw'] = records['Component Name'].astype(str)
            records['component_normalized'] = records['component_raw'].map(normalize_text)
            records['equipment_family'] = records['component_normalized'].map(classify_family)
            records['unit_no'] = re
