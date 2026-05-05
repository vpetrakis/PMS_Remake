import io
import re
import shutil
import subprocess
import tempfile
import os
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title='POSEIDON Parser Workbench', layout='wide', initial_sidebar_state='expanded')

MONTHS = 'JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER'
ALIASES = {
    'M.E.': 'MAIN ENGINE',
    'ME': 'MAIN ENGINE',
    'DG': 'DIESEL GENERATOR',
    'AIR COND.': 'AIR CONDITION',
    'AIR CONDITIONED': 'AIR CONDITION',
    'COMP.': 'COMPRESSOR',
    'L.O.': 'LUBE OIL',
    'F.O.': 'FUEL OIL',
    'EXH.': 'EXHAUST',
    'TC': 'TURBOCHARGER',
}


def normalize_text(text: str) -> str:
    t = (text or '').upper().strip()
    for k, v in sorted(ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        t = re.sub(rf'\b{re.escape(k)}\b', v, t)
    t = re.sub(r'[^A-Z0-9\s\-\./]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def extract_word_text(uploaded_file):
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()
    if name.endswith('.docx'):
        doc = Document(io.BytesIO(raw))
        parts = []
        for table in doc.tables:
            for row in table.rows:
                vals = [c.text.strip() for c in row.cells if c.text.strip()]
                if vals:
                    parts.append(' | '.join(vals))
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text.strip())
        return '\n'.join(parts), None
    if name.endswith('.doc'):
        antiword = shutil.which('antiword')
        if not antiword:
            return None, 'antiword is not installed; .doc parsing unavailable in this environment'
        with tempfile.NamedTemporaryFile(delete=False, suffix='.doc') as tmp:
            tmp.write(raw)
            tmp_path = tmp.name
        try:
            proc = subprocess.run([antiword, tmp_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
            if proc.returncode != 0:
                return None, proc.stderr.decode('utf-8', errors='ignore') or 'antiword failed'
            return proc.stdout.decode('utf-8', errors='ignore'), None
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    return None, 'Unsupported file type'


def split_tec_segments(text: str):
    t = re.sub(r'\s+', ' ', text)
    t = re.sub(rf'({MONTHS}(?:\s+\d{{4}})?)', r'\n\1', t)
    t = re.sub(r'(WEEK\s+\d+)', r'\n\1', t)
    t = re.sub(r'(?<!\d)(\d{1,3})(?=[A-Z])', r'\n\1', t)
    return [x.strip() for x in t.split('\n') if x.strip()]


def parse_tec_file(uploaded_file, vessel_name: str):
    raw_text, err = extract_word_text(uploaded_file)
    if err:
        return pd.DataFrame(), pd.DataFrame([{'source_file': uploaded_file.name, 'reason': err}]), raw_text or ''

    month = None
    week = None
    rows = []
    quarantine = []
    for seg in split_tec_segments(raw_text):
        if re.match(rf'^{MONTHS}', seg):
            month = seg
            continue
        m_week = re.match(r'^WEEK\s+(\d+)', seg)
        if m_week:
            week = int(m_week.group(1))
            continue
        dates = re.findall(r'\d{2}-\d{2}-\d{2}', seg)
        clean = re.sub(r'\d{2}-\d{2}-\d{2}', ' ', seg)
        clean = re.sub(r'\bCE\b', ' ', clean, flags=re.I)
        clean = re.sub(r'\s+', ' ', clean).strip()
        norm = normalize_text(clean)
        if len(norm) < 5:
            quarantine.append({'source_file': uploaded_file.name, 'raw_segment': seg, 'reason': 'too_short'})
            continue
        rows.append({
            'vessel_name': vessel_name,
            'source_file': uploaded_file.name,
            'source_month': month,
            'week_no': week,
            'raw_segment': seg,
            'clean_text': clean,
            'normalized_text': norm,
            'date_count': len(dates),
            'dates_found': ' | '.join(dates),
            'segment_type': 'SCHEDULED' if dates else 'EXECUTED_OR_FREE_TEXT'
        })
    return pd.DataFrame(rows), pd.DataFrame(quarantine), raw_text


def parse_pms_file(uploaded_file, vessel_name: str):
    name = uploaded_file.name.lower()
    if name.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(uploaded_file, engine='openpyxl' if name.endswith('.xlsx') else None)
        df.columns = [str(c).strip() for c in df.columns]
        preview = df.copy()
        if len(preview.columns) > 0:
            preview['__row_text__'] = preview.astype(str).agg(' | '.join, axis=1)
            preview['__normalized__'] = preview['__row_text__'].map(normalize_text)
        quarantine = pd.DataFrame()
        return preview, quarantine, 'SPREADSHEET MODE'

    raw_text, err = extract_word_text(uploaded_file)
    if err:
        return pd.DataFrame(), pd.DataFrame([{'source_file': uploaded_file.name, 'reason': err}]), raw_text or ''

    txt = re.sub(r'\s+', ' ', raw_text)
    sections = []
    headings = ['MAIN ENGINE', 'TURBOCHARGER', 'COOLERS', 'AC REFR. COMPRESSORS', 'AUXILIARY BOILER', 'EXH GAS BOILER', 'MAIN AIR COMPRESSORS', 'AUX. ENGINE']
    for h in headings:
        if h in txt.upper():
            sections.append(h)

    blocks = []
    for h in sections:
        idx = txt.upper().find(h)
        snippet = txt[idx: idx + 1000] if idx >= 0 else ''
        blocks.append({'section_name': h, 'raw_excerpt': snippet, 'normalized_excerpt': normalize_text(snippet)})

    quarantine = pd.DataFrame()
    if not blocks:
        quarantine = pd.DataFrame([{'source_file': uploaded_file.name, 'reason': 'no_known_sections_detected'}])
    return pd.DataFrame(blocks), quarantine, raw_text


st.title('POSEIDON Parser Workbench')
st.caption('Revised parser-first build: stable ingestion, extraction visibility, and quarantine before matching.')

with st.sidebar:
    st.header('Inputs')
    vessel_name = st.text_input('Vessel name', value='MV ALEXIS')
    pms_file = st.file_uploader('PMS / Running Hours (.doc, .docx, .xlsx, .xls)', type=['doc', 'docx', 'xlsx', 'xls'])
    tec_files = st.file_uploader('TEC Logs (.doc, .docx)', type=['doc', 'docx'], accept_multiple_files=True)
    run_btn = st.button('Parse files', type='primary', use_container_width=True)

if not run_btn:
    st.info('Upload files and run the parser workbench. This version intentionally stops before fuzzy matching and forensic scoring.')
    st.stop()

if not pms_file and not tec_files:
    st.error('Upload at least one PMS file or one TEC file.')
    st.stop()

pms_df, pms_q, pms_raw = pd.DataFrame(), pd.DataFrame(), ''
if pms_file:
    pms_df, pms_q, pms_raw = parse_pms_file(pms_file, vessel_name)

tec_frames = []
tec_q_frames = []
tec_raw_bundle = []
for f in tec_files or []:
    df, q, raw = parse_tec_file(f, vessel_name)
    tec_frames.append(df)
    tec_q_frames.append(q)
    tec_raw_bundle.append({'file_name': f.name, 'raw_text': raw[:15000]})

tec_df = pd.concat(tec_frames, ignore_index=True) if tec_frames else pd.DataFrame()
tec_q = pd.concat(tec_q_frames, ignore_index=True) if tec_q_frames else pd.DataFrame()

c1, c2, c3, c4 = st.columns(4)
c1.metric('PMS parsed rows', len(pms_df))
c2.metric('PMS quarantine', len(pms_q))
c3.metric('TEC parsed rows', len(tec_df))
c4.metric('TEC quarantine', len(tec_q))

tabs = st.tabs(['Overview', 'PMS Parsed', 'PMS Quarantine', 'TEC Parsed', 'TEC Quarantine', 'Raw Text'])

with tabs[0]:
    st.subheader('Parser summary')
    st.dataframe(pd.DataFrame([{
        'timestamp': datetime.now().isoformat(timespec='seconds'),
        'vessel_name': vessel_name,
        'pms_file': pms_file.name if pms_file else None,
        'tec_files': ', '.join([f.name for f in tec_files]) if tec_files else None,
        'pms_rows': len(pms_df),
        'pms_quarantine': len(pms_q),
        'tec_rows': len(tec_df),
        'tec_quarantine': len(tec_q),
    }]), use_container_width=True)
    st.markdown('This build is intentionally conservative: it only proves ingestion and extraction quality before any matching logic is reintroduced.')

with tabs[1]:
    st.subheader('PMS parsed output')
    st.dataframe(pms_df, use_container_width=True, height=520)

with tabs[2]:
    st.subheader('PMS quarantine')
    st.dataframe(pms_q, use_container_width=True, height=520)

with tabs[3]:
    st.subheader('TEC parsed output')
    st.dataframe(tec_df, use_container_width=True, height=520)

with tabs[4]:
    st.subheader('TEC quarantine')
    st.dataframe(tec_q, use_container_width=True, height=520)

with tabs[5]:
    st.subheader('Raw text inspection')
    if pms_file:
        st.markdown(f'### PMS raw text: {pms_file.name}')
        st.text_area('PMS raw', pms_raw[:20000], height=250)
    for item in tec_raw_bundle:
        st.markdown(f"### TEC raw text: {item['file_name']}")
        st.text_area(item['file_name'], item['raw_text'], height=250)
