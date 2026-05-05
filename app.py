"""
TEMPORAL PIPELINE v2.0 — Zero-Trust Maritime PMS Audit Engine
MV ALEXIS | MTS Marine Ltd. | Production Release
"""

import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import struct
from datetime import datetime, date
from docx import Document
from rapidfuzz import fuzz, process

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Temporal Pipeline",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Maritime abbreviation normalization dictionary
MARITIME_ABBREV = {
    r'\bM/?E\b': 'MAIN ENGINE',
    r'\bD/?G\b': 'DIESEL GENERATOR',
    r'\bA/?E\b': 'AUX ENGINE',
    r'\bF/?O\b': 'FUEL OIL',
    r'\bL/?O\b': 'LUBE OIL',
    r'\bH/?F/?O\b': 'HEAVY FUEL OIL',
    r'\bM/?D/?O\b': 'MARINE DIESEL OIL',
    r'\bF/?W\b': 'FRESH WATER',
    r'\bS/?W\b': 'SEA WATER',
    r'\bT/?C\b': 'TURBOCHARGER',
    r'\bC/?W\b': 'COOLING WATER',
    r'\bA/?C\b': 'AIR CONDITIONING',
    r'\bEXH\b': 'EXHAUST',
    r'\bINJ\b': 'INJECTOR',
    r'\bCYL\b': 'CYLINDER',
    r'\bASS\'?Y\b': 'ASSEMBLY',
    r'\bELECT\b': 'ELECTRIC',
    r'\bMOTOT?\b': 'MOTOR',
    r'\bPURIF\b': 'PURIFIER',
    r'\bPRESSURE\b': 'PRESSURE',
    r'\bPUMP\b': 'PUMP',
    r'\bFILTER\b': 'FILTER',
    r'\bVALVE\b': 'VALVE',
    r'\bCOOLER\b': 'COOLER',
    r'\bNO\.?\s*(\d+)\b': r'NUMBER \1',
    r'#\s*(\d+)': r'NUMBER \1',
}

# Action words that indicate maintenance was performed (not just inspection)
ACTION_WORDS = {
    'OVERHAUL', 'OVERHAULED', 'REPLACE', 'REPLACED', 'RENEW', 'RENEWED',
    'PULL OUT', 'PULLED OUT', 'DISMANTLE', 'DISMANTLED', 'RECONDITION',
    'RECONDITIONED', 'CHANGE', 'CHANGED', 'FABRICATE', 'FABRICATED',
    'REBUILD', 'REBUILT', 'REPAIR', 'REPAIRED', 'INSTALL', 'INSTALLED',
}

MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

# ─────────────────────────────────────────────────────────────────────────────
# TEXT NORMALIZATION
# ─────────────────────────────────────────────────────────────────────────────

def normalize_maritime_text(text: str) -> str:
    """Expand maritime abbreviations to full form for accurate matching."""
    text = str(text).upper().strip()
    for pattern, replacement in MARITIME_ABBREV.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    # Remove special chars except letters, digits, spaces
    text = re.sub(r'[^\w\s]', ' ', text)
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def is_action_entry(text: str) -> bool:
    """Return True if the text describes actual maintenance work performed."""
    upper = str(text).upper()
    return any(word in upper for word in ACTION_WORDS)

# ─────────────────────────────────────────────────────────────────────────────
# FILE EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_strings_from_binary(data: bytes, min_length: int = 8) -> list[str]:
    """
    Pure-Python equivalent of the Unix `strings` command.
    Finds sequences of printable ASCII characters in binary data.
    Works reliably on old .doc (Word 97-2003) binary files.
    """
    pattern = re.compile(b'[ -~]{' + str(min_length).encode() + b',}')
    results = []
    for match in pattern.finditer(data):
        try:
            s = match.group().decode('ascii', errors='ignore').strip()
            if len(s) >= min_length:
                results.append(s)
        except Exception:
            continue
    return results

def extract_text_from_doc(file_bytes: bytes) -> list[str]:
    """Extract text lines from a binary .doc file using ASCII string scanning."""
    raw_strings = extract_strings_from_binary(file_bytes, min_length=8)
    # Filter for lines that contain alphabetic content
    return [s for s in raw_strings if sum(1 for c in s if c.isalpha()) >= 4]

def extract_text_from_docx(file_bytes: bytes) -> list[str]:
    """Extract text lines from a modern .docx file."""
    doc = Document(io.BytesIO(file_bytes))
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text and sum(1 for c in text if c.isalpha()) >= 4:
            lines.append(text)
    return lines

def extract_lines_from_word(file_bytes: bytes, filename: str) -> list[str]:
    """Route to correct extractor based on file extension."""
    if filename.lower().endswith('.docx'):
        try:
            return extract_text_from_docx(file_bytes)
        except Exception:
            # Fall back to binary extraction
            return extract_text_from_doc(file_bytes)
    else:
        # .doc binary format
        lines = extract_text_from_doc(file_bytes)
        # Also try as docx (some .doc files are actually docx)
        if not lines:
            try:
                return extract_text_from_docx(file_bytes)
            except Exception:
                pass
        return lines

# ─────────────────────────────────────────────────────────────────────────────
# TEC-19 PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_tec19_files(uploaded_files: list) -> list[dict]:
    """
    Parse TEC-19 work diary Word documents.
    Returns list of {text, date, normalized_text, is_action, source}
    """
    DATE_PATTERN = re.compile(r'\b(\d{2})[/-](\d{2})[/-](\d{2,4})\b')
    all_entries = []

    for f in uploaded_files:
        try:
            file_bytes = f.getvalue()
            lines = extract_lines_from_word(file_bytes, f.name)

            if not lines:
                st.warning(f"⚠️ No readable text extracted from: {f.name}")
                continue

            # State-machine parser: associate each job description with its nearest date
            current_date = None
            job_buffer = []

            for line in lines:
                dates_in_line = DATE_PATTERN.findall(line)

                if dates_in_line:
                    # Flush buffered job text if we had one
                    if job_buffer and current_date:
                        job_text = ' '.join(job_buffer).strip()
                        if len(job_text) > 8 and is_action_entry(job_text):
                            all_entries.append({
                                'text': job_text,
                                'date': current_date,
                                'normalized_text': normalize_maritime_text(job_text),
                                'is_action': True,
                                'source': f.name,
                            })
                    job_buffer = []

                    # Parse and store the date(s) from this line
                    for d, m, y in dates_in_line:
                        day, month = int(d), int(m)
                        year = int(y)
                        if year < 100:
                            year += 2000
                        if 1 <= day <= 31 and 1 <= month <= 12:
                            try:
                                current_date = datetime(year, month, day)
                                break
                            except ValueError:
                                continue

                    # The line itself might also contain a job description (minus the date)
                    job_part = DATE_PATTERN.sub('', line).strip()
                    if len(job_part) > 8 and sum(c.isalpha() for c in job_part) > 4:
                        job_buffer.append(job_part)
                else:
                    # Pure description line
                    if len(line) > 8:
                        job_buffer.append(line)

                    # If we have a big enough job description and a date, flush it
                    combined = ' '.join(job_buffer).strip()
                    if len(combined) > 15 and current_date and is_action_entry(combined):
                        all_entries.append({
                            'text': combined,
                            'date': current_date,
                            'normalized_text': normalize_maritime_text(combined),
                            'is_action': True,
                            'source': f.name,
                        })
                        job_buffer = []

            # Flush any remaining buffer
            if job_buffer and current_date:
                job_text = ' '.join(job_buffer).strip()
                if len(job_text) > 8 and is_action_entry(job_text):
                    all_entries.append({
                        'text': job_text,
                        'date': current_date,
                        'normalized_text': normalize_maritime_text(job_text),
                        'is_action': True,
                        'source': f.name,
                    })

        except Exception as e:
            st.error(f"❌ Error parsing {f.name}: {e}")

    # Deduplicate identical entries
    seen = set()
    unique_entries = []
    for entry in all_entries:
        key = (entry['text'][:80], entry['date'])
        if key not in seen:
            seen.add(key)
            unique_entries.append(entry)

    return unique_entries

# ─────────────────────────────────────────────────────────────────────────────
# PMS EXCEL PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_pms_excel(file_bytes: bytes) -> tuple[pd.DataFrame | None, str | None, dict]:
    """
    Parse the PMS Excel file with the MTS Marine structure.

    Returns:
        (df, error_message, column_map)
        df columns: code, component, job_type, interval_hrs, last_oh_date,
                    hrs_end_last_yr, current_hrs, monthly_jan..dec, normalized_component
    """
    meta = {}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, f"Cannot open Excel file: {e}", {}

    # Suppress openpyxl warnings
    import warnings
    warnings.filterwarnings('ignore', category=UserWarning)

    # Try PMS sheet first, then any sheet with component data
    target_sheet = None
    for name in wb.sheetnames:
        if 'PMS' in name.upper():
            target_sheet = wb[name]
            break
    if target_sheet is None:
        target_sheet = wb.active

    ws = target_sheet
    meta['sheet_name'] = ws.title

    # ── Find header row (look for 'CODE' or 'ITEMS' keyword) ──
    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=20, values_only=True):
        row_text = ' '.join(str(v).upper() for v in row if v)
        if ('CODE' in row_text or 'ITEMS' in row_text) and 'INTERVAL' in row_text:
            header_row_idx = row[0]  # store for reference
            break
        if ('CODE' in row_text) and 'INSPECTION' in row_text:
            break

    # ── Read with pandas, searching for the actual header ──
    df_scan = pd.read_excel(io.BytesIO(file_bytes), sheet_name=ws.title,
                            engine='openpyxl', header=None)

    header_idx = 7  # Default (row 8 in 1-indexed = idx 7)
    for idx, row in df_scan.head(15).iterrows():
        row_str = ' '.join(str(v).upper() for v in row if pd.notna(v))
        if ('ITEMS' in row_str or 'CODE' in row_str) and 'INTERVAL' in row_str:
            header_idx = idx
            break

    # ── Re-read with correct header ──
    # Use data_only to get computed formula values
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=ws.title,
        engine='openpyxl',
        header=header_idx,
        dtype={0: str}  # Keep code as string
    )

    # Normalize column names
    df.columns = [str(c).strip().upper() if pd.notna(c) else f'COL_{i}'
                  for i, c in enumerate(df.columns)]

    meta['raw_columns'] = list(df.columns[:22])

    # ── Map columns to standard names ──
    col_map = {}

    def find_col(keywords, df_cols):
        for keyword in keywords:
            for col in df_cols:
                if keyword in col:
                    return col
        return None

    cols = list(df.columns)

    # Code column: typically first col, format ME-01-01
    col_map['code'] = cols[0] if cols else None

    # Component name: ITEMS or DESCRIPTION
    col_map['component'] = (find_col(['ITEMS', 'DESCRIPTION', 'EQUIPMENT', 'NAME'], cols)
                            or cols[1] if len(cols) > 1 else None)

    # Job type
    col_map['job'] = find_col(['JOB'], cols) or (cols[2] if len(cols) > 2 else None)

    # Interval hours
    col_map['interval'] = find_col(['INTERVAL'], cols) or (cols[3] if len(cols) > 3 else None)

    # Date of last inspection/overhaul
    col_map['date'] = find_col(['DATE OF LAST', 'LAST INSPECTION', 'LAST OVERHAUL'], cols)
    if not col_map['date']:
        # Look for a column containing dates
        for col in cols[4:10]:
            if 'DATE' in col:
                col_map['date'] = col
                break

    # Operating hours at end of last year
    col_map['hrs_end_yr'] = find_col(['OPERATING HOURS AT', 'END OF LAST'], cols)
    if not col_map['hrs_end_yr']:
        # It's typically the column after the date column
        if col_map['date']:
            date_pos = list(df.columns).index(col_map['date'])
            if date_pos + 1 < len(cols):
                col_map['hrs_end_yr'] = cols[date_pos + 1]

    # Current operating hours
    col_map['current_hrs'] = find_col(['CURRENT OPERAT'], cols)
    if not col_map['current_hrs']:
        if col_map['hrs_end_yr']:
            pos = list(df.columns).index(col_map['hrs_end_yr'])
            if pos + 1 < len(cols):
                col_map['current_hrs'] = cols[pos + 1]

    # Monthly hours: Jan through Dec (12 consecutive columns after current_hrs or est. date)
    monthly_cols = {}
    month_patterns = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                      'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

    for i, (month_key, pattern) in enumerate(zip(
            ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
             'jul', 'aug', 'sep', 'oct', 'nov', 'dec'],
            month_patterns)):
        found = find_col([pattern], cols)
        if found:
            monthly_cols[month_key] = found
        else:
            # Fallback: use positional (cols 9-20 in the known structure)
            target_idx = 9 + i
            if target_idx < len(cols):
                monthly_cols[month_key] = cols[target_idx]

    meta['col_map'] = col_map
    meta['monthly_cols'] = monthly_cols

    # ── Build clean DataFrame ──
    required = ['code', 'component', 'date']
    missing = [k for k in required if not col_map.get(k)]
    if missing:
        return None, (f"Could not locate required columns: {missing}. "
                      f"Found headers: {list(df.columns[:15])}"), meta

    # Build the output dataframe
    output_rows = []
    for _, row in df.iterrows():
        code = str(row.get(col_map['code'], '')).strip()
        component = str(row.get(col_map['component'], '')).strip()

        # Only process rows with a valid component code (e.g., ME-01-01)
        if not re.match(r'^[A-Z]{2,4}-\d{2}-\d{2,3}$', code):
            continue
        if not component or component.upper() in ('NAN', 'NONE', ''):
            continue

        # Get date
        raw_date = row.get(col_map['date']) if col_map.get('date') else None
        last_oh_date = None
        if pd.notna(raw_date) and raw_date:
            if isinstance(raw_date, (datetime, date)):
                last_oh_date = raw_date if isinstance(raw_date, datetime) else datetime.combine(raw_date, datetime.min.time())
            else:
                for fmt in ['%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y']:
                    try:
                        last_oh_date = datetime.strptime(str(raw_date)[:10], fmt)
                        break
                    except ValueError:
                        continue

        # Get job type
        job_type = str(row.get(col_map.get('job', ''), '')).strip()

        # Get interval hours
        interval_raw = row.get(col_map.get('interval', ''), None)
        interval_hrs = None
        if pd.notna(interval_raw) and interval_raw:
            try:
                interval_hrs = float(str(interval_raw).replace(',', ''))
            except (ValueError, TypeError):
                interval_hrs = None  # MONTHLY or non-numeric

        # Get hours at end of last year
        hrs_end_yr = None
        if col_map.get('hrs_end_yr'):
            v = row.get(col_map['hrs_end_yr'])
            if pd.notna(v):
                try:
                    hrs_end_yr = float(v)
                except (ValueError, TypeError):
                    pass

        # Get current hours
        current_hrs = None
        if col_map.get('current_hrs'):
            v = row.get(col_map['current_hrs'])
            if pd.notna(v):
                try:
                    current_hrs = float(v)
                except (ValueError, TypeError):
                    pass

        # Get monthly hours
        monthly = {}
        for month_key, col_name in monthly_cols.items():
            v = row.get(col_name)
            if pd.notna(v):
                try:
                    monthly[month_key] = float(v)
                except (ValueError, TypeError):
                    monthly[month_key] = 0.0
            else:
                monthly[month_key] = 0.0

        output_rows.append({
            'code': code,
            'component': component,
            'job_type': job_type,
            'interval_hrs': interval_hrs,
            'last_oh_date': last_oh_date,
            'hrs_end_last_yr': hrs_end_yr,
            'current_hrs': current_hrs,
            'normalized_component': normalize_maritime_text(component),
            **{f'hrs_{k}': v for k, v in monthly.items()}
        })

    if not output_rows:
        return None, "No valid component rows found in PMS sheet.", meta

    result_df = pd.DataFrame(output_rows)
    return result_df, None, meta

# ─────────────────────────────────────────────────────────────────────────────
# FUZZY MATCHING ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def find_best_pms_match(tec_normalized: str, pms_df: pd.DataFrame,
                        threshold: int = 58) -> dict | None:
    """
    Find the best matching PMS component for a TEC diary entry.
    Uses multi-strategy fuzzy matching:
      1. token_set_ratio  – handles word reordering and extra words
      2. token_sort_ratio – handles word reordering
      3. partial_ratio    – handles substring containment
    Returns best match above threshold, or None.
    """
    if pms_df.empty:
        return None

    best_score = 0
    best_row = None

    for _, row in pms_df.iterrows():
        comp_norm = row['normalized_component']

        # Calculate all three scores
        score_set = fuzz.token_set_ratio(tec_normalized, comp_norm)
        score_sort = fuzz.token_sort_ratio(tec_normalized, comp_norm)
        score_partial = fuzz.partial_ratio(tec_normalized, comp_norm)

        # Weighted combination: token_set is most reliable for maritime text
        combined = max(score_set, score_sort * 0.9, score_partial * 0.85)

        if combined > best_score:
            best_score = combined
            best_row = row.copy()

    if best_score >= threshold and best_row is not None:
        best_row['match_score'] = round(best_score, 1)
        return best_row.to_dict()

    return None

# ─────────────────────────────────────────────────────────────────────────────
# PHYSICS VALIDATOR
# ─────────────────────────────────────────────────────────────────────────────

def calculate_max_possible_hours(pms_row: dict, audit_date: datetime) -> tuple:
    """
    Calculate max possible hours and flag if this is a current-year overhaul.

    For current-year overhauls: exact monthly tracking is available.
      max_hours = sum of monthly hours from overhaul month to audit month.

    For prior-year overhauls: we use G (hrs_end_last_yr) as the validated baseline.
      max_hours = hrs_end_last_yr + sum of this year's monthly hours.
      These should match current_hrs by formula definition.

    Returns: (max_hours: float, is_current_year: bool)
    """
    last_oh = pms_row.get('last_oh_date')
    if not last_oh:
        return float('inf'), False

    if isinstance(last_oh, str):
        try:
            last_oh = datetime.fromisoformat(last_oh)
        except Exception:
            return float('inf'), False

    month_keys = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                  'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

    audit_month = min(audit_date.month, 12)
    current_year = audit_date.year

    if last_oh.year == current_year:
        # Current-year overhaul: physics check using exact monthly data
        overhaul_month = last_oh.month
        total = sum(
            pms_row.get(f'hrs_{month_keys[i]}', 0.0) or 0.0
            for i in range(overhaul_month - 1, audit_month)
        )
        return total, True

    elif last_oh.year < current_year:
        # Prior-year overhaul: G-value is the pre-validated baseline
        hrs_end_yr = pms_row.get('hrs_end_last_yr') or 0.0
        current_yr_total = sum(
            pms_row.get(f'hrs_{k}', 0.0) or 0.0
            for k in month_keys[:audit_month]
        )
        return hrs_end_yr + current_yr_total, False

    else:
        return 0.0, False

def classify_physics_violation(current_hrs: float, max_possible: float,
                               interval_hrs: float | None,
                               is_current_year: bool = False) -> tuple[str, str]:
    """
    Returns (violation_type, description).
    Physics violations are ONLY flagged for current-year overhauls where we have
    exact monthly data. Prior-year overhauls use interval checks only.
    """
    if current_hrs is None:
        return 'UNKNOWN', 'Current hours not recorded in PMS'

    # PHYSICS VIOLATION: Only for current-year overhauls
    # (we have exact monthly hour data to compare against)
    if is_current_year and max_possible != float('inf') and max_possible >= 0:
        if current_hrs > max_possible + 10:  # 10hr tolerance for rounding
            excess = current_hrs - max_possible
            return ('PHYSICS_VIOLATION',
                    f"PMS claims {current_hrs:.0f} hrs but only {max_possible:.0f} hrs "
                    f"possible since this year's overhaul (excess: +{excess:.0f} hrs). "
                    f"Hour counter was NOT reset after overhaul.")

    # OVERDUE: Component has exceeded its service interval
    if interval_hrs and current_hrs and interval_hrs > 0 and current_hrs > interval_hrs:
        excess = current_hrs - interval_hrs
        return ('OVERDUE',
                f"Component has run {current_hrs:.0f} hrs against a "
                f"{interval_hrs:.0f} hr service interval (+{excess:.0f} hrs overdue).")

    # HEALTHY
    if is_current_year:
        return 'OK', f"Hours ({current_hrs:.0f} / {max_possible:.0f} max since overhaul)"
    else:
        return 'OK', f"Hours ({current_hrs:.0f} / {interval_hrs:.0f} interval)" if interval_hrs else f"Hours: {current_hrs:.0f}"

# ─────────────────────────────────────────────────────────────────────────────
# MAIN AUDIT ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def run_audit(pms_df: pd.DataFrame, tec_entries: list[dict],
              audit_date: datetime, match_threshold: int) -> dict:
    """
    Cross-reference PMS components against TEC-19 diary entries.
    Returns a 4-tier triage result dict.
    """
    results = {
        'verified_syncs': [],      # Date + hours match, physics OK
        'physics_violations': [],   # Hours physically impossible
        'date_conflicts': [],       # TEC date ≠ PMS date
        'ghost_overhauls': [],      # PMS entry with NO TEC evidence
        'unlogged_jobs': [],        # TEC entry with NO PMS entry
        'missing_hours': [],        # TEC match found but hours not recorded
    }

    # Index TEC entries by normalized text for fast lookup
    tec_by_date = {}
    for entry in tec_entries:
        d = entry['date']
        if d:
            month_key = d.strftime('%Y-%m')
            tec_by_date.setdefault(month_key, []).append(entry)

    matched_tec_ids = set()

    # ── For each PMS component, find matching TEC entry ──
    for _, pms_row in pms_df.iterrows():
        component = pms_row['component']
        code = pms_row['code']
        pms_date = pms_row['last_oh_date']
        current_hrs = pms_row['current_hrs']
        interval_hrs = pms_row['interval_hrs']
        pms_norm = pms_row['normalized_component']

        # Skip components with no overhaul date or NaT
        if pms_date is None or (hasattr(pms_date, 'isnull') and pms_date.isnull()):
            continue
        try:
            _ = pms_date.strftime('%d-%m-%Y')
        except (ValueError, AttributeError):
            continue

        # ── Find best TEC match ──
        best_tec = None
        best_score = 0

        for i, tec_entry in enumerate(tec_entries):
            tec_norm = tec_entry['normalized_text']
            score_set = fuzz.token_set_ratio(pms_norm, tec_norm)
            score_sort = fuzz.token_sort_ratio(pms_norm, tec_norm)
            combined = max(score_set, score_sort * 0.9)

            if combined > best_score:
                best_score = combined
                best_tec = (i, tec_entry)

        if best_score < match_threshold or best_tec is None:
            # No TEC evidence found → Ghost overhaul
            results['ghost_overhauls'].append({
                'Code': code,
                'Component': component,
                'PMS Date': pms_date.strftime('%d-%m-%Y') if pms_date else '—',
                'Current Hrs': f"{current_hrs:.0f}" if current_hrs else '—',
                'Interval Hrs': f"{interval_hrs:.0f}" if interval_hrs else '—',
                'Status': '❌ No TEC-19 Evidence Found',
                'Confidence': f'{best_score:.0f}%' if best_score > 0 else '0%',
            })
            continue

        tec_idx, tec_match = best_tec
        matched_tec_ids.add(tec_idx)

        # ── Compare dates ──
        tec_date = tec_match['date']
        # Full date match: same year AND same month
        date_match = (tec_date and pms_date and
                      tec_date.year == pms_date.year and
                      tec_date.month == pms_date.month)
        # Year-only match: same year but different month
        year_match = (tec_date and pms_date and tec_date.year == pms_date.year)

        # ── Physics check ──
        max_hrs, is_curr_yr = calculate_max_possible_hours(pms_row.to_dict(), audit_date)
        phys_status, phys_msg = classify_physics_violation(current_hrs, max_hrs, interval_hrs, is_curr_yr)

        # ── Build result record ──
        record = {
            'Code': code,
            'Component': component,
            'PMS Date': pms_date.strftime('%d-%m-%Y') if pms_date else '—',
            'TEC Date': tec_date.strftime('%d-%m-%Y') if tec_date else '—',
            'TEC Evidence': tec_match['text'][:120],
            'Match Confidence': f'{best_score:.0f}%',
            'Current Hrs': f"{current_hrs:.1f}" if current_hrs is not None else '—',
            'Max Hrs Since OH': f"{max_hrs:.1f}" if max_hrs != float('inf') else 'N/A (prior yr)',
            'Interval Hrs': f"{interval_hrs:.0f}" if interval_hrs else '—',
            'OH Year': 'Current' if is_curr_yr else 'Prior',
            'Physics Check': phys_msg,
            'Source File': tec_match['source'],
        }

        # ── Classify ──
        if phys_status == 'PHYSICS_VIOLATION':
            record['Status'] = '🔴 PHYSICS VIOLATION — Counter Not Reset'
            results['physics_violations'].append(record)
        elif phys_status == 'OVERDUE':
            record['Status'] = '🟠 OVERDUE — Exceeds Service Interval'
            record['Overdue Detail'] = phys_msg
            results['physics_violations'].append(record)
        elif not year_match:
            # TEC evidence found but belongs to a different year entirely
            # → treat as Ghost (no corroborating evidence for the claimed year)
            record['Status'] = '👻 GHOST — No TEC Evidence for Claimed Year'
            record['Note'] = (f"PMS claims {pms_date.strftime('%b %Y')} but "
                              f"best TEC match is from {tec_date.strftime('%b %Y') if tec_date else '?'} "
                              f"(different year — upload the corresponding diary)")
            results['ghost_overhauls'].append(record)
        elif not date_match:
            # Same year, different month → genuine date discrepancy
            record['Conflict'] = (f"PMS claims {pms_date.strftime('%b %Y')}, "
                                  f"TEC diary says {tec_date.strftime('%b %Y') if tec_date else 'unknown'}")
            record['Status'] = '⚠️ DATE CONFLICT — Same Year, Different Month'
            results['date_conflicts'].append(record)
        elif current_hrs is None or current_hrs == 0.0:
            record['Status'] = '🟡 INCOMPLETE — Hours Not Recorded in PMS'
            results['missing_hours'].append(record)
        else:
            record['Status'] = '✅ VERIFIED SYNC'
            results['verified_syncs'].append(record)

    # ── Find TEC entries with NO PMS match ──
    for i, tec_entry in enumerate(tec_entries):
        if i in matched_tec_ids:
            continue
        results['unlogged_jobs'].append({
            'TEC Description': tec_entry['text'][:150],
            'TEC Date': tec_entry['date'].strftime('%d-%m-%Y') if tec_entry['date'] else '—',
            'Source': tec_entry['source'],
            'Status': '📋 JOB PERFORMED — Not Found in PMS Ledger',
        })

    return results

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def generate_excel_report(results: dict, audit_date: datetime,
                          vessel_meta: dict) -> bytes:
    """Generate a color-coded Excel audit report."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book

        # Formats
        hdr_fmt = wb.add_format({
            'bold': True, 'bg_color': '#1a1a2e', 'font_color': '#FFFFFF',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        red_fmt = wb.add_format({'bg_color': '#FFE4E1', 'border': 1, 'text_wrap': True})
        orange_fmt = wb.add_format({'bg_color': '#FFF3CD', 'border': 1, 'text_wrap': True})
        green_fmt = wb.add_format({'bg_color': '#E8F5E9', 'border': 1, 'text_wrap': True})
        yellow_fmt = wb.add_format({'bg_color': '#FFFDE7', 'border': 1, 'text_wrap': True})
        normal_fmt = wb.add_format({'border': 1, 'text_wrap': True})

        # ── Summary Sheet ──
        summary_data = {
            'Audit Metric': [
                'Vessel', 'Audit Date', 'Total PMS Components Analyzed',
                '✅ Verified Syncs', '🔴 Physics Violations',
                '⚠️ Date Conflicts', '🟡 Incomplete Records',
                '👻 Ghost Overhauls', '📋 Unlogged Jobs'
            ],
            'Count / Value': [
                vessel_meta.get('vessel', 'MV ALEXIS'),
                audit_date.strftime('%d-%m-%Y'),
                str(len(results['verified_syncs']) + len(results['physics_violations']) +
                    len(results['date_conflicts']) + len(results['ghost_overhauls']) +
                    len(results['missing_hours'])),
                str(len(results['verified_syncs'])),
                str(len(results['physics_violations'])),
                str(len(results['date_conflicts'])),
                str(len(results['missing_hours'])),
                str(len(results['ghost_overhauls'])),
                str(len(results['unlogged_jobs'])),
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

        # ── Write each tier ──
        sheets = [
            ('Physics Violations', results['physics_violations'], red_fmt),
            ('Date Conflicts', results['date_conflicts'], orange_fmt),
            ('Verified Syncs', results['verified_syncs'], green_fmt),
            ('Ghost Overhauls', results['ghost_overhauls'], yellow_fmt),
            ('Incomplete Records', results['missing_hours'], orange_fmt),
            ('Unlogged Jobs', results['unlogged_jobs'], yellow_fmt),
        ]

        for sheet_name, data, row_fmt in sheets:
            if not data:
                pd.DataFrame([{'Note': 'No items in this category.'}]).to_excel(
                    writer, sheet_name=sheet_name, index=False)
                continue

            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            ws = writer.sheets[sheet_name]

            # Write headers
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, hdr_fmt)

            # Apply row formatting
            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    ws.write(row_idx + 1, col_idx,
                             str(df.iloc[row_idx, col_idx]), row_fmt)

            # Auto-fit columns
            for col_idx, col_name in enumerate(df.columns):
                max_width = max(len(col_name), 15)
                ws.set_column(col_idx, col_idx, min(max_width, 40))

    return output.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

def apply_custom_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

    html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; }

    .main { background: #0d0d1a; }

    .metric-card {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 1px solid #2a2a4a;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        text-align: center;
        transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-2px); }

    .metric-value { font-size: 2.5rem; font-weight: 700; font-family: 'JetBrains Mono'; }
    .metric-label { font-size: 0.8rem; color: #8888aa; text-transform: uppercase; letter-spacing: 0.1em; margin-top: 0.25rem; }

    .tier-header {
        font-family: 'Space Grotesk'; font-weight: 700; font-size: 1rem;
        text-transform: uppercase; letter-spacing: 0.15em;
        padding: 0.5rem 1rem; border-radius: 6px; margin-bottom: 1rem;
        display: inline-block;
    }
    .tier-red { background: #3d1a1a; color: #ff6b6b; border: 1px solid #ff6b6b40; }
    .tier-orange { background: #3d2a1a; color: #ffa06b; border: 1px solid #ffa06b40; }
    .tier-green { background: #1a3d1a; color: #6bffa0; border: 1px solid #6bffa040; }
    .tier-yellow { background: #3d3d1a; color: #ffe06b; border: 1px solid #ffe06b40; }
    .tier-blue { background: #1a2a3d; color: #6bb5ff; border: 1px solid #6bb5ff40; }

    .stDataFrame { border-radius: 8px; overflow: hidden; }
    div[data-testid="stSidebar"] { background: #0d0d1a; border-right: 1px solid #1a1a2e; }
    .stButton>button {
        background: linear-gradient(135deg, #6c63ff, #5a54d6);
        color: white; border: none; border-radius: 8px;
        font-family: 'Space Grotesk'; font-weight: 600;
        letter-spacing: 0.05em; transition: all 0.2s;
    }
    .stButton>button:hover { transform: translateY(-1px); box-shadow: 0 4px 20px #6c63ff40; }

    .pipeline-title {
        font-family: 'JetBrains Mono'; font-size: 2.5rem; font-weight: 700;
        background: linear-gradient(135deg, #6c63ff, #a56cff, #ff6ca5);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        letter-spacing: -0.02em;
    }
    .pipeline-sub {
        color: #6666aa; font-size: 0.9rem; letter-spacing: 0.1em;
        text-transform: uppercase; margin-top: -0.5rem;
    }
    .anchor-line {
        height: 2px;
        background: linear-gradient(90deg, #6c63ff, #ff6ca5, transparent);
        margin: 1rem 0 2rem 0;
    }
    </style>
    """, unsafe_allow_html=True)

def render_metric_card(value, label, color="#6c63ff"):
    return f"""
    <div class="metric-card">
        <div class="metric-value" style="color:{color}">{value}</div>
        <div class="metric-label">{label}</div>
    </div>
    """

def main():
    apply_custom_css()

    if 'audit_results' not in st.session_state:
        st.session_state.audit_results = None
    if 'audit_meta' not in st.session_state:
        st.session_state.audit_meta = {}
    if 'pms_df' not in st.session_state:
        st.session_state.pms_df = None
    if 'tec_entries' not in st.session_state:
        st.session_state.tec_entries = None

    # ── SIDEBAR ──
    with st.sidebar:
        st.markdown("""
        <div class="pipeline-title">⚓ TEMPORAL<br>PIPELINE</div>
        <div class="pipeline-sub">Zero-Trust Audit Engine v2.0</div>
        <div class="anchor-line"></div>
        """, unsafe_allow_html=True)

        st.markdown("**VESSEL CONFIGURATION**")
        vessel_name = st.text_input("Vessel Name", value="MV ALEXIS", label_visibility="collapsed",
                                     placeholder="Vessel Name")

        st.markdown("**AUDIT DATE**")
        audit_date_input = st.date_input(
            "Select audit date",
            value=date.today(),
            label_visibility="collapsed"
        )

        st.markdown("**MATCH SENSITIVITY**")
        match_threshold = st.slider(
            "Fuzzy match threshold (%)",
            min_value=40, max_value=90, value=58, step=2,
            help="Lower = more matches (may include false positives). Higher = stricter matching."
        )

        st.divider()

        st.markdown("**① UPLOAD PMS LEDGER**")
        pms_file = st.file_uploader(
            "PMS Excel file",
            type=['xlsx', 'xls'],
            label_visibility="collapsed",
            key="pms_upload"
        )

        st.markdown("**② UPLOAD TEC-19 DIARY**")
        tec_files = st.file_uploader(
            "TEC-19 Word file(s)",
            type=['doc', 'docx'],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="tec_upload"
        )

        st.divider()

        run_clicked = st.button(
            "▶  EXECUTE AUDIT",
            type="primary",
            use_container_width=True,
            disabled=(pms_file is None or not tec_files)
        )

        if pms_file is None or not tec_files:
            st.caption("Upload both files to enable the audit.")

    # ── MAIN CONTENT ──
    st.markdown("""
    <div class="pipeline-title" style="font-size:2rem">RECONCILIATION DASHBOARD</div>
    <div class="pipeline-sub">Cross-referencing PMS Ledger ↔ TEC-19 Maintenance Diary</div>
    <div class="anchor-line"></div>
    """, unsafe_allow_html=True)

    # ── RUN AUDIT ──
    if run_clicked:
        audit_dt = datetime.combine(audit_date_input, datetime.min.time())

        with st.status("🔬 Running Zero-Trust Audit Pipeline...", expanded=True) as status:
            # Step 1: Parse PMS
            st.write("📊 Parsing PMS Ledger (Excel)...")
            pms_df, pms_error, pms_meta = parse_pms_excel(pms_file.getvalue())
            if pms_error:
                st.error(f"PMS Parse Error: {pms_error}")
                status.update(label="Audit failed.", state="error")
                st.stop()

            st.write(f"✅ PMS: {len(pms_df)} components extracted from sheet '{pms_meta.get('sheet_name')}'")

            # Step 2: Parse TEC-19
            st.write(f"📝 Parsing {len(tec_files)} TEC-19 diary file(s)...")
            tec_entries = parse_tec19_files(tec_files)

            if not tec_entries:
                st.error("No maintenance action entries found in the TEC-19 diary files. "
                         "Check that the files contain work descriptions with dates.")
                status.update(label="Audit failed.", state="error")
                st.stop()

            st.write(f"✅ TEC-19: {len(tec_entries)} maintenance action entries extracted")

            # Step 3: Run audit
            st.write("🔗 Running fuzzy cross-reference and physics validation...")
            results = run_audit(pms_df, tec_entries, audit_dt, match_threshold)

            total = (len(results['verified_syncs']) + len(results['physics_violations']) +
                     len(results['date_conflicts']) + len(results['ghost_overhauls']) +
                     len(results['missing_hours']))

            st.session_state.audit_results = results
            st.session_state.audit_meta = {
                'vessel': vessel_name,
                'audit_date': audit_dt,
                'pms_meta': pms_meta,
                'tec_count': len(tec_entries),
                'pms_count': len(pms_df),
            }
            st.session_state.pms_df = pms_df
            st.session_state.tec_entries = tec_entries

            status.update(label=f"✅ Audit Complete — {total} components processed.", state="complete")

    # ── DISPLAY RESULTS ──
    if st.session_state.audit_results:
        res = st.session_state.audit_results
        meta = st.session_state.audit_meta
        audit_dt = meta.get('audit_date', datetime.now())

        # ── Metrics Row ──
        total = (len(res['verified_syncs']) + len(res['physics_violations']) +
                 len(res['date_conflicts']) + len(res['ghost_overhauls']) +
                 len(res['missing_hours']))

        cols = st.columns(6)
        metrics = [
            (total, "Total Analyzed", "#6c63ff"),
            (len(res['verified_syncs']), "Verified Syncs", "#4caf50"),
            (len(res['physics_violations']), "Physics Violations", "#f44336"),
            (len(res['date_conflicts']), "Date Conflicts", "#ff9800"),
            (len(res['ghost_overhauls']), "Ghost Overhauls", "#9c27b0"),
            (len(res['missing_hours']), "Missing Hours", "#ffc107"),
        ]

        for col, (val, label, color) in zip(cols, metrics):
            with col:
                st.markdown(render_metric_card(val, label, color), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Export Button ──
        report_bytes = generate_excel_report(res, audit_dt, meta)
        st.download_button(
            label="📥 Download Full Audit Report (Excel)",
            data=report_bytes,
            file_name=f"audit_report_{meta.get('vessel', 'vessel').replace(' ', '_')}_{audit_dt.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary"
        )

        st.divider()

        # ── Tabbed Results ──
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            f"🔴 Physics Violations ({len(res['physics_violations'])})",
            f"⚠️ Date Conflicts ({len(res['date_conflicts'])})",
            f"✅ Verified Syncs ({len(res['verified_syncs'])})",
            f"👻 Ghost Overhauls ({len(res['ghost_overhauls'])})",
            f"🟡 Missing Hours ({len(res['missing_hours'])})",
            f"📋 Unlogged Jobs ({len(res['unlogged_jobs'])})",
        ])

        with tab1:
            st.markdown('<div class="tier-header tier-red">🔴 PHYSICS VIOLATIONS — Critical Risk</div>', unsafe_allow_html=True)
            st.caption("These components show running hours that EXCEED what was physically possible since the reported overhaul date. "
                       "This proves the hour counter was never reset, or the overhaul date is fabricated.")
            if res['physics_violations']:
                st.dataframe(pd.DataFrame(res['physics_violations']), use_container_width=True, hide_index=True)
            else:
                st.success("✅ No physics violations detected.")

        with tab2:
            st.markdown('<div class="tier-header tier-orange">⚠️ DATE CONFLICTS — Requires Investigation</div>', unsafe_allow_html=True)
            st.caption("A matching job was found in the TEC-19 diary, but the completion date differs from what the PMS ledger claims.")
            if res['date_conflicts']:
                st.dataframe(pd.DataFrame(res['date_conflicts']), use_container_width=True, hide_index=True)
            else:
                st.success("✅ No date conflicts detected.")

        with tab3:
            st.markdown('<div class="tier-header tier-green">✅ VERIFIED SYNCS — Clean Records</div>', unsafe_allow_html=True)
            st.caption("Both the date and hours match between PMS and TEC-19 diary, and the physics check passes.")
            if res['verified_syncs']:
                st.dataframe(pd.DataFrame(res['verified_syncs']), use_container_width=True, hide_index=True)
            else:
                st.info("No fully verified syncs yet. Adjust the match threshold if needed.")

        with tab4:
            st.markdown('<div class="tier-header tier-yellow">👻 GHOST OVERHAULS — No TEC Evidence</div>', unsafe_allow_html=True)
            st.caption("The PMS ledger claims an overhaul was performed, but no corresponding entry was found in ANY TEC-19 diary. "
                       "This may indicate phantom record entries.")
            if res['ghost_overhauls']:
                st.dataframe(pd.DataFrame(res['ghost_overhauls']), use_container_width=True, hide_index=True)
            else:
                st.success("✅ No ghost overhauls detected.")

        with tab5:
            st.markdown('<div class="tier-header tier-orange">🟡 INCOMPLETE RECORDS — Admin Deficiency</div>', unsafe_allow_html=True)
            st.caption("A matching TEC-19 entry confirms the work was done, but the running hours were not properly recorded in the PMS.")
            if res['missing_hours']:
                st.dataframe(pd.DataFrame(res['missing_hours']), use_container_width=True, hide_index=True)
            else:
                st.success("✅ No incomplete records detected.")

        with tab6:
            st.markdown('<div class="tier-header tier-blue">📋 UNLOGGED JOBS — Work Without PMS Entry</div>', unsafe_allow_html=True)
            st.caption("These jobs were recorded in the TEC-19 diary but have NO matching component in the PMS ledger. "
                       "This may indicate unplanned maintenance or components not yet added to the PMS.")
            if res['unlogged_jobs']:
                st.dataframe(pd.DataFrame(res['unlogged_jobs']), use_container_width=True, hide_index=True)
            else:
                st.success("✅ All diary entries are accounted for in the PMS.")

        # ── Debug Info ──
        with st.expander("🔧 Audit Diagnostics & Column Mapping", expanded=False):
            pms_meta = meta.get('pms_meta', {})
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**PMS Column Mapping**")
                st.json(pms_meta.get('col_map', {}))
            with col2:
                st.markdown("**Monthly Hours Columns**")
                st.json(pms_meta.get('monthly_cols', {}))

            st.markdown(f"**PMS Components Parsed:** {meta.get('pms_count', 0)}")
            st.markdown(f"**TEC-19 Action Entries Extracted:** {meta.get('tec_count', 0)}")
            st.markdown(f"**Fuzzy Match Threshold Used:** {match_threshold}%")

    else:
        # ── Welcome State ──
        st.markdown("""
        <div style="text-align:center; padding: 4rem 2rem; color: #4444aa;">
            <div style="font-size:4rem; margin-bottom:1rem">⚓</div>
            <div style="font-size:1.3rem; font-family: 'JetBrains Mono'; color:#6666cc;">
                AWAITING UPLINK
            </div>
            <div style="font-size:0.9rem; margin-top:0.5rem; color:#444477; max-width:500px; margin-left:auto; margin-right:auto;">
                Upload the PMS Excel ledger and TEC-19 diary files in the sidebar,
                then press EXECUTE AUDIT to begin the zero-trust cross-reference.
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Info cards
        c1, c2, c3 = st.columns(3)
        with c1:
            st.info("**① Upload PMS Ledger**\n\nThe Excel spreadsheet containing component codes, overhaul dates, and running hours.")
        with c2:
            st.info("**② Upload TEC-19 Diary**\n\nThe Word document(s) where the Chief Engineer logged completed maintenance work.")
        with c3:
            st.info("**③ Set Audit Date**\n\nChoose the inspection date. The engine uses this to calculate maximum possible elapsed hours since each overhaul.")

if __name__ == "__main__":
    main()
