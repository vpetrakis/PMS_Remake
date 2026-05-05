import streamlit as st
import pandas as pd
from datetime import datetime
from temporal_audit.ingest import parse_running_hours_excel, parse_tec_files
from temporal_audit.matching import build_match_table
from temporal_audit.audit import run_integrity_audit, summarize_results
from temporal_audit.exports import to_excel_bytes

st.set_page_config(page_title='Temporal Integrity Audit', layout='wide', initial_sidebar_state='expanded')
st.title('Temporal Integrity Audit')
st.caption('Data-integrity-first reconciliation between PMS running hours and TEC schedule / job evidence.')

with st.sidebar:
    st.header('Inputs')
    audit_date = st.date_input('Audit date', value=datetime.today())
    pms_file = st.file_uploader('PMS / Running Hours (.xlsx, .xls, .doc, .docx)', type=['xlsx', 'xls', 'doc', 'docx'])
    tec_files = st.file_uploader('TEC logs (.doc, .docx)', type=['doc', 'docx'], accept_multiple_files=True)
    execute = st.button('Run audit', type='primary', use_container_width=True)

if execute:
    if not pms_file or not tec_files:
        st.error('Upload both PMS and TEC files.')
        st.stop()
    with st.spinner('Parsing, normalizing, matching, and auditing...'):
        pms_result = parse_running_hours_excel(pms_file)
        tec_result = parse_tec_files(tec_files)
        if pms_result.errors:
            st.error('PMS parsing errors detected.')
            st.dataframe(pd.DataFrame({'error': pms_result.errors}), use_container_width=True)
            st.stop()
        match_df = build_match_table(tec_result.records, pms_result.records)
        audit_bundle = run_integrity_audit(tec_df=tec_result.records, pms_df=pms_result.records, match_df=match_df, audit_date=pd.Timestamp(audit_date))
        summary = summarize_results(audit_bundle)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric('PMS rows', summary['pms_rows'])
    c2.metric('TEC rows', summary['tec_rows'])
    c3.metric('Verified', summary['verified_matches'])
    c4.metric('Exceptions', summary['exceptions'])
    c5.metric('Review', summary['review_queue'])
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(['Overview', 'Verified', 'Matches', 'Exceptions', 'Review', 'Exports'])
    with tab1:
        st.dataframe(pd.DataFrame([summary]), use_container_width=True)
        if tec_result.errors:
            st.warning('TEC parser reported issues.')
            st.dataframe(pd.DataFrame({'error': tec_result.errors}), use_container_width=True)
    with tab2:
        st.dataframe(audit_bundle['verified'], use_container_width=True)
    with tab3:
        st.dataframe(audit_bundle['matches'], use_container_width=True)
    with tab4:
        st.dataframe(audit_bundle['exceptions'], use_container_width=True)
    with tab5:
        st.dataframe(audit_bundle['review'], use_container_width=True)
        quarantine = pd.concat([
            pms_result.quarantine.assign(source='PMS') if not pms_result.quarantine.empty else pd.DataFrame(),
            tec_result.quarantine.assign(source='TEC') if not tec_result.quarantine.empty else pd.DataFrame()
        ], ignore_index=True)
        if not quarantine.empty:
            st.markdown('### Parser quarantine')
            st.dataframe(quarantine, use_container_width=True)
    with tab6:
        excel_bytes = to_excel_bytes(audit_bundle, summary)
        st.download_button('Download audit workbook', data=excel_bytes, file_name='temporal_integrity_audit.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
else:
    st.info('Upload PMS and TEC files, then run the audit.')
