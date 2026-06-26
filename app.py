# -*- coding: utf-8 -*-
"""
가톨릭혈액병원 경영팀 일일현황 자동생성 웹앱
사용법: streamlit run app.py
"""

import streamlit as st
import tempfile, os, sys, traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from auto_report import (
    load_all_data, load_files_from_dir,
    calculate_table1, calculate_table2,
    create_clean_excel, find_template
)

# ─── 페이지 설정 ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="경영팀 일일현황 자동생성",
    page_icon="🏥",
    layout="wide",
)

# ─── CSS: 업로드 영역 화면 전체 수준으로 ────────────────────────────────────
st.markdown("""
<style>
.block-container { padding-top: 1.5rem !important; }

[data-testid="stFileUploader"] { width: 100%; }

[data-testid="stFileUploader"] section {
    min-height: 45vh !important;
    padding: 5vh 2rem !important;
    border: 4px dashed #4A90D9 !important;
    border-radius: 20px !important;
    background: linear-gradient(135deg, #f0f6ff 0%, #e8f0fe 100%) !important;
    text-align: center;
    cursor: pointer;
    transition: all 0.2s ease;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
}
[data-testid="stFileUploader"] section:hover {
    background: linear-gradient(135deg, #ddeeff 0%, #d5e8ff 100%) !important;
    border-color: #1a73e8 !important;
    box-shadow: 0 4px 20px rgba(74,144,217,0.25) !important;
}
[data-testid="stFileUploadDropzone"] svg {
    width: 80px !important;
    height: 80px !important;
}
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section span {
    font-size: 1.2rem !important;
}
[data-testid="stFileUploader"] section button {
    font-size: 1rem !important;
    padding: 0.5rem 2rem !important;
    margin-top: 0.5rem !important;
}
</style>
""", unsafe_allow_html=True)

# ─── 헤더 ────────────────────────────────────────────────────────────────────
st.title("🏥 경영팀 일일현황 자동생성")
st.markdown("**zip 파일** 또는 **개별 xlsx/xls 파일들**을 업로드하면 표1·표2 Excel이 자동으로 만들어집니다.")
st.divider()

# ─── 템플릿 상태 확인 ────────────────────────────────────────────────────────
template_path = find_template(SCRIPT_DIR)
if not template_path:
    st.error(
        "템플릿 Excel을 찾을 수 없습니다.  \n"
        "`교수님 자료/AI_agent_개발_우선순위 포함.xlsx` 파일이 같은 폴더에 있어야 합니다."
    )
    st.stop()

# ─── 업로드 영역 ─────────────────────────────────────────────────────────────
tab_zip, tab_files = st.tabs(["📦  zip 파일 업로드", "📄  개별 파일 업로드"])

with tab_zip:
    uploaded_zip = st.file_uploader(
        "zip 파일을 드래그하거나 클릭해서 선택하세요",
        type=["zip"],
        key="zip_uploader",
        label_visibility="visible",
    )

with tab_files:
    st.caption("교수님 자료 zip에서 **설명 파일 제외한 나머지 xlsx/xls 파일 전체**를 선택하세요")
    uploaded_files = st.file_uploader(
        "개별 파일들을 드래그하거나 클릭해서 선택하세요 (여러 파일 동시 선택 가능)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="files_uploader",
        label_visibility="visible",
    )

# 어떤 입력이 있는지 결정
has_zip   = uploaded_zip is not None
has_files = bool(uploaded_files)

if not has_zip and not has_files:
    st.markdown("""
    <div style="text-align:center; color:#888; margin-top:1rem; font-size:0.95rem;">
        <b>zip 파일</b> 탭 또는 <b>개별 파일 업로드</b> 탭을 선택하여 파일을 올려주세요
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─── 처리 ────────────────────────────────────────────────────────────────────
with st.spinner("데이터 처리 중..."):
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            import io as _io
            old_stdout = sys.stdout
            sys.stdout = _io.StringIO()
            try:
                if has_zip:
                    zip_path = os.path.join(tmpdir, uploaded_zip.name)
                    with open(zip_path, "wb") as f:
                        f.write(uploaded_zip.getvalue())
                    data = load_all_data(zip_path)
                else:
                    # 개별 파일 저장 후 로드
                    saved_paths = []
                    for uf in uploaded_files:
                        fpath = os.path.join(tmpdir, uf.name)
                        with open(fpath, "wb") as f:
                            f.write(uf.getvalue())
                        saved_paths.append(fpath)
                    data = load_files_from_dir(saved_paths)

                t1  = calculate_table1(data)
                t2  = calculate_table2(data)

                ds = data['report_date'].strftime('%Y%m%d')
                output_filename = "경영팀_일일현황_{}.xlsx".format(ds)
                output_path = os.path.join(tmpdir, output_filename)

                create_clean_excel(t1, t2, template_path, output_path)

                with open(output_path, "rb") as ef:
                    excel_bytes = ef.read()

                log_lines = sys.stdout.getvalue().splitlines()
            finally:
                sys.stdout = old_stdout

    except Exception as e:
        st.error("오류 발생: {}".format(e))
        st.code(traceback.format_exc())
        st.stop()

# ─── 결과 ────────────────────────────────────────────────────────────────────
s  = t1['seoul']
rd = data['report_date']

st.success("✅  {} 데이터 처리 완료!".format(rd.strftime('%Y년 %m월 %d일')))

# 핵심 수치
c1, c2, c3, c4 = st.columns(4)
c1.metric("전체 입원환자",    "{}명".format(s['total']))
c1.metric("전용병상 가동율",  "{:.1f}%".format(s['total_rate']))
c2.metric("이식병실",         "{}명".format(s['transplant']))
c2.metric("무균병실",         "{}명".format(s['sterile']))
c3.metric("응급실 재원",      "{}명".format(s['er']))
c3.metric("입원+응급 합계",   "{}명".format(s['total_er']))
c4.metric("전일입원",         "{}명".format(s['prev_adm']))
c4.metric("전일퇴원",         "{}명".format(s['prev_dis']))

st.divider()

# 표1 상세
with st.expander("표1 서울성모 상세", expanded=True):
    rows = [
        ("전체 입원환자 수",        s['total']),
        ("전용대비 가동율(%)",       s['total_rate']),
        ("전용병상 입원환자 수",     s['dedicated']),
        ("전용병상 가동율(%)",       s['dedicated_rate']),
        ("이식병실 입원환자 수",     s['transplant']),
        ("이식병실 가동율(%)",       s['transplant_rate']),
        ("무균병실 입원환자 수",     s['sterile']),
        ("무균병실 가동율(%)",       s['sterile_rate']),
        ("일반병실 입원환자 수",     s['general']),
        ("  전용 일반병실",          s['gen_ded']),
        ("  비전용 일반병실",        s['gen_non']),
        ("전용 일반병실 가동율(%)",  s['gen_rate']),
        ("응급실 재원",              s['er']),
        ("입원+응급",                s['total_er']),
        ("입원대기(외래)",           s['waiting']),
        ("전일입원",                 s['prev_adm']),
        ("전일퇴원",                 s['prev_dis']),
        ("BMT D-Day",                s['bmt']),
    ]
    import pandas as pd
    st.dataframe(pd.DataFrame(rows, columns=["항목", "서울성모"]),
                 use_container_width=True, hide_index=True)

# 표2 상세
with st.expander("표2 의사별 집계 ({}명)".format(len(t2))):
    t2_rows = [{"주치의": d,
                 "GW+ICU": v.get("GW+ICU", 0),
                "GW": v.get("GW", 0),
                "ICU": v.get("ICU", 0),
                "응급실재원": v.get("응급실재원", 0)}
               for d, v in sorted(t2.items())]
    if t2_rows:
        import pandas as pd
        st.dataframe(pd.DataFrame(t2_rows), use_container_width=True, hide_index=True)

with st.expander("처리 로그"):
    st.code("\n".join(log_lines), language=None)

st.divider()

st.download_button(
    label="📥  Excel 다운로드  ({})".format(output_filename),
    data=excel_bytes,
    file_name=output_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)
