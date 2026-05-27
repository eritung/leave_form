import streamlit as st
from datetime import date
from generate_leave import generate_leave_xlsx

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="艾迪英特 請假申請",
    page_icon="📋",
    layout="centered",
)

# ── Custom CSS — Japanese minimalist ────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@300;400;600&family=Noto+Sans+TC:wght@300;400;500&display=swap');

/* ---- base ---- */
html, body, [class*="css"] {
    font-family: 'Noto Sans TC', 'Hiragino Kaku Gothic ProN', sans-serif;
    color: #1a1a1a;
}

.main > div { padding-top: 2rem; }

/* ---- page background ---- */
.stApp {
    background-color: #f7f5f0;
}

/* ---- header ---- */
.form-header {
    text-align: center;
    padding: 2.4rem 0 1.6rem;
    border-bottom: 1px solid #c8c0b4;
    margin-bottom: 2rem;
}
.form-header h1 {
    font-family: 'Noto Serif TC', serif;
    font-weight: 600;
    font-size: 1.7rem;
    letter-spacing: .18em;
    color: #1a1a1a;
    margin-bottom: .3rem;
}
.form-header p {
    font-size: .78rem;
    color: #888;
    letter-spacing: .06em;
}

/* ---- section labels ---- */
.section-label {
    font-size: .7rem;
    letter-spacing: .14em;
    text-transform: uppercase;
    color: #9a8f82;
    margin: 1.8rem 0 .5rem;
    padding-bottom: .25rem;
    border-bottom: 1px solid #ddd8d0;
}

/* ---- card / field container ---- */
.stSelectbox label,
.stTextInput label,
.stTextArea label,
.stNumberInput label,
.stRadio label { 
    font-size: .78rem !important;
    color: #555 !important;
    letter-spacing: .04em !important;
    font-weight: 400 !important;
}

/* inputs */
.stTextInput input,
.stTextArea textarea,
.stNumberInput input {
    border-radius: 0 !important;
    border: none !important;
    border-bottom: 1px solid #b8b0a4 !important;
    background: transparent !important;
    font-size: .9rem !important;
    padding: .35rem .1rem !important;
    box-shadow: none !important;
    transition: border-color .2s;
}
.stTextInput input:focus,
.stTextArea textarea:focus,
.stNumberInput input:focus {
    border-bottom-color: #1a1a1a !important;
    box-shadow: none !important;
}

/* select */
.stSelectbox > div > div {
    border-radius: 0 !important;
    border: none !important;
    border-bottom: 1px solid #b8b0a4 !important;
    background: transparent !important;
    box-shadow: none !important;
}

/* radio */
.stRadio > div {
    gap: .4rem !important;
    flex-wrap: wrap;
}
.stRadio [data-testid="stMarkdownContainer"] p {
    font-size: .84rem !important;
}

/* ---- divider ---- */
hr { border: none; border-top: 1px solid #ddd8d0; margin: 1.6rem 0; }

/* ---- submit button ---- */
.stButton > button {
    width: 100%;
    border-radius: 0 !important;
    background-color: #1a1a1a !important;
    color: #f7f5f0 !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: .8rem !important;
    letter-spacing: .12em !important;
    padding: .75rem 2rem !important;
    border: none !important;
    margin-top: 1.4rem;
    transition: background-color .2s;
}
.stButton > button:hover {
    background-color: #3a3a3a !important;
}

/* ---- download button ---- */
.stDownloadButton > button {
    width: 100%;
    border-radius: 0 !important;
    background-color: transparent !important;
    color: #1a1a1a !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: .8rem !important;
    letter-spacing: .12em !important;
    padding: .75rem 2rem !important;
    border: 1px solid #1a1a1a !important;
    margin-top: .6rem;
    transition: all .2s;
}
.stDownloadButton > button:hover {
    background-color: #1a1a1a !important;
    color: #f7f5f0 !important;
}

/* ---- success box ---- */
.stSuccess {
    border-radius: 0 !important;
    border-left: 3px solid #1a1a1a !important;
    background-color: #f0ede8 !important;
}

/* hide streamlit branding */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="form-header">
    <h1>艾迪英特股份有限公司</h1>
    <p>請假申請單　Leave Request Form</p>
</div>
""", unsafe_allow_html=True)

# ── Helper — ROC year ─────────────────────────────────────────────────────
def roc_year(y: int) -> int:
    return y - 1911

today = date.today()
current_roc = roc_year(today.year)

# ── Form ─────────────────────────────────────────────────────────────────────
with st.form("leave_form"):

    # Section: 申請日期
    st.markdown('<p class="section-label">申請日期　Application Date</p>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        apply_year  = st.number_input("民國　年", min_value=100, max_value=200, value=current_roc, step=1)
    with c2:
        apply_month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1)
    with c3:
        apply_day   = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1)

    # Section: 人員
    st.markdown('<p class="section-label">人員資訊　Personnel</p>', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        applicant = st.text_input("請假人　Applicant")
    with col_b:
        proxy = st.text_input("代理人　Proxy")

    # Section: 假別
    st.markdown('<p class="section-label">假別　Leave Type</p>', unsafe_allow_html=True)
    leave_type = st.radio(
        "假別",
        options=["事假", "病假", "特休", "產/婚/喪假", "其他"],
        horizontal=True,
        label_visibility="collapsed",
    )

    # Section: 事由
    st.markdown('<p class="section-label">事由　Reason</p>', unsafe_allow_html=True)
    reason = st.text_area("請填寫請假事由", height=80, label_visibility="collapsed",
                          placeholder="請填寫請假事由 …")

    # Section: 請假時間
    st.markdown('<p class="section-label">請假起始　Start Time</p>', unsafe_allow_html=True)
    s1, s2, s3, s4, s5 = st.columns(5)
    with s1:
        s_year   = st.number_input("民國年", min_value=100, max_value=200, value=current_roc, step=1, key="sy")
    with s2:
        s_month  = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="sm")
    with s3:
        s_day    = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1, key="sd")
    with s4:
        s_hour   = st.number_input("時", min_value=0, max_value=23, value=9, step=1, key="sh")
    with s5:
        s_minute = st.selectbox("分", options=[0, 30], format_func=lambda x: f"{x:02d}", key="smi")

    st.markdown('<p class="section-label">請假結束　End Time</p>', unsafe_allow_html=True)
    e1, e2, e3, e4, e5 = st.columns(5)
    with e1:
        e_year   = st.number_input("民國年", min_value=100, max_value=200, value=current_roc, step=1, key="ey")
    with e2:
        e_month  = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="em")
    with e3:
        e_day    = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1, key="ed")
    with e4:
        e_hour   = st.number_input("時", min_value=0, max_value=23, value=18, step=1, key="eh")
    with e5:
        e_minute = st.selectbox("分", options=[0, 30], format_func=lambda x: f"{x:02d}", key="emi")

    st.markdown('<p class="section-label">合計天數　Total Days</p>', unsafe_allow_html=True)
    total_days = st.text_input("天數（例：1、0.5）", placeholder="1", label_visibility="collapsed")

    st.markdown("---")

    submitted = st.form_submit_button("生成請假單　Generate")

# ── Output ───────────────────────────────────────────────────────────────────
if submitted:
    if not applicant.strip():
        st.warning("請填寫請假人姓名。")
    else:
        xlsx_bytes = generate_leave_xlsx(
            apply_year=int(apply_year),
            apply_month=int(apply_month),
            apply_day=int(apply_day),
            applicant=applicant.strip(),
            proxy=proxy.strip(),
            leave_type=leave_type,
            reason=reason.strip(),
            start_year=int(s_year),
            start_month=int(s_month),
            start_day=int(s_day),
            start_hour=int(s_hour),
            start_minute=int(s_minute),
            end_year=int(e_year),
            end_month=int(e_month),
            end_day=int(e_day),
            end_hour=int(e_hour),
            end_minute=int(e_minute),
            total_days=total_days.strip() if total_days.strip() else "",
        )

        filename = f"請假單_{applicant.strip()}_{apply_year}年{apply_month:02d}月{apply_day:02d}日.xlsx"

        st.success("✓ 請假單已生成，點擊下方按鈕下載。")
        st.download_button(
            label="下載請假單　Download xlsx",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
