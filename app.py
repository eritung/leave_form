import streamlit as st
from datetime import date
from generate_leave import generate_leave_xlsx

st.set_page_config(
    page_title="艾迪英特 請假申請",
    page_icon="📋",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@300;400;600&family=Noto+Sans+TC:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Sans TC', 'Hiragino Kaku Gothic ProN', sans-serif;
    color: #1a1a1a;
}
.main > div { padding-top: 1.5rem; }
.stApp { background-color: #f7f5f0; }

/* ── header ── */
.form-header {
    text-align: center;
    padding: 1.6rem 0 1.2rem;
    border-bottom: 1px solid #c8c0b4;
    margin-bottom: 1.6rem;
}
.form-header h1 {
    font-family: 'Noto Serif TC', serif;
    font-weight: 600;
    font-size: 1.55rem;
    letter-spacing: .18em;
    color: #1a1a1a;
    margin-bottom: .2rem;
}
.form-header p {
    font-size: .75rem;
    color: #999;
    letter-spacing: .06em;
}

/* ── section label ── */
.section-label {
    font-size: .68rem;
    letter-spacing: .14em;
    text-transform: uppercase;
    color: #9a8f82;
    margin: 1.2rem 0 .4rem;
    padding-bottom: .2rem;
    border-bottom: 1px solid #ddd8d0;
}

/* ── inputs ── */
.stTextInput label, .stTextArea label,
.stNumberInput label, .stSelectbox label, .stRadio label {
    font-size: .76rem !important;
    color: #666 !important;
    letter-spacing: .04em !important;
    font-weight: 400 !important;
}
.stTextInput input, .stNumberInput input {
    border-radius: 0 !important;
    border: none !important;
    border-bottom: 1px solid #b8b0a4 !important;
    background: transparent !important;
    font-size: .88rem !important;
    padding: .3rem .05rem !important;
    box-shadow: none !important;
}
.stTextInput input:focus, .stNumberInput input:focus {
    border-bottom-color: #1a1a1a !important;
    box-shadow: none !important;
}
.stTextArea textarea {
    border-radius: 0 !important;
    border: 1px solid #c8c0b4 !important;
    background: transparent !important;
    font-size: .88rem !important;
    box-shadow: none !important;
}
.stSelectbox > div > div {
    border-radius: 0 !important;
    border: none !important;
    border-bottom: 1px solid #b8b0a4 !important;
    background: transparent !important;
    box-shadow: none !important;
}
.stRadio > div { gap: .3rem !important; flex-wrap: wrap; }

/* ── card panel ── */
.panel {
    background: #fff;
    border: 1px solid #e0dbd4;
    padding: 1.2rem 1.4rem 1rem;
    margin-bottom: .8rem;
}
.panel-title {
    font-size: .68rem;
    letter-spacing: .14em;
    color: #9a8f82;
    margin-bottom: .8rem;
    padding-bottom: .25rem;
    border-bottom: 1px solid #ece8e2;
}

/* ── button ── */
.stButton > button {
    border-radius: 0 !important;
    background-color: #1a1a1a !important;
    color: #f7f5f0 !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: .78rem !important;
    letter-spacing: .12em !important;
    padding: .65rem 2.4rem !important;
    border: none !important;
    transition: background-color .2s;
}
.stButton > button:hover { background-color: #3d3d3d !important; }

.stDownloadButton > button {
    border-radius: 0 !important;
    background-color: transparent !important;
    color: #1a1a1a !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: .78rem !important;
    letter-spacing: .12em !important;
    padding: .65rem 2.4rem !important;
    border: 1px solid #1a1a1a !important;
    transition: all .2s;
    width: 100%;
}
.stDownloadButton > button:hover {
    background-color: #1a1a1a !important;
    color: #f7f5f0 !important;
}

/* ── divider ── */
.stDivider { border-color: #ddd8d0 !important; }

#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="form-header">
    <h1>艾迪英特股份有限公司　請假申請</h1>
    <p>Leave Request Form</p>
</div>
""", unsafe_allow_html=True)

today = date.today()
roc_now = today.year - 1911

# ── Session state init ───────────────────────────────────────────────────────
if "reason_val" not in st.session_state:
    st.session_state.reason_val = ""
if "generated" not in st.session_state:
    st.session_state.generated = None

# ── Layout: two main columns ─────────────────────────────────────────────────
left, right = st.columns([3, 2], gap="large")

with left:
    # ─ 申請日期 + 人員 ──────────────────────────────────────────────────────
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<p class="panel-title">申請資訊　APPLICATION INFO</p>', unsafe_allow_html=True)

    d1, d2, d3 = st.columns(3)
    with d1:
        apply_year  = st.number_input("申請日期　民國年", min_value=100, max_value=200, value=roc_now, step=1)
    with d2:
        apply_month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1)
    with d3:
        apply_day   = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1)

    p1, p2 = st.columns(2)
    with p1:
        applicant = st.text_input("請假人", value="董伊淇")
    with p2:
        proxy = st.text_input("代理人", value="葉詩宣")
    st.markdown('</div>', unsafe_allow_html=True)

    # ─ 假別 + 事由 ───────────────────────────────────────────────────────────
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<p class="panel-title">假別與事由　LEAVE TYPE & REASON</p>', unsafe_allow_html=True)

    leave_type = st.radio(
        "假別",
        options=["事假", "病假", "特休", "產/婚/喪假", "其他"],
        horizontal=True,
    )

    # Auto-fill reason when 特休 is selected
    auto_reason = "特休" if leave_type == "特休" else ""
    reason = st.text_area(
        "事由",
        value=auto_reason,
        height=72,
        placeholder="請填寫請假事由 …",
    )
    st.markdown('</div>', unsafe_allow_html=True)

with right:
    # ─ 請假時間 ──────────────────────────────────────────────────────────────
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<p class="panel-title">請假時間　LEAVE PERIOD</p>', unsafe_allow_html=True)

    st.markdown('<span style="font-size:.72rem;color:#9a8f82;letter-spacing:.1em">開始　START</span>', unsafe_allow_html=True)
    s1, s2, s3 = st.columns(3)
    with s1:
        s_year  = st.number_input("民國年", min_value=100, max_value=200, value=roc_now, step=1, key="sy")
    with s2:
        s_month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="sm")
    with s3:
        s_day   = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1, key="sd")

    s4, s5 = st.columns(2)
    with s4:
        s_hour   = st.number_input("時", min_value=0, max_value=23, value=9, step=1, key="sh")
    with s5:
        s_minute = st.selectbox("分", options=[0, 30], format_func=lambda x: f"{x:02d}", key="smi")

    st.markdown('<span style="font-size:.72rem;color:#9a8f82;letter-spacing:.1em;display:block;margin-top:.6rem">結束　END</span>', unsafe_allow_html=True)
    e1, e2, e3 = st.columns(3)
    with e1:
        e_year  = st.number_input("民國年", min_value=100, max_value=200, value=roc_now, step=1, key="ey")
    with e2:
        e_month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="em")
    with e3:
        e_day   = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1, key="ed")

    e4, e5 = st.columns(2)
    with e4:
        e_hour   = st.number_input("時", min_value=0, max_value=23, value=18, step=1, key="eh")
    with e5:
        e_minute = st.selectbox("分", options=[0, 30], format_func=lambda x: f"{x:02d}", key="emi")

    total_days = st.text_input("合計天數（例：1、0.5）", placeholder="1")
    st.markdown('</div>', unsafe_allow_html=True)

    # ─ Generate ──────────────────────────────────────────────────────────────
    st.markdown("")
    if st.button("生成請假單　Generate"):
        if not applicant.strip():
            st.warning("請填寫請假人姓名。")
        else:
            xlsx_bytes = generate_leave_xlsx(
                apply_year=int(apply_year), apply_month=int(apply_month), apply_day=int(apply_day),
                applicant=applicant.strip(), proxy=proxy.strip(),
                leave_type=leave_type, reason=reason.strip(),
                start_year=int(s_year), start_month=int(s_month), start_day=int(s_day),
                start_hour=int(s_hour), start_minute=int(s_minute),
                end_year=int(e_year), end_month=int(e_month), end_day=int(e_day),
                end_hour=int(e_hour), end_minute=int(e_minute),
                total_days=total_days.strip(),
            )
            st.session_state.generated = (xlsx_bytes, applicant.strip(), apply_year, apply_month, apply_day)

    if st.session_state.generated:
        xlsx_bytes, name, yr, mo, dy = st.session_state.generated
        fname = f"請假單_{name}_{yr}年{mo:02d}月{dy:02d}日.xlsx"
        st.success("✓ 請假單已生成")
        st.download_button(
            label="下載請假單　Download xlsx",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
