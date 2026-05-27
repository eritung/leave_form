import base64
import streamlit as st
import streamlit.components.v1 as components
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
    padding: 1.4rem 0 1rem;
    border-bottom: 1px solid #c8c0b4;
    margin-bottom: 1.4rem;
}
.form-header h1 {
    font-family: 'Noto Serif TC', serif;
    font-weight: 600;
    font-size: 1.5rem;
    letter-spacing: .2em;
    color: #1a1a1a;
    margin-bottom: .15rem;
}
.form-header p { font-size: .73rem; color: #aaa; letter-spacing: .08em; }

/* ── panel card ── */
.panel {
    background: #ffffff;
    border: 1px solid #e2ddd7;
    padding: 1.1rem 1.3rem .9rem;
    margin-bottom: .75rem;
}
.panel-title {
    font-size: .64rem;
    letter-spacing: .16em;
    color: #a09890;
    margin-bottom: .7rem;
    padding-bottom: .2rem;
    border-bottom: 1px solid #ece7e0;
}

/* ── inputs ── */
.stTextInput label, .stTextArea label,
.stNumberInput label, .stSelectbox label {
    font-size: .74rem !important; color: #777 !important;
    letter-spacing: .04em !important; font-weight: 400 !important;
}
.stTextInput input, .stNumberInput input {
    border-radius: 0 !important; border: none !important;
    border-bottom: 1px solid #c4bdb5 !important;
    background: transparent !important; font-size: .88rem !important;
    padding: .28rem .05rem !important; box-shadow: none !important;
}
.stTextInput input:focus, .stNumberInput input:focus {
    border-bottom-color: #1a1a1a !important; box-shadow: none !important;
}
.stSelectbox > div > div {
    border-radius: 0 !important; border: none !important;
    border-bottom: 1px solid #c4bdb5 !important;
    background: transparent !important; box-shadow: none !important;
}

/* ── leave type chips ── */
.stRadio > label { display: none !important; }
.stRadio > div {
    display: flex !important; flex-wrap: wrap; gap: .4rem !important;
}
.stRadio > div > label {
    display: flex !important; align-items: center;
    background: #f5f2ed !important;
    border: 1px solid #d5cfc8 !important;
    border-radius: 2px !important;
    padding: .32rem .75rem !important;
    cursor: pointer !important;
    transition: all .15s !important;
    font-size: .82rem !important;
    color: #444 !important;
}
.stRadio > div > label:hover {
    background: #eae6df !important;
    border-color: #b0a89e !important;
}
.stRadio > div > label[data-checked="true"],
.stRadio > div > label:has(input:checked) {
    background: #1a1a1a !important;
    border-color: #1a1a1a !important;
    color: #f7f5f0 !important;
}
.stRadio > div > label > div:first-child { display: none !important; }

/* ── time preset chips ── */
.time-preset .stRadio > div > label { 
    padding: .28rem .65rem !important;
    font-size: .78rem !important;
}

/* ── generate button ── */
.stButton > button {
    width: 100%;
    border-radius: 0 !important;
    background-color: #1a1a1a !important; color: #f7f5f0 !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: .78rem !important; letter-spacing: .14em !important;
    padding: .7rem 1rem !important; border: none !important;
    transition: background-color .18s;
}
.stButton > button:hover { background-color: #3c3c3c !important; }

/* ── success ── */
.stSuccess { border-radius: 0 !important; border-left: 3px solid #1a1a1a !important; }

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

# ─────────────────────────────────────────────────────────────────────────────
left, right = st.columns([3, 2], gap="large")

with left:
    # ── 申請資訊 ──────────────────────────────────────────────────────────────
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<p class="panel-title">申請資訊　APPLICATION INFO</p>', unsafe_allow_html=True)

    d1, d2, d3 = st.columns(3)
    with d1: apply_year  = st.number_input("申請日期　民國年", min_value=100, max_value=200, value=roc_now, step=1)
    with d2: apply_month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1)
    with d3: apply_day   = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1)

    p1, p2 = st.columns(2)
    with p1: applicant = st.text_input("請假人", value="董伊淇")
    with p2: proxy     = st.text_input("代理人", value="葉詩宣")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 假別 & 事由 ───────────────────────────────────────────────────────────
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<p class="panel-title">假別與事由　LEAVE TYPE & REASON</p>', unsafe_allow_html=True)

    leave_type = st.radio(
        "假別",
        options=["事假", "病假", "特休", "產/婚/喪假", "其他"],
        horizontal=True,
    )
    auto_reason = "特休" if leave_type == "特休" else ""
    reason = st.text_area("事由", value=auto_reason, height=68,
                          placeholder="請填寫請假事由 …", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with right:
    # ── 請假時間 ──────────────────────────────────────────────────────────────
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<p class="panel-title">請假時間　LEAVE PERIOD</p>', unsafe_allow_html=True)

    # Date row
    d1, d2, d3 = st.columns(3)
    with d1: leave_year  = st.number_input("民國年 ", min_value=100, max_value=200, value=roc_now, step=1, key="ly")
    with d2: leave_month = st.number_input("月 ", min_value=1, max_value=12, value=today.month, step=1, key="lm")
    with d3: leave_day   = st.number_input("日 ", min_value=1, max_value=31, value=today.day, step=1, key="ld")

    # Time preset
    st.markdown('<div class="time-preset">', unsafe_allow_html=True)
    time_preset = st.radio(
        "時段",
        options=["整天", "半天上午", "半天下午"],
        horizontal=True,
        key="tp",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Derive times from preset
    if time_preset == "整天":
        sh, sm_val, eh, em_val = 9, 30, 18, 30
        days_default = "1"
    elif time_preset == "半天上午":
        sh, sm_val, eh, em_val = 9, 0, 13, 0
        days_default = "0.5"
    else:  # 半天下午
        sh, sm_val, eh, em_val = 14, 0, 18, 0
        days_default = "0.5"

    # Display derived time (read-only look)
    st.markdown(
        f'<div style="background:#f5f2ed;border:1px solid #ddd8d0;padding:.55rem .8rem;'
        f'font-size:.84rem;color:#444;letter-spacing:.04em;margin:.4rem 0 .6rem;">'
        f'&nbsp;&nbsp;{sh:02d}:{sm_val:02d}&ensp;→&ensp;{eh:02d}:{em_val:02d}'
        f'&emsp;<span style="color:#999;font-size:.72rem">{days_default} 天</span></div>',
        unsafe_allow_html=True
    )

    total_days = days_default
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 生成按鈕 ──────────────────────────────────────────────────────────────
    st.markdown("")
    if st.button("生成並下載請假單　Generate & Download"):
        if not applicant.strip():
            st.warning("請填寫請假人姓名。")
        else:
            xlsx_bytes = generate_leave_xlsx(
                apply_year=int(apply_year), apply_month=int(apply_month), apply_day=int(apply_day),
                applicant=applicant.strip(), proxy=proxy.strip(),
                leave_type=leave_type, reason=reason.strip(),
                start_year=int(leave_year), start_month=int(leave_month), start_day=int(leave_day),
                start_hour=sh, start_minute=sm_val,
                end_year=int(leave_year), end_month=int(leave_month), end_day=int(leave_day),
                end_hour=eh, end_minute=em_val,
                total_days=total_days,
            )

            fname = f"請假單_{applicant.strip()}_{leave_year}年{leave_month:02d}月{leave_day:02d}日.xlsx"
            b64 = base64.b64encode(xlsx_bytes).decode()
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            # Auto-download via JS
            components.html(f"""
<script>
  (function() {{
    var a = document.createElement('a');
    a.href = 'data:{mime};base64,{b64}';
    a.download = '{fname}';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }})();
</script>
""", height=0)
            st.success(f"✓ {fname} 已下載")
