import base64
import json
import datetime
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
    font-weight: 600; font-size: 1.5rem;
    letter-spacing: .2em; color: #1a1a1a; margin-bottom: .15rem;
}
.form-header p { font-size: .73rem; color: #aaa; letter-spacing: .08em; }

/* ── panel ── */
.panel {
    background: #ffffff; border: 1px solid #e2ddd7;
    padding: 0 1.3rem 1rem; margin-bottom: .75rem;
}
.panel-title {
    font-size: .64rem; letter-spacing: .16em; color: #a09890;
    margin: 0 -1.3rem .9rem; padding: .55rem 1.3rem .45rem;
    border-bottom: 1px solid #ece7e0;
    background: #ffffff;
}

/* ── remove top ghost block ── */
[data-testid=\"stMainBlockContainer\"],
.block-container,
[data-testid=\"stAppViewBlockContainer\"] {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding-top: 0 !important;
}
.main > div { padding-top: 0 !important; }

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

/* ── sub-section label ── */
.sub-label {
    font-size: .68rem; letter-spacing: .1em; color: #b0a89e;
    margin: .6rem 0 .2rem; display: block;
}

/* ── time summary box ── */
.time-box {
    background: #f5f2ed; border: 1px solid #ddd8d0;
    padding: .5rem .9rem;
    display: flex; align-items: center; justify-content: center;
    gap: .5rem; margin: .5rem 0 .3rem;
}
.time-box .time-val { font-size: .86rem; color: #333; }
.time-box .days-val {
    font-size: .72rem; color: #aaa; 
    border-left: 1px solid #d0c8c0; padding-left: .5rem; margin-left: .2rem;
}

/* ═════════════════════════════════════════
   RADIO CHIPS
   ═════════════════════════════════════════ */
[data-testid="stRadio"] > label { display: none !important; }
[data-testid="stRadio"] > div {
    display: flex !important; flex-wrap: wrap !important;
    gap: .4rem !important; align-items: center !important;
}
[data-testid="stRadio"] label[data-baseweb="radio"] {
    display: inline-flex !important;
    align-items: center !important; justify-content: center !important;
    background: #f5f2ed !important; border: 1px solid #d0c8c0 !important;
    border-radius: 2px !important; padding: .36rem .9rem !important;
    cursor: pointer !important;
    transition: background .15s, border-color .15s !important;
    min-width: 62px !important;
}
[data-testid="stRadio"] label[data-baseweb="radio"]:hover {
    background: #eae6df !important; border-color: #b0a89e !important;
}
/* Hide circle indicator */
[data-testid="stRadio"] label[data-baseweb="radio"] > div:first-child {
    display: none !important;
}
/* Unselected text */
[data-testid="stRadio"] label[data-baseweb="radio"] > div:last-child,
[data-testid="stRadio"] label[data-baseweb="radio"] > div:last-child *,
[data-testid="stRadio"] label[data-baseweb="radio"] span,
[data-testid="stRadio"] label[data-baseweb="radio"] p {
    color: #3a3a3a !important; font-size: .83rem !important;
    line-height: 1 !important; text-align: center !important;
    margin: 0 !important; padding: 0 !important;
}
/* Selected: black background */
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) {
    background: #1a1a1a !important; border-color: #1a1a1a !important;
}
/* Selected: force WHITE on every child */
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked),
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) *,
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) > div,
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) > div *,
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) > div:last-child,
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) > div:last-child p,
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) > div:last-child span,
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) [data-testid="stMarkdownContainer"],
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) [data-testid="stMarkdownContainer"] p {
    color: #f5f0e8 !important;
}

/* ── generate button ── */
.stButton > button {
    width: 100%; border-radius: 0 !important;
    background-color: #1a1a1a !important; color: #f7f5f0 !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: .78rem !important; letter-spacing: .14em !important;
    padding: .7rem 1rem !important; border: none !important;
    transition: background-color .18s;
}
.stButton > button:hover { background-color: #3c3c3c !important; }
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

for k, v in [("xlsx_data", None), ("xlsx_fname", None), ("download_id", 0)]:
    if k not in st.session_state:
        st.session_state[k] = v

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
    leave_type = st.radio("假別", options=["事假", "病假", "特休", "產/婚/喪假", "其他"], horizontal=True)
    auto_reason = "特休" if leave_type == "特休" else ""
    reason = st.text_area("事由", value=auto_reason, height=68,
                          placeholder="請填寫請假事由 …", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with right:
    # ── 請假時間 ──────────────────────────────────────────────────────────────
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<p class="panel-title">請假時間　LEAVE PERIOD</p>', unsafe_allow_html=True)

    # 開始日期
    st.markdown('<span class="sub-label">開始日期　START DATE</span>', unsafe_allow_html=True)
    s1, s2, s3 = st.columns(3)
    with s1: start_year  = st.number_input("民國年", min_value=100, max_value=200, value=roc_now, step=1, key="sy")
    with s2: start_month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="sm")
    with s3: start_day   = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1, key="sd")

    # 結束日期
    st.markdown('<span class="sub-label">結束日期　END DATE</span>', unsafe_allow_html=True)
    e1, e2, e3 = st.columns(3)
    with e1: end_year  = st.number_input("民國年", min_value=100, max_value=200, value=roc_now, step=1, key="ey")
    with e2: end_month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="em")
    with e3: end_day   = st.number_input("日", min_value=1, max_value=31, value=today.day, step=1, key="ed")

    # 時段
    st.markdown('<span class="sub-label">時段　TIME SLOT</span>', unsafe_allow_html=True)
    time_preset = st.radio("時段", options=["整天", "半天上午", "半天下午"], horizontal=True, key="tp")

    TIME_MAP = {
        "整天":    (9, 30, 18, 30),
        "半天上午": (9,  0, 13,  0),
        "半天下午": (14, 0, 18,  0),
    }
    sh, sm_val, eh, em_val = TIME_MAP[time_preset]

    # 計算天數
    try:
        sd = date(int(start_year) + 1911, int(start_month), int(start_day))
        ed = date(int(end_year)   + 1911, int(end_month),   int(end_day))
        delta = (ed - sd).days + 1
        if time_preset == "整天":
            computed_days = str(delta) if delta > 1 else "1"
        else:
            computed_days = str(delta * 0.5) if delta > 1 else "0.5"
    except Exception:
        computed_days = "1"

    # 時間摘要顯示
    st.markdown(
        f'<div class="time-box">'
        f'<span class="time-val">{sh:02d}:{sm_val:02d}&ensp;—&ensp;{eh:02d}:{em_val:02d}</span>'
        f'<span class="days-val">{computed_days} 天</span>'
        f'</div>',
        unsafe_allow_html=True
    )

    # 可手動修改天數
    total_days_input = st.text_input(
        "合計天數（可手動修改）",
        value=computed_days,
        label_visibility="visible",
    )
    total_days = total_days_input.strip() or computed_days

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Generate ──────────────────────────────────────────────────────────────
    st.markdown("")
    if st.button("生成並下載請假單　Generate & Download"):
        if not applicant.strip():
            st.warning("請填寫請假人姓名。")
        else:
            xlsx_bytes = generate_leave_xlsx(
                apply_year=int(apply_year), apply_month=int(apply_month), apply_day=int(apply_day),
                applicant=applicant.strip(), proxy=proxy.strip(),
                leave_type=leave_type, reason=reason.strip(),
                start_year=int(start_year), start_month=int(start_month), start_day=int(start_day),
                start_hour=sh, start_minute=sm_val,
                end_year=int(end_year), end_month=int(end_month), end_day=int(end_day),
                end_hour=eh, end_minute=em_val,
                total_days=total_days,
            )
            st.session_state.xlsx_data  = xlsx_bytes
            st.session_state.xlsx_fname = (
                f"請假單_{applicant.strip()}_{start_year}年{start_month:02d}月{start_day:02d}日.xlsx"
            )
            st.session_state.download_id += 1

    # ── Auto-download ─────────────────────────────────────────────────────────
    if st.session_state.xlsx_data is not None:
        dl_id    = st.session_state.download_id
        dl_fname = st.session_state.xlsx_fname
        b64  = base64.b64encode(st.session_state.xlsx_data).decode()
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        components.html(f"""<!DOCTYPE html><html><body><script>
(function() {{
  var id = {dl_id};
  if (window.parent._dlId === id) return;
  window.parent._dlId = id;
  try {{
    var a = window.parent.document.createElement('a');
    a.href = 'data:{mime};base64,{b64}';
    a.download = {json.dumps(dl_fname)};
    window.parent.document.body.appendChild(a);
    a.click();
    setTimeout(function(){{ window.parent.document.body.removeChild(a); }}, 200);
  }} catch(e) {{ console.warn('download error', e); }}
}})();
</script></body></html>""", height=0)

        st.success(f"✓ {dl_fname} 已下載")
