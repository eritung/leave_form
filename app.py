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
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;500;600&family=Noto+Sans+TC:wght@400;500;600&display=swap');

:root {
    --bg:            #FAF9F6;
    --panel:         #FFFFFF;
    --border:        #E7E4DA;
    --border-soft:   #EFEDE4;
    --text:          #3D3929;
    --text-soft:     #83807A;
    --text-faint:    #ACA89F;
    --accent:        #C2673F;
    --accent-hover:  #AC5834;
    --accent-soft:   #F3E4D8;
    --radius-lg:      20px;
    --radius-md:      14px;
    --radius-pill:    999px;
}

html, body, [class*="css"] {
    font-family: 'Noto Sans TC', 'Hiragino Kaku Gothic ProN', sans-serif;
    color: var(--text);
}
.stApp { background-color: var(--bg); }

/* ── remove top ghost block ── */
[data-testid="stMainBlockContainer"],
.block-container,
[data-testid="stAppViewBlockContainer"] {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding-top: 1rem !important;
    max-width: 900px !important;
}
.main > div { padding-top: 0 !important; }

/* ── header ── */
.form-header {
    text-align: center;
    padding: .4rem 0 1.8rem;
    margin-bottom: .4rem;
}
.form-header .mark {
    display: inline-flex; align-items: center; justify-content: center;
    width: 38px; height: 38px; border-radius: var(--radius-pill);
    background: var(--accent); color: #fff; font-size: 1.1rem;
    margin-bottom: .9rem;
}
.form-header h1 {
    font-family: 'Noto Serif TC', serif;
    font-weight: 500; font-size: 1.55rem;
    letter-spacing: .01em; color: var(--text); margin-bottom: .3rem;
}
.form-header p { font-size: .82rem; color: var(--text-faint); letter-spacing: .03em; }

/* ── panel (chat-card style) ──
   Targets the wrapper div that st.container(key=...) generates,
   e.g. st.container(key="panel_apply") → .st-key-panel_apply           */
.st-key-panel_apply,
.st-key-panel_leave,
.st-key-panel_period {
    background: var(--panel);
    border: 1px solid var(--border-soft);
    border-radius: var(--radius-lg);
    padding: 1.1rem 1.4rem 1.3rem;
    box-shadow: 0 1px 2px rgba(61, 57, 41, 0.04);
}
.panel-title {
    font-size: .78rem; font-weight: 600; letter-spacing: .02em; color: var(--text);
    margin: 0 0 .9rem;
    padding: 0 0 .7rem;
    border-bottom: 1px solid var(--border-soft);
    display: flex; align-items: center; gap: .4rem;
}
.panel-title::before {
    content: ""; width: 6px; height: 6px; border-radius: 50%;
    background: var(--accent); display: inline-block;
}

/* ── date dropdowns (st.selectbox) ── */
div[data-baseweb="select"] > div {
    border-radius: var(--radius-md) !important;
    border: 1px solid var(--border) !important;
    background: var(--bg) !important;
    min-height: 2.5rem !important;
}
div[data-baseweb="select"] > div:hover {
    border-color: var(--accent) !important;
}
div[data-baseweb="popover"] li {
    font-size: .86rem !important;
}
div[data-baseweb="popover"] li:hover {
    background: var(--accent-soft) !important;
}

/* ── inputs ── */
.stTextInput label, .stTextArea label,
.stNumberInput label, .stSelectbox label {
    font-size: .78rem !important; color: var(--text-soft) !important;
    letter-spacing: .01em !important; font-weight: 500 !important;
}
.stTextInput input, .stNumberInput input, .stTextArea textarea {
    border-radius: var(--radius-md) !important;
    border: 1px solid var(--border) !important;
    background: var(--bg) !important; font-size: .9rem !important;
    padding: .5rem .75rem !important; box-shadow: none !important;
    color: var(--text) !important;
}
.stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {
    border-color: var(--accent) !important; box-shadow: 0 0 0 3px var(--accent-soft) !important;
    background: #ffffff !important;
}

/* ── sub-section label ── */
.sub-label {
    font-size: .74rem; font-weight: 500; letter-spacing: .01em; color: var(--text-soft);
    margin: .7rem 0 .35rem; display: block;
}

/* ── time summary box ── */
.time-box {
    background: var(--accent-soft); border: 1px solid #ecd8c9;
    border-radius: var(--radius-md);
    padding: .6rem 1rem;
    display: flex; align-items: center; justify-content: center;
    gap: .6rem; margin: .6rem 0 .4rem;
}
.time-box .time-val { font-size: .92rem; color: var(--text); font-weight: 500; }
.time-box .days-val {
    font-size: .78rem; color: var(--accent-hover); font-weight: 600;
    border-left: 1px solid #e0c3ac; padding-left: .6rem; margin-left: .2rem;
}

/* ═════════════════════════════════════════
   RADIO CHIPS (Claude suggestion-chip style)
   ═════════════════════════════════════════ */
[data-testid="stRadio"] > label { display: none !important; }
[data-testid="stRadio"] [role="radiogroup"],
[data-testid="stRadio"] > div {
    display: flex !important; flex-wrap: wrap !important;
    gap: .5rem !important; align-items: center !important;
}
[data-testid="stRadio"] label {
    display: inline-flex !important;
    align-items: center !important; justify-content: center !important;
    background: var(--bg) !important; border: 1px solid var(--border) !important;
    border-radius: var(--radius-pill) !important; padding: .42rem 1.05rem !important;
    cursor: pointer !important;
    transition: background .15s, border-color .15s, color .15s !important;
    min-width: 62px !important;
}
[data-testid="stRadio"] label:hover {
    background: var(--accent-soft) !important; border-color: #e0c3ac !important;
}
/* Hide circle indicator (always the first child of the label) */
[data-testid="stRadio"] label > div:first-child {
    display: none !important;
}
/* Text styling for whatever is left in the label */
[data-testid="stRadio"] label div,
[data-testid="stRadio"] label span,
[data-testid="stRadio"] label p {
    color: var(--text) !important; font-size: .84rem !important;
    line-height: 1 !important; text-align: center !important;
    margin: 0 !important; padding: 0 !important; font-weight: 500 !important;
}
/* Selected: coral background, white text on every descendant */
[data-testid="stRadio"] label:has(input:checked) {
    background: var(--accent) !important; border-color: var(--accent) !important;
}
[data-testid="stRadio"] label:has(input:checked) * {
    color: #ffffff !important;
}

/* ── generate button (Claude CTA style) ── */
.stButton > button {
    width: 100%; border-radius: var(--radius-lg) !important;
    background-color: var(--accent) !important; color: #ffffff !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: .88rem !important; font-weight: 600 !important; letter-spacing: .02em !important;
    padding: .75rem 1rem !important; border: none !important;
    transition: background-color .18s;
    box-shadow: 0 1px 2px rgba(194, 103, 63, 0.25);
}
.stButton > button:hover { background-color: var(--accent-hover) !important; }
div[data-testid="stButton"] { margin-top: .2rem; }
.stSuccess {
    border-radius: var(--radius-md) !important;
    border: 1px solid #ecd8c9 !important;
    background-color: var(--accent-soft) !important;
    color: var(--text) !important;
}
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="form-header">
    <div class="mark">✦</div>
    <h1>艾迪英特股份有限公司　請假申請</h1>
    <p>Leave Request Form</p>
</div>
""", unsafe_allow_html=True)

today = date.today()
roc_now = today.year - 1911

YEAR_OPTIONS  = list(range(roc_now - 2, roc_now + 3))
MONTH_OPTIONS = list(range(1, 13))
DAY_OPTIONS   = list(range(1, 32))

for k, v in [("xlsx_data", None), ("xlsx_fname", None), ("download_id", 0)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────────────────────────────────────
left, right = st.columns([3, 2], gap="large")

with left:
    # ── 申請資訊 ──────────────────────────────────────────────────────────────
    with st.container(key="panel_apply"):
        st.markdown('<p class="panel-title">申請資訊　APPLICATION INFO</p>', unsafe_allow_html=True)
        d1, d2, d3 = st.columns(3)
        with d1: apply_year  = st.selectbox("申請日期　民國年", options=YEAR_OPTIONS, index=YEAR_OPTIONS.index(roc_now))
        with d2: apply_month = st.selectbox("月", options=MONTH_OPTIONS, index=MONTH_OPTIONS.index(today.month))
        with d3: apply_day   = st.selectbox("日", options=DAY_OPTIONS, index=DAY_OPTIONS.index(today.day))
        p1, p2 = st.columns(2)
        with p1: applicant = st.text_input("請假人", value="董伊淇")
        with p2: proxy     = st.text_input("代理人", value="葉詩宣")

    # ── 假別 & 事由 ───────────────────────────────────────────────────────────
    with st.container(key="panel_leave"):
        st.markdown('<p class="panel-title">假別與事由　LEAVE TYPE & REASON</p>', unsafe_allow_html=True)
        leave_type = st.radio("假別", options=["事假", "病假", "特休", "產/婚/喪假", "其他"], horizontal=True)
        auto_reason = "特休" if leave_type == "特休" else ""
        reason = st.text_area("事由", value=auto_reason, height=150,
                              placeholder="請填寫請假事由 …", label_visibility="collapsed")

with right:
    # ── 請假時間 ──────────────────────────────────────────────────────────────
    with st.container(key="panel_period"):
        st.markdown('<p class="panel-title">請假時間　LEAVE PERIOD</p>', unsafe_allow_html=True)

        # 開始日期
        st.markdown('<span class="sub-label">開始日期　START DATE</span>', unsafe_allow_html=True)
        s1, s2, s3 = st.columns(3)
        with s1: start_year  = st.selectbox("民國年", options=YEAR_OPTIONS, index=YEAR_OPTIONS.index(roc_now), key="sy")
        with s2: start_month = st.selectbox("月", options=MONTH_OPTIONS, index=MONTH_OPTIONS.index(today.month), key="sm")
        with s3: start_day   = st.selectbox("日", options=DAY_OPTIONS, index=DAY_OPTIONS.index(today.day), key="sd")

        # 結束日期
        st.markdown('<span class="sub-label">結束日期　END DATE</span>', unsafe_allow_html=True)
        e1, e2, e3 = st.columns(3)
        with e1: end_year  = st.selectbox("民國年", options=YEAR_OPTIONS, index=YEAR_OPTIONS.index(roc_now), key="ey")
        with e2: end_month = st.selectbox("月", options=MONTH_OPTIONS, index=MONTH_OPTIONS.index(today.month), key="em")
        with e3: end_day   = st.selectbox("日", options=DAY_OPTIONS, index=DAY_OPTIONS.index(today.day), key="ed")

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

    # ── Generate ──────────────────────────────────────────────────────────────
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

# ── 左右欄等高同步 ─────────────────────────────────────────────────────────
# 直接量測「請假時間」卡片的實際高度，反推「假別與事由」卡片要多高，
# 讓左右兩欄的底部對齊（比用 CSS 猜測 DOM 結構更可靠）。
components.html("""
<script>
(function() {
  function sync() {
    var doc = window.parent.document;
    var apply = doc.querySelector('.st-key-panel_apply');
    var leave = doc.querySelector('.st-key-panel_leave');
    var period = doc.querySelector('.st-key-panel_period');
    if (!apply || !leave || !period) return;

    leave.style.minHeight = '0px';
    var applyBox  = apply.getBoundingClientRect();
    var leaveBox  = leave.getBoundingClientRect();
    var periodBox = period.getBoundingClientRect();

    var gap = leaveBox.top - applyBox.bottom; // 左欄兩張卡片之間目前的間距
    var needed = periodBox.height - applyBox.height - gap;

    if (needed > leaveBox.height) {
      leave.style.minHeight = needed + 'px';
    }
  }
  setTimeout(sync, 60);
  setTimeout(sync, 300);
  setTimeout(sync, 800);
  window.addEventListener('resize', sync);
})();
</script>
""", height=0)
