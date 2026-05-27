from __future__ import annotations

from datetime import date, datetime, time
from io import BytesIO
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

TEMPLATE_PATH = Path(__file__).with_name("艾迪英特請假單.xlsx")
SHEET_NAME = "假單"

LEAVE_TYPES = ["事假", "病假", "特休", "產/婚/喪假", "其他"]

# 兩聯式表單：上半部公司留存，下半部請假者留存
# 這些都是各欄位合併儲存格的左上角，寫入左上角即可保留原版型。
SECTION_ROWS = {
    "公司留存": {
        "date": 4,
        "name": 5,
        "leave_type": 8,
        "reason": 9,
        "start": 11,
        "end": 12,
    },
    "請假者留存": {
        "date": 25,
        "name": 26,
        "leave_type": 29,
        "reason": 30,
        "start": 32,
        "end": 33,
    },
}

LEAVE_TYPE_CELLS = {
    "事假": "D",
    "病假": "F",
    "特休": "H",
    "產/婚/喪假": "J",
    "其他": "M",
}


def roc_year(value: date) -> int:
    """Convert western year to ROC/Minguo year."""
    return value.year - 1911


def calculate_leave_days(start_date: date, start_time: time, end_date: date, end_time: time) -> float:
    """Rough default: elapsed hours / 8. User can manually override in the form."""
    start_dt = datetime.combine(start_date, start_time)
    end_dt = datetime.combine(end_date, end_time)
    if end_dt <= start_dt:
        return 0.0
    hours = (end_dt - start_dt).total_seconds() / 3600
    return round(hours / 8, 2)


def write_centered(ws, cell: str, value):
    ws[cell] = value
    ws[cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[cell].font = Font(name="PMingLiu", size=12)


def fill_section(ws, rows: dict[str, int], data: dict):
    # 申請日期
    date_row = rows["date"]
    write_centered(ws, f"I{date_row}", roc_year(data["apply_date"]))
    write_centered(ws, f"K{date_row}", data["apply_date"].month)
    write_centered(ws, f"M{date_row}", data["apply_date"].day)

    # 請假人 / 代理人
    name_row = rows["name"]
    write_centered(ws, f"D{name_row}", data["applicant"])
    write_centered(ws, f"K{name_row}", data["proxy"])

    # 假別：先清空勾選列，再勾選目標假別
    leave_row = rows["leave_type"]
    for col in LEAVE_TYPE_CELLS.values():
        write_centered(ws, f"{col}{leave_row}", "")
    selected_col = LEAVE_TYPE_CELLS[data["leave_type"]]
    check_text = "✓" if data["leave_type"] != "其他" or not data["other_leave_note"] else f"✓ {data['other_leave_note']}"
    write_centered(ws, f"{selected_col}{leave_row}", check_text)

    # 事由
    reason_row = rows["reason"]
    ws[f"D{reason_row}"] = data["reason"]
    ws[f"D{reason_row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws[f"D{reason_row}"].font = Font(name="PMingLiu", size=12)

    # 請假起訖時間
    start_row = rows["start"]
    end_row = rows["end"]
    start_date = data["start_date"]
    end_date = data["end_date"]
    start_time = data["start_time"]
    end_time = data["end_time"]

    write_centered(ws, f"C{start_row}", roc_year(start_date))
    write_centered(ws, f"E{start_row}", start_date.month)
    write_centered(ws, f"G{start_row}", start_date.day)
    write_centered(ws, f"I{start_row}", f"{start_time.hour:02d}")
    write_centered(ws, f"K{start_row}", f"{start_time.minute:02d}")
    write_centered(ws, f"M{start_row}", data["leave_days"])

    write_centered(ws, f"C{end_row}", roc_year(end_date))
    write_centered(ws, f"E{end_row}", end_date.month)
    write_centered(ws, f"G{end_row}", end_date.day)
    write_centered(ws, f"I{end_row}", f"{end_time.hour:02d}")
    write_centered(ws, f"K{end_row}", f"{end_time.minute:02d}")


def build_leave_form_xlsx(data: dict) -> bytes:
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET_NAME]

    for rows in SECTION_ROWS.values():
        fill_section(ws, rows, data)

    # 讓輸出範圍乾淨，避免 Excel 誤抓到很後面的空白格。
    ws.print_area = "A1:N39"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.set_page_config(page_title="艾迪英特請假單產生器", page_icon="📝", layout="centered")

st.markdown(
    """
    <style>
    .block-container {max-width: 860px; padding-top: 2rem;}
    .ad2-card {
        padding: 1.1rem 1.25rem;
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        background: #ffffff;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
        margin-bottom: 1rem;
    }
    .hint {color: #64748b; font-size: 0.92rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("艾迪英特請假單產生器")
st.caption("填寫資料後，直接下載可使用的 xlsx 請假單。兩聯內容會同步帶入：公司留存＋請假者留存。")

with st.form("leave_form"):
    st.markdown('<div class="ad2-card">', unsafe_allow_html=True)
    st.subheader("基本資料")
    col1, col2 = st.columns(2)
    with col1:
        apply_date = st.date_input("申請日期", value=date.today())
        applicant = st.text_input("請假人", placeholder="請輸入姓名")
    with col2:
        proxy = st.text_input("代理人", placeholder="請輸入代理人姓名，可留空")
        leave_type = st.selectbox("假別", LEAVE_TYPES)

    other_leave_note = ""
    if leave_type == "其他":
        other_leave_note = st.text_input("其他假別說明", placeholder="例如：公假、補休")

    reason = st.text_area("事由", placeholder="請簡述請假原因", height=90)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="ad2-card">', unsafe_allow_html=True)
    st.subheader("請假時間")
    col3, col4 = st.columns(2)
    with col3:
        start_date = st.date_input("開始日期", value=date.today())
        start_time = st.time_input("開始時間", value=time(9, 0))
    with col4:
        end_date = st.date_input("結束日期", value=date.today())
        end_time = st.time_input("結束時間", value=time(18, 0))

    suggested_days = calculate_leave_days(start_date, start_time, end_date, end_time)
    leave_days = st.number_input("請假天數", min_value=0.0, max_value=365.0, value=suggested_days, step=0.5)
    st.markdown('<p class="hint">小提醒：系統會先用「起訖時間 ÷ 8 小時」估算天數，但仍可手動調整，例如半天填 0.5。</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    submitted = st.form_submit_button("產生請假單 xlsx", use_container_width=True)

if submitted:
    errors = []
    if not applicant.strip():
        errors.append("請填寫請假人。")
    if not reason.strip():
        errors.append("請填寫事由。")
    if datetime.combine(end_date, end_time) <= datetime.combine(start_date, start_time):
        errors.append("結束時間需晚於開始時間。")

    if errors:
        for error in errors:
            st.error(error)
    else:
        data = {
            "apply_date": apply_date,
            "applicant": applicant.strip(),
            "proxy": proxy.strip(),
            "leave_type": leave_type,
            "other_leave_note": other_leave_note.strip(),
            "reason": reason.strip(),
            "start_date": start_date,
            "start_time": start_time,
            "end_date": end_date,
            "end_time": end_time,
            "leave_days": leave_days,
        }
        xlsx_bytes = build_leave_form_xlsx(data)
        filename = f"艾迪英特請假單_{applicant.strip()}_{apply_date.strftime('%Y%m%d')}.xlsx"
        st.success("請假單已產生，可以下載囉。")
        st.download_button(
            label="下載 xlsx",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
