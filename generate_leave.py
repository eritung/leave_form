import io
import openpyxl

TEMPLATE_PATH = "艾迪英特請假單_template.xlsx"

LEAVE_TYPES = {
    "事假": ("D8", "D29"),
    "病假": ("F8", "F29"),
    "特休": ("H8", "H29"),
    "產/婚/喪假": ("J8", "J29"),
    "其他": ("M8", "M29"),
}

def generate_leave_xlsx(
    apply_year: int,
    apply_month: int,
    apply_day: int,
    applicant: str,
    proxy: str,
    leave_type: str,
    reason: str,
    start_year: int,
    start_month: int,
    start_day: int,
    start_hour: int,
    start_minute: int,
    end_year: int,
    end_month: int,
    end_day: int,
    end_hour: int,
    end_minute: int,
    total_days: str,
) -> bytes:
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["假單"]

    # Enable print centering
    ws.print_options.horizontalCentered = True

    def fill(cell_addr, value):
        ws[cell_addr] = value

    # === 公司留存 (rows 4-18) ===
    fill("I4", apply_year)
    fill("K4", apply_month)
    fill("M4", apply_day)

    fill("D5", applicant)
    fill("K5", proxy)

    # Clear all leave type cells first, then mark selected
    for lt, (c1, _) in LEAVE_TYPES.items():
        ws[c1] = ""
    if leave_type in LEAVE_TYPES:
        fill(LEAVE_TYPES[leave_type][0], "○")

    fill("D9", reason)

    fill("C11", start_year)
    fill("E11", start_month)
    fill("G11", start_day)
    fill("I11", start_hour)
    fill("K11", start_minute)

    fill("C12", end_year)
    fill("E12", end_month)
    fill("G12", end_day)
    fill("I12", end_hour)
    fill("K12", end_minute)

    fill("M11", total_days)

    # === 請假者留存 (rows 25-39) ===
    fill("I25", apply_year)
    fill("K25", apply_month)
    fill("M25", apply_day)

    fill("D26", applicant)
    fill("K26", proxy)

    for lt, (_, c2) in LEAVE_TYPES.items():
        ws[c2] = ""
    if leave_type in LEAVE_TYPES:
        fill(LEAVE_TYPES[leave_type][1], "○")

    fill("D30", reason)

    fill("C32", start_year)
    fill("E32", start_month)
    fill("G32", start_day)
    fill("I32", start_hour)
    fill("K32", start_minute)

    fill("C33", end_year)
    fill("E33", end_month)
    fill("G33", end_day)
    fill("I33", end_hour)
    fill("K33", end_minute)

    fill("M32", total_days)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
