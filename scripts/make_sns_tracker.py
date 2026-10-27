"""LINE登録数・投稿メモ / 月別ストーリーズ閲覧数 の記録用スプレッドシートを作成する。"""
import calendar
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = "sns_tracking/SNS記録シート.xlsx"
START = date(2026, 10, 1)
MONTHS = [(2026, 10), (2026, 11), (2026, 12), (2027, 1), (2027, 2), (2027, 3)]
STORY_ROWS = 40  # 1か月あたりのストーリーズ記入枠

HEAD_FILL = PatternFill("solid", fgColor="06C755")  # LINEグリーン
STORY_FILL = PatternFill("solid", fgColor="C13584")  # インスタ系
SUB_FILL = PatternFill("solid", fgColor="F3F3F3")
WEEKEND_FILL = PatternFill("solid", fgColor="FFF4E5")
WHITE_BOLD = Font(bold=True, color="FFFFFF")
thin = Side(style="thin", color="CCCCCC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center")
WRAP = Alignment(vertical="top", wrap_text=True)
WD = "月火水木金土日"


def header(ws, row, labels, fill):
    for col, label in enumerate(labels, 1):
        c = ws.cell(row=row, column=col, value=label)
        c.fill, c.font, c.alignment, c.border = fill, WHITE_BOLD, CENTER, BORDER


wb = Workbook()

# --- LINE登録数（毎日） ---
ws = wb.active
ws.title = "LINE登録数"
header(ws, 1, ["日付", "曜日", "LINE登録数", "累計", "📝 その日の投稿内容", "投稿先", "メモ"], HEAD_FILL)
end = date(*MONTHS[-1], calendar.monthrange(*MONTHS[-1])[1])
d, r = START, 2
while d <= end:
    ws.cell(row=r, column=1, value=d).number_format = "yyyy/mm/dd"
    ws.cell(row=r, column=2, value=WD[d.weekday()])
    ws.cell(row=r, column=4, value=f"=IF(C{r}=\"\",\"\",SUM($C$2:C{r}))")
    for col in range(1, 8):
        c = ws.cell(row=r, column=col)
        c.border = BORDER
        c.alignment = WRAP if col in (5, 7) else CENTER
        if d.weekday() >= 5:
            c.fill = WEEKEND_FILL
    d += timedelta(days=1)
    r += 1
last = r - 1
for col, w in zip("ABCDEFG", [12, 6, 12, 10, 50, 14, 30]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:G{last}"

# --- 月別まとめ ---
sm = wb.create_sheet("月別まとめ")
header(sm, 1, ["月", "LINE登録数（合計）", "1日平均", "ストーリーズ本数", "閲覧数（合計）", "閲覧数（平均）"], HEAD_FILL)
for i, (y, m) in enumerate(MONTHS, 2):
    first, lastday = date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])
    s = f"{m}月ストーリーズ"
    sm.cell(row=i, column=1, value=f"{y}年{m}月")
    sm.cell(row=i, column=2, value=f"=SUMIFS('LINE登録数'!C:C,'LINE登録数'!A:A,\">=\"&DATE({y},{m},1),'LINE登録数'!A:A,\"<=\"&DATE({y},{m},{lastday.day}))")
    sm.cell(row=i, column=3, value=f"=ROUND(B{i}/{lastday.day},1)")
    sm.cell(row=i, column=4, value=f"=COUNT('{s}'!E5:E{4 + STORY_ROWS})")
    sm.cell(row=i, column=5, value=f"=SUM('{s}'!E5:E{4 + STORY_ROWS})")
    sm.cell(row=i, column=6, value=f"=IFERROR(ROUND(AVERAGE('{s}'!E5:E{4 + STORY_ROWS}),0),\"\")")
    for col in range(1, 7):
        sm.cell(row=i, column=col).border = BORDER
        sm.cell(row=i, column=col).alignment = CENTER
for col, w in zip("ABCDEF", [12, 18, 10, 16, 16, 16]):
    sm.column_dimensions[col].width = w

# --- 月別ストーリーズ ---
for y, m in MONTHS:
    st = wb.create_sheet(f"{m}月ストーリーズ")
    first = 5
    lastrow = first + STORY_ROWS - 1
    st["A1"] = f"{y}年{m}月 ストーリーズ記録"
    st["A1"].font = Font(bold=True, size=14)
    for col, (label, formula) in enumerate([
        ("本数", f"=COUNT(E{first}:E{lastrow})"),
        ("閲覧数合計", f"=SUM(E{first}:E{lastrow})"),
        ("閲覧数平均", f"=IFERROR(ROUND(AVERAGE(E{first}:E{lastrow}),0),\"\")"),
        ("最高閲覧数", f"=IF(COUNT(E{first}:E{lastrow})=0,\"\",MAX(E{first}:E{lastrow}))"),
    ]):
        lc = st.cell(row=2, column=2 + col * 1, value=label)
        vc = st.cell(row=3, column=2 + col * 1, value=formula)
        lc.fill, lc.alignment, lc.border = SUB_FILL, CENTER, BORDER
        vc.font, vc.alignment, vc.border = Font(bold=True, size=13), CENTER, BORDER
    header(st, 4, ["No.", "投稿日", "📷 スクショ（ここに貼る）", "内容メモ", "閲覧数", "リアクション等", "備考"], STORY_FILL)
    for i in range(STORY_ROWS):
        r = first + i
        st.cell(row=r, column=1, value=i + 1)
        st.cell(row=r, column=2).number_format = "m/d"
        st.row_dimensions[r].height = 220
        for col in range(1, 8):
            c = st.cell(row=r, column=col)
            c.border = BORDER
            c.alignment = WRAP if col in (4, 6, 7) else CENTER
    for col, w in zip("ABCDEFG", [5, 9, 22, 30, 10, 16, 20]):
        st.column_dimensions[col].width = w
    st.column_dimensions["E"].width = 12
    st.conditional_formatting.add(
        f"E{first}:E{lastrow}",
        FormulaRule(formula=[f'AND(E{first}<>"",E{first}=MAX($E${first}:$E${lastrow}))'], fill=PatternFill("solid", fgColor="FFE08A")),
    )
    st.freeze_panes = "A5"

wb.save(OUT)
print("saved", OUT)
