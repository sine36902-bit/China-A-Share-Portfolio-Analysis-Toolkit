
import baostock as bs
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

# =====================================
# 1. 选择保存位置
# =====================================

root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)

save_path = filedialog.asksaveasfilename(
    title="选择半强式有效性分析Excel的保存位置",
    initialfile="半强式有效性_含Excel公式.xlsx",
    defaultextension=".xlsx",
    filetypes=[("Excel 工作簿", "*.xlsx")]
)

root.destroy()

if not save_path:
    raise SystemExit("已取消保存")

# =====================================
# 2. 股票与事件日期
# =====================================

stocks = {
    "美诺华": ("sh.603538", "1"),
    "三人行": ("sh.605168", "1"),
    "沪深300": ("sh.000300", "3")
}

events = {
    "美诺华": "2021-04-23",
    "三人行": "2023-03-28"
}

# =====================================
# 3. 从BaoStock获取行情
# =====================================

login = bs.login()

if login.error_code != "0":
    raise RuntimeError(login.error_msg)

prices = {}

try:
    for name, (code, adjustflag) in stocks.items():

        rs = bs.query_history_k_data_plus(
            code,
            "date,close",
            start_date="2021-04-01",
            end_date="2023-05-15",
            frequency="d",
            adjustflag=adjustflag
        )

        if rs.error_code != "0":
            raise RuntimeError(f"{name}: {rs.error_msg}")

        rows = []

        while rs.next():
            rows.append(rs.get_row_data())

        df = pd.DataFrame(rows, columns=rs.fields)

        if df.empty:
            raise ValueError(f"{name}没有获取到数据")

        df["date"] = pd.to_datetime(df["date"])
        df["close"] = pd.to_numeric(
            df["close"], errors="coerce"
        )

        prices[name] = (
            df.set_index("date")["close"].sort_index()
        )

        print(f"{name}：获取 {len(df)} 条记录")

finally:
    bs.logout()

raw = pd.DataFrame(prices).sort_index()

# =====================================
# 4. 创建Excel工作簿
# =====================================

wb = Workbook()
ws_raw = wb.active
ws_raw.title = "原始价格"

ws_raw.append([
    "日期",
    "美诺华（后复权）",
    "三人行（后复权）",
    "沪深300（原始点位）"
])

date_to_row = {}

for date, values in raw.iterrows():

    row_num = ws_raw.max_row + 1
    date_to_row[date] = row_num

    ws_raw.append([
        date.to_pydatetime(),
        float(values["美诺华"])
            if pd.notna(values["美诺华"]) else None,
        float(values["三人行"])
            if pd.notna(values["三人行"]) else None,
        float(values["沪深300"])
            if pd.notna(values["沪深300"]) else None
    ])

    ws_raw.cell(row_num, 1).number_format = "yyyy-mm-dd"

# =====================================
# 5. 创建事件研究表：使用Excel公式
# =====================================

stock_cols = {
    "美诺华": "B",
    "三人行": "C"
}

for name, event_date in events.items():

    ws = wb.create_sheet(f"{name}事件研究")

    ws.append([
        "T",
        "交易日期",
        "个股收盘价",
        "沪深300点位",
        "个股日收益率",
        "市场日收益率",
        "AR",
        "CAR"
    ])

    # 统一使用个股和指数均有收盘数据的日期
    common = raw[[name, "沪深300"]].dropna()
    event_date = pd.Timestamp(event_date)

    if event_date not in common.index:
        raise ValueError(
            f"{name}：找不到事件日 {event_date.date()}"
        )

    event_pos = common.index.get_loc(event_date)

    # 提前多取一天T=-6，用来计算T=-5的日收益率
    window = common.iloc[
        event_pos - 6:event_pos + 11
    ]

    if len(window) != 17:
        raise ValueError(
            f"{name}：缺少计算所需的交易日"
        )

    for i, date in enumerate(window.index):

        excel_row = i + 2
        original_row = date_to_row[date]

        stock_col = stock_cols[name]

        # T=-6作为计算基期，不纳入CAR
        t = i - 6

        ws.cell(excel_row, 1, t)

        # 日期和价格直接引用原始价格表
        ws.cell(
            excel_row, 2,
            f"='原始价格'!A{original_row}"
        )

        ws.cell(
            excel_row, 3,
            f"='原始价格'!{stock_col}{original_row}"
        )

        ws.cell(
            excel_row, 4,
            f"='原始价格'!D{original_row}"
        )

        ws.cell(excel_row, 2).number_format = "yyyy-mm-dd"

        # T=-6只用于提供前一日价格
        if i == 0:
            continue

        prev = excel_row - 1

        # 个股日收益率
        ws.cell(
            excel_row, 5,
            f"=C{excel_row}/C{prev}-1"
        )

        # 沪深300日收益率
        ws.cell(
            excel_row, 6,
            f"=D{excel_row}/D{prev}-1"
        )

        # 异常收益率 AR
        ws.cell(
            excel_row, 7,
            f"=E{excel_row}-F{excel_row}"
        )

        # 累计异常收益率 CAR
        ws.cell(
            excel_row, 8,
            f"=SUM($G$3:G{excel_row})"
        )

        # 将收益率格式化成百分比
        for col in range(5, 9):
            ws.cell(
                excel_row, col
            ).number_format = "0.00%"

# =====================================
# 6. 结果汇总（也使用Excel公式）
# =====================================

summary = wb.create_sheet("结果汇总", 0)

summary.append([
    "股票",
    "T=0 AR",
    "T=0 CAR",
    "T=+10 CAR",
    "公告后CAR变化（百分点）"
])

for row, name in enumerate(events, start=2):

    sheet = f"'{name}事件研究'"

    summary.cell(row, 1, name)

    # T=-6在第2行，因此T=0在第8行
    summary.cell(row, 2, f"={sheet}!G8")
    summary.cell(row, 3, f"={sheet}!H8")

    # T=+10在第18行
    summary.cell(row, 4, f"={sheet}!H18")

    # 公告后CAR变化：+10日CAR减去第0日CAR
    summary.cell(
        row, 5,
        f"=(D{row}-C{row})*100"
    )

    for col in (2, 3, 4):
        summary.cell(
            row, col
        ).number_format = "0.00%"

    summary.cell(row, 5).number_format = "0.00"

# =====================================
# 7. 统一表格格式
# =====================================

for ws in wb.worksheets:

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )
        cell.fill = PatternFill(
            "solid",
            fgColor="24476B"
        )
        cell.alignment = Alignment(
            horizontal="center"
        )

    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        ws.column_dimensions[letter].width = 19

summary.column_dimensions["E"].width = 30

# 打开时让Excel自动重新计算公式
wb.calculation = CalcProperties(
    calcMode="auto",
    fullCalcOnLoad=True,
    forceFullCalc=True
)

# =====================================
# 8. 保存
# =====================================

wb.save(save_path)

print("\nExcel文件已成功保存")
print("保存位置：", Path(save_path).resolve())
