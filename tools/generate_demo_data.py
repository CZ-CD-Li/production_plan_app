"""生成用于本地演示的 Excel 数据。

运行：python tools/generate_demo_data.py
"""

from __future__ import annotations

import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DEMO_DIR = ROOT / "demo_data"
PROJECT_FILE = DEMO_DIR / "新产品项目任务分布清单.xlsx"

BORDER_COLOR = "D5DEEA"
THIN = Side(style="thin", color=BORDER_COLOR)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TITLE_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FILL = PatternFill("solid", fgColor="2F5597")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
LINE_FILL = PatternFill("solid", fgColor="EAF0F9")

MASS_COLORS = {
    4: "8B1A1A",
    3: "C55A11",
    2: "2F855A",
    1: "6B46C1",
}


PRODUCTS = [
    # 主机, 产品名称, 公司代号, 生产线, 焊接, 装配, 机加, mass
    ("北汽制造", "星耀 X5", "LK-A301", "C1线", "张伟", "王强", "孙敏", 4),
    ("吉利汽车", "星越 L", "LK-A302", "C1线", "赵磊", "张伟", "孙敏", 3),
    ("比亚迪", "秦 PLUS", "LK-B205", "C2线", "刘洋", "陈晨", "周涛", 4),
    ("长安汽车", "深蓝 S7", "LK-C106", "C2线", "刘洋", "陈晨", "周涛", 2),
    ("奇瑞汽车", "瑞虎 8", "LK-D408", "B1线", "何军", "李雪", "吴磊", 1),
    ("上汽集团", "荣威 RX5", "LK-E512", "B1线", "何军", "李雪", "吴磊", 3),
    ("广汽埃安", "AION Y", "LK-F617", "B2线", "郑凯", "林静", "马超", 2),
    ("东风日产", "轩逸", "LK-G721", "B2线", "郑凯", "林静", "马超", 4),
    ("长城汽车", "哈弗 H6", "LK-H825", "C1线", "赵磊", "王强", "孙敏", 1),
    ("一汽红旗", "HS5", "LK-J930", "B1线", "何军", "李雪", "吴磊", 2),
    ("江淮汽车", "瑞风 S7", "LK-K034", "C2线", "刘洋", "陈晨", "周涛", 0),
    ("零跑汽车", "C11", "LK-L138", "B2线", "郑凯", "林静", "马超", 0),
    ("东风本田", "思域", "LK-M441", "C3线", "沈岩", "顾清", "许亮", 4),
    ("广汽传祺", "影豹", "LK-N552", "C3线", "沈岩", "顾清", "许亮", 2),
    ("蔚来汽车", "ET5", "LK-R773", "C4线", "范宁", "夏雨", "秦朗", 3),
    ("小鹏汽车", "G6", "LK-S886", "C4线", "范宁", "夏雨", "秦朗", 1),
    ("北汽新能源", "EU5", "LK-T884", "B3线", "杜鹏", "苏婉", "蒋涛", 3),
    ("岚图汽车", "FREE", "LK-U995", "B3线", "杜鹏", "苏婉", "蒋涛", 1),
    ("极氪", "001", "LK-V116", "B4线", "白桦", "温宁", "程亮", 4),
    ("哪吒汽车", "S", "LK-W227", "B4线", "白桦", "温宁", "程亮", 2),
]

ASSEMBLY_CODES = [
    "LK-A301", "LK-A302", "LK-B205", "LK-C106",
    "LK-H825", "LK-M441", "LK-N552", "LK-R773",
]
WELDING_CODES = [
    "LK-D408", "LK-E512", "LK-F617", "LK-G721",
    "LK-J930", "LK-T884", "LK-U995", "LK-V116",
]


def make_project_file() -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "新产品项目任务分布清单"
    sheet.sheet_view.showGridLines = False

    sheet.merge_cells("A1:L1")
    title = sheet["A1"]
    title.value = "新产品项目任务分布清单"
    title.font = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
    title.fill = TITLE_FILL
    title.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 34

    headers = ["序号", "主机", "产品名称", "公司代号", "生产线", "焊接", "装配", "机加", "备注"]
    for index, header in enumerate(headers, start=1):
        cell = sheet.cell(2, index, header)
        cell.font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    sheet.cell(1, 13, "颜色等级").alignment = Alignment(horizontal="center", vertical="center")
    sheet.cell(1, 13).font = Font(name="微软雅黑", size=10, color="7C8AA0")

    for index, (mass, color) in enumerate(MASS_COLORS.items()):
        cell = sheet.cell(1, 14 + index)
        cell.value = mass
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for row_index, product in enumerate(PRODUCTS, start=3):
        host, name, code, line, welding, assembly, machining, mass = product
        values = [row_index - 2, host, name, code, line, welding, assembly, machining, ""]
        for col_index, value in enumerate(values, start=1):
            cell = sheet.cell(row_index, col_index, value)
            cell.font = Font(name="微软雅黑", size=10, color="26364D")
            cell.alignment = Alignment(
                horizontal="left" if col_index in (2, 3) else "center",
                vertical="center",
            )
            cell.border = BORDER
            if (row_index - 3) % 2 == 1:
                cell.fill = PatternFill("solid", fgColor="F8FAFD")
        code_cell = sheet.cell(row_index, 4)
        if mass in MASS_COLORS:
            code_cell.fill = PatternFill("solid", fgColor=MASS_COLORS[mass])
            code_cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        else:
            code_cell.fill = PatternFill("solid", fgColor="FFFFFF")
        sheet.row_dimensions[row_index].height = 24

    widths = [8, 15, 18, 16, 11, 10, 10, 10, 18]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.column_dimensions["M"].width = 10
    for index in range(14, 18):
        sheet.column_dimensions[get_column_letter(index)].width = 10
    sheet.freeze_panes = "A3"

    PROJECT_FILE.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(PROJECT_FILE)
    print(f"已生成项目分布清单：{PROJECT_FILE}")


def make_plan_file(today: date) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = today.strftime("%Y%m%d")
    sheet.sheet_view.showGridLines = False

    days = [today + timedelta(days=offset) for offset in (-2, -1, 0, 1, 2)]
    block_width = 4

    last_col = 2 + block_width * len(days)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    title = sheet["A1"]
    title.value = "301车间 · 五日滚动生产计划"
    title.font = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
    title.fill = TITLE_FILL
    title.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 34

    sheet.cell(2, 1, "生产线").font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    sheet.cell(2, 2, "主机厂").font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    for col in (1, 2):
        cell = sheet.cell(2, col)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for index, day in enumerate(days):
        start_col = 3 + index * block_width
        end_col = start_col + block_width - 1
        sheet.merge_cells(
            start_row=2, start_column=start_col, end_row=2, end_column=end_col
        )
        cell = sheet.cell(2, start_col, datetime.combine(day, time.min))
        cell.number_format = "yyyy/m/d"
        cell.font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(start_col, end_col + 1):
            sheet.cell(2, col).border = BORDER

        labels = ["产品代号", "计划数量", "节拍", "备注"]
        for offset, label in enumerate(labels):
            cell = sheet.cell(3, start_col + offset, label)
            cell.font = Font(name="微软雅黑", size=9, bold=True, color="3B4A61")
            cell.fill = SUBHEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
    sheet.row_dimensions[2].height = 28
    sheet.row_dimensions[3].height = 24

    line_groups = [
        (4, 5, "C1线"),
        (6, 7, "C2线"),
        (8, 9, "C3线"),
        (10, 11, "C4线"),
        (12, 13, "B1线"),
        (14, 15, "B2线"),
        (16, 17, "B3线"),
        (18, 19, "B4线"),
    ]
    for start_row, end_row, line_name in line_groups:
        sheet.merge_cells(
            start_row=start_row, start_column=1, end_row=end_row, end_column=1
        )
        sheet.merge_cells(
            start_row=start_row, start_column=2, end_row=end_row, end_column=2
        )
        cell = sheet.cell(start_row, 1, line_name)
        cell.font = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
        cell.fill = LINE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in range(start_row, end_row + 1):
            for col in (1, 2):
                border_cell = sheet.cell(row, col)
                border_cell.border = BORDER
                if col == 1:
                    border_cell.fill = LINE_FILL
        sheet.cell(start_row, 2).alignment = Alignment(horizontal="center", vertical="center")

    for row in range(4, 20):
        sheet.row_dimensions[row].height = 26

    for day_index, _day in enumerate(days):
        start_col = 3 + day_index * block_width
        slot_codes = []
        slot_codes.extend(
            ASSEMBLY_CODES[(day_index + offset) % len(ASSEMBLY_CODES)]
            for offset in range(8)
        )
        slot_codes.extend(
            WELDING_CODES[(day_index + offset) % len(WELDING_CODES)]
            for offset in range(8)
        )

        for row_index, code in enumerate(slot_codes):
            row = 4 + row_index
            quantity = (7 - (row_index % 8)) * 10 + 60 + day_index * 5
            remark_pool = ["优先排产", "", "等待物料", "加急", ""]
            remark = remark_pool[(day_index + row_index) % len(remark_pool)]

            code_cell = sheet.cell(row, start_col, code)
            code_cell.font = Font(name="微软雅黑", size=10, bold=True, color="22324A")
            code_cell.alignment = Alignment(horizontal="center", vertical="center")
            code_cell.border = BORDER

            qty_cell = sheet.cell(row, start_col + 1, quantity)
            qty_cell.font = Font(name="微软雅黑", size=10, color="26364D")
            qty_cell.alignment = Alignment(horizontal="center", vertical="center")
            qty_cell.border = BORDER

            cycle_cell = sheet.cell(row, start_col + 2, 45 + row_index)
            cycle_cell.font = Font(name="微软雅黑", size=10, color="7C8AA0")
            cycle_cell.alignment = Alignment(horizontal="center", vertical="center")
            cycle_cell.border = BORDER

            remark_cell = sheet.cell(row, start_col + 3, remark)
            remark_cell.font = Font(name="微软雅黑", size=10, color="52627A")
            remark_cell.alignment = Alignment(horizontal="left", vertical="center")
            remark_cell.border = BORDER

    sheet.column_dimensions["A"].width = 12
    sheet.column_dimensions["B"].width = 13
    for day_index in range(len(days)):
        start_col = 3 + day_index * block_width
        sheet.column_dimensions[get_column_letter(start_col)].width = 14
        sheet.column_dimensions[get_column_letter(start_col + 1)].width = 10
        sheet.column_dimensions[get_column_letter(start_col + 2)].width = 8
        sheet.column_dimensions[get_column_letter(start_col + 3)].width = 18

    sheet.freeze_panes = "C4"
    sheet.auto_filter.ref = f"A2:{get_column_letter(last_col)}19"

    month_dir = DEMO_DIR / f"{today.year}年" / f"{today.month}月"
    month_dir.mkdir(parents=True, exist_ok=True)
    for old_file in DEMO_DIR.rglob("301_*.xlsx"):
        try:
            old_file.unlink()
        except OSError:
            pass
    plan_file = month_dir / f"301_生产计划_{today:%Y%m}.xlsx"
    workbook.save(plan_file)
    print(f"已生成生产计划：{plan_file}")
    return plan_file


def main() -> None:
    DEMO_DIR.mkdir(parents=True, exist_ok=True)
    make_project_file()
    make_plan_file(date.today())


if __name__ == "__main__":
    sys.exit(main())
