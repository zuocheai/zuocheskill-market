# -*- coding: utf-8 -*-
"""生成「学员台账」Excel（可自动算剩余课时、自动挑出今天该催谁）。

用法：
    python 建学员台账.py [输出路径]

不填任何业务数据 —— 只建骨架、公式和配色，老板自己填。
三条判据（见 资产/学员台账模板说明.md）：
    1. 剩余课时 ≤ 2 节          → 课时告急，催续费
    2. 距卡到期 ≤ 14 天         → 将到期，给续费钩
    3. 距上次上课 ≥ 14 天       → 流失风险，发关心式促活（不硬推）
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

DEFAULT_OUT = os.path.expanduser("~/Documents/瑜伽工作台/学员台账.xlsx")

HEAD_FILL = PatternFill("solid", fgColor="1D9E75")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
ALERT_FILL = PatternFill("solid", fgColor="FCEBEB")   # 课时告急
SOON_FILL = PatternFill("solid", fgColor="FAEEDA")    # 将到期
RISK_FILL = PatternFill("solid", fgColor="E6F1FB")    # 流失风险
TITLE_FONT = Font(bold=True, size=14, color="04342C")
THIN = Side(style="thin", color="D3D1C7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

# 与 资产/学员台账模板.csv 同序，共 13 列（A..M）
COLS = [
    ("学员名", 12),
    ("微信备注", 14),
    ("手机号", 13),
    ("卡种", 16),
    ("报课日期", 12),
    ("总课时", 9),
    ("已耗课时", 10),
    ("剩余课时", 10),   # H 公式
    ("卡到期日", 12),
    ("上次上课", 12),
    ("负责教练", 12),
    ("状态", 12),       # L 手动/参考
    ("备注", 30),
]

STATES = "活跃,将到期,课时告急,流失风险,已流失"

GUIDE = [
    ("怎么用这份表", ""),
    ("① 填", "每行一个学员，从第 4 行开始往下填。只填你自己知道的，别的先空着。"),
    ("② 不用算", "「剩余课时」「状态」两列自带公式，填完上面的数字它们自己算，别手动改。"),
    ("③ 配色", "整行变红＝课时告急（剩2节内·该催）；变黄＝将到期（14天内）；变蓝＝流失风险（14天没来）。"),
    ("④ 手机号", "属个人信息，只存老板自己发来的、只在本机，不外传。"),
    ("", ""),
    ("每天只做一件事", ""),
    ("早上打开", "只看颜色。三种颜色的人，就是今天该联系的——红的最急，黄的次之，蓝的慢慢聊。"),
    ("", ""),
    ("三条判据（写死·见 资产/学员台账模板说明.md）", ""),
    ("判据一", "剩余课时 ≤ 2 节 → 课时告急，催续费。"),
    ("判据二", "距卡到期 ≤ 14 天 → 将到期，给续费优惠钩。"),
    ("判据三", "距上次上课 ≥ 14 天 且非已流失 → 流失风险，发一条关心式促活（不是硬推销）。"),
    ("", ""),
    ("表里没有的，怎么补", ""),
    ("排课", "排课是另一张表，跟这张勾稽不上，建议单独做。"),
    ("课程价格", "在「卡种」后面加几列，把单次价、课包价写死，以后算钱不用回忆。"),
]


def build(out_path=DEFAULT_OUT):
    wb = Workbook()

    ws = wb.active
    ws.title = "学员台账"
    ws["A1"] = "学员台账"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("填完「总课时/已耗课时/卡到期日/上次上课」后，「剩余课时」「状态」自动算，"
                "整行按三条判据变色")
    ws["A2"].font = Font(size=10, color="5F5E5A")
    n = len(COLS)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)

    HEADER_ROW = 3
    for idx, (name, width) in enumerate(COLS, start=1):
        c = ws.cell(row=HEADER_ROW, column=idx, value=name)
        c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, WRAP, BORDER
        ws.column_dimensions[get_column_letter(idx)].width = width

    FIRST, LAST = HEADER_ROW + 1, 500
    for r in range(FIRST, LAST + 1):
        # H 剩余课时 = 总课时 - 已耗课时
        ws.cell(row=r, column=8, value=f'=IF(OR(F{r}="",G{r}=""),"",F{r}-G{r})')
        # L 状态：三条判据按优先级重算（学员名/已耗课时未填则留空）
        ws.cell(row=r, column=12, value=(
            f'=IF(OR($A{r}="",H{r}=""),"",'
            f'IF(AND(ISNUMBER(H{r}),H{r}<=2),"课时告急",'
            f'IF(AND(I{r}<>"",I{r}-TODAY()<=14,I{r}-TODAY()>=0),"将到期",'
            f'IF(AND(J{r}<>"",TODAY()-J{r}>=14),"流失风险",'
            f'"活跃"))))'
        ))
        for col in range(1, n + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDER
            if col in (6, 7, 8):
                cell.number_format = "0"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col in (5, 9, 10):
                cell.number_format = "yyyy-mm-dd"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col in (3, 4, 11, 12):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 13:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 20

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A{HEADER_ROW}:M{LAST}"

    dv = DataValidation(type="list", formula1=f'"{STATES}"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"L{FIRST}:L{LAST}")

    rng = f"A{FIRST}:M{LAST}"
    # 判据一：课时告急（红）——必须 ISNUMBER，否则空行会被 <=2 误判
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'AND($A{FIRST}<>"",ISNUMBER($H{FIRST}),$H{FIRST}<=2)'],
                         fill=ALERT_FILL))
    # 判据二：将到期（黄）——须剩余课时>2（课时告急优先）
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'AND($A{FIRST}<>"",ISNUMBER($H{FIRST}),$H{FIRST}>2,$I{FIRST}<>"",$I{FIRST}-TODAY()<=14,$I{FIRST}-TODAY()>=0)'],
                         fill=SOON_FILL))
    # 判据三：流失风险（蓝）——须未进入前两种
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'AND($A{FIRST}<>"",ISNUMBER($H{FIRST}),$H{FIRST}>2,$I{FIRST}<>"",$I{FIRST}-TODAY()>14,$J{FIRST}<>"",TODAY()-$J{FIRST}>=14)'],
                         fill=RISK_FILL))

    ws2 = wb.create_sheet("怎么填")
    ws2.column_dimensions["A"].width = 20
    ws2.column_dimensions["B"].width = 74
    for r, (k, v) in enumerate(GUIDE, start=1):
        a, b = ws2.cell(row=r, column=1, value=k), ws2.cell(row=r, column=2, value=v)
        if k and not v:
            a.font = TITLE_FONT
        else:
            a.font = Font(bold=True, size=10)
        b.font = Font(size=10)
        b.alignment = Alignment(vertical="center", wrap_text=True)
        ws2.row_dimensions[r].height = 30

    parent = os.path.dirname(os.path.abspath(out_path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT
    p = build(out)
    print("已生成:", p)
    print("大小:", round(os.path.getsize(p) / 1024, 1), "KB")
