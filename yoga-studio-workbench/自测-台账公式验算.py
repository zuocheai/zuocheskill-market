# -*- coding: utf-8 -*-
"""用真数据验算台账公式：四条边界必须判对，否则表一建就是错的。"""
import datetime
import os
import sys

from openpyxl import load_workbook

XLSX = sys.argv[1] if len(sys.argv) > 1 else "C:/Users/Admin/.workbuddy/scratch/验算台账.xlsx"
TODAY = datetime.date(2026, 10, 5)

CASES = [
    # (学员名, 总课时, 已耗, 到期日, 上次上课, 期望状态, 期望颜色)
    # ★ 判据一是「剩余课时 ≤ 2」，所以总20已耗17=剩3节 → 不是告急（这里专门验这条）
    ("剩3节-不该告急", 20, 17, datetime.date(2027, 6, 1), datetime.date(2026, 9, 30), "活跃", "无"),
    ("剩2节-该告急", 20, 18, datetime.date(2027, 6, 1), datetime.date(2026, 9, 30), "课时告急", "红"),
    ("剩1节-该告急", 12, 11, datetime.date(2027, 5, 1), datetime.date(2026, 10, 4), "课时告急", "红"),
    ("将到期样例", 60, 55, datetime.date(2026, 10, 14), datetime.date(2026, 9, 20), "将到期", "黄"),
    ("流失风险样例", 30, 8, datetime.date(2027, 5, 1), datetime.date(2026, 9, 1), "流失风险", "蓝"),
    ("正常活跃样例", 30, 8, datetime.date(2027, 5, 1), datetime.date(2026, 10, 4), "活跃", "无"),
    ("边界-刚好14天", 30, 8, datetime.date(2026, 10, 19), datetime.date(2026, 10, 4), "将到期", "黄"),
    ("边界-15天", 30, 8, datetime.date(2026, 10, 20), datetime.date(2026, 10, 4), "活跃", "无"),
    ("边界-刚好14天没来", 30, 8, datetime.date(2027, 5, 1), datetime.date(2026, 9, 21), "流失风险", "蓝"),
    ("边界-13天没来", 30, 8, datetime.date(2027, 5, 1), datetime.date(2026, 9, 22), "活跃", "无"),
    ("边界-过期卡", 30, 8, datetime.date(2026, 9, 1), datetime.date(2026, 10, 4), "活跃", "无"),
    ("空行-只填名字", None, None, None, None, "", "无"),
    ("空行-什么都没填", None, None, None, None, "", "无"),
]

wb = load_workbook(XLSX)
ws = wb["学员台账"]

# 写入样例
for i, (name, total, used, exp, last, _, _) in enumerate(CASES):
    r = 4 + i
    ws.cell(row=r, column=1, value=name or None)
    ws.cell(row=r, column=4, value="测试卡")
    ws.cell(row=r, column=6, value=total)
    ws.cell(row=r, column=7, value=used)
    ws.cell(row=r, column=9, value=exp)
    ws.cell(row=r, column=10, value=last)

tmp = XLSX.replace(".xlsx", "_计算.xlsx")
wb.save(tmp)

# 计算引擎认的是临时文件名——从 tmp 推，别写死
BOOK = os.path.splitext(os.path.basename(tmp))[0]

import formulas
xl = formulas.ExcelModel().loads(tmp).finish()
sol = xl.calculate()


# 取 L 列（状态）算出来的值
def get(letter, row):
    want = f"'[{BOOK}.XLSX]学员台账'!{letter}{row}"
    for k, v in sol.items():
        if k.upper().endswith(want.upper()):
            try:
                return v.value[0, 0]
            except Exception:
                return v
    return None

def remain(total, used):
    if total is None or used is None:
        return ""
    return total - used

ok = True
print(f"{'样例':<18}{'剩余':>5}{'期望状态':>10}{'实算':>10}  {'判':<4}{'期望色':>6}{'实判色':>6}")
print("-" * 62)
for i, (name, total, used, exp, last, want, want_color) in enumerate(CASES):
    r = 4 + i
    rem = remain(total, used)
    got = get("L", r)
    got = "" if got is None else str(got)
    passed = (got == want)
    ok = ok and passed
    print(f"{name:<18}{str(rem):>5}{want:>10}{got:>10}  {'OK' if passed else 'FAIL':<4}{want_color:>6}{'-':>6}")

print("-" * 62)
print("总判定:", "全部通过 ✅" if ok else "有失败 ❌")
os.remove(tmp)
sys.exit(0 if ok else 1)
