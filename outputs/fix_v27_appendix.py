# -*- coding: utf-8 -*-
"""Fix 8.1 appendix table in V2.7.docx: remove stale 热点榜单 row, add 热点助手 row."""
from docx import Document
from docx.oxml.ns import qn

P = r"C:\Users\chenhong\WorkBuddy\2026-06-26-14-49-33\szmgzc_pm\数媒智创平台_PRD_V2.7.docx"
doc = Document(P)

target = None
for tb in doc.tables:
    hdr = " ".join(c.text for r in tb.rows[:1] for c in r.cells)
    if hdr.startswith("页面") and "URL" in hdr:
        target = tb
        break
if target is None:
    raise SystemExit("appendix table not found")

# 1) delete row whose first cell == 热点榜单
for ri in range(len(target.rows) - 1, 0, -1):
    if target.rows[ri].cells[0].text.strip() == "热点榜单":
        target._tbl.remove(target.rows[ri]._tr)
        print("removed 热点榜单 row at", ri)
        break

# 2) insert 热点助手 row after 主站 (row index 1)
new_row = target.add_row()
master_tr = target.rows[1]._tr
master_tr.addnext(new_row._tr)
new_row.cells[0].text = "热点助手（4 个页面：热点搜索/热点预测/国内热榜/国际热榜）"
new_row.cells[1].text = ("https://x.sztv.com.cn/hotspot-assistant?tab=search（搜索）/ ?tab=predict（预测）/ "
                          "?tab=domestic（国内热榜）/ ?tab=international（国际热榜）")
print("inserted 热点助手 row after 主站")

doc.save(P)
print("saved", P)
