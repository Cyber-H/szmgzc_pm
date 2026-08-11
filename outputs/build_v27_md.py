# -*- coding: utf-8 -*-
"""Rebuild 数媒智创平台_PRD_V2.7.md with corrected chapter-4 architecture.
Source: V2.6.md.  Scan-confirmed sidebar (2026-08-11):
  顶层菜单: 新建任务 / 知识库 / 热点助手 / 生产工具 / 技能市场
  热点榜单 = 热点助手内「国内热榜/国际热榜」页面 (非独立菜单)
  历史记录 = 新建任务首页组件 (非独立菜单)
  个人中心 = 右上角头像菜单
"""
import io, os, re

SRC = r"C:\Users\chenhong\WorkBuddy\2026-06-26-14-49-33\szmgzc_pm\docs\数媒智创平台_PRD_V2.6.md"
DST = r"C:\Users\chenhong\WorkBuddy\2026-06-26-14-49-33\szmgzc_pm\docs\数媒智创平台_PRD_V2.7.md"

with io.open(SRC, encoding="utf-8") as f:
    txt = f.read()

def idx(prefix):
    i = txt.find(prefix)
    if i < 0:
        raise SystemExit("anchor not found: %r" % prefix)
    return i

# anchors
A41 = "4.1 新建任务（AI对话工作台）"
A42 = "4.2 知识库"
A43 = "4.3 热点榜单"
A44 = "4.4 热点助手"
A45 = "4.5 生产工具"
A46 = "4.6 技能市场（SkillHub）"
A47 = "4.7 历史记录（原\"历史对话\"）"
A48 = "4.8 个人中心"
A5  = "5. 典型业务流程"

i41=idx(A41); i42=idx(A42); i43=idx(A43); i44=idx(A44)
i45=idx(A45); i46=idx(A46); i47=idx(A47); i48=idx(A48); i5=idx(A5)

pre = txt[:i41]                      # everything before chapter 4
block41 = txt[i41:i42]               # 4.1 新建任务
block42 = txt[i42:i43]               # 4.2 知识库
block43 = txt[i43:i44]               # OLD 4.3 热点榜单 (deleted as standalone; content reused)
block44 = txt[i44:i45]               # OLD 4.4 热点助手
block45 = txt[i45:i46]               # OLD 4.5 生产工具
block46 = txt[i46:i47]               # OLD 4.6 技能市场
block47 = txt[i47:i48]               # OLD 4.7 历史记录 -> fold into 4.1.5
block48 = txt[i48:i5]                # OLD 4.8 个人中心
post = txt[i5:]                      # chapter 5 onward

def renum_block(s, mapping):
    # mapping: old->new, apply longer first to avoid partial collisions
    for old, new in sorted(mapping.items(), key=lambda kv: -len(kv[0])):
        s = s.replace(old, new)
    return s

# ---- block44 (OLD 热点助手) -> new 4.3, drop 4.4.6, split before 4.4.5 ----
split45 = block44.find("4.4.5 事件详情")
head44 = block44[:split45]            # 4.4.1..4.4.4
tail44 = block44[split45:]           # 4.4.5 事件详情 ... 4.4.6 与热点榜单的关系 ...
# remove 4.4.6 section from tail44
end46 = tail44.find("4.5 生产工具")   # not present in block44 (block ends before 4.5) -> guard
# 4.4.6 is the last subsection inside block44; cut everything from "4.4.6 与热点榜单的关系"
pos46 = tail44.find("4.4.6 与热点榜单的关系")
if pos46 >= 0:
    tail44 = tail44[:pos46].rstrip() + "\n"
# renumber head44 4.4.{1..4} -> 4.3.{1..4} and heading
head44 = head44.replace("4.4 热点助手", "4.3 热点助手")
head44 = renum_block(head44, {"4.4.4":"4.3.4","4.4.3":"4.3.3","4.4.2":"4.3.2","4.4.1":"4.3.1"})
tail44 = tail44.replace("4.4.5 事件详情", "4.3.6 事件详情")

# new 4.3.1 功能定位 (rewrite)
new_intro = (
"4.3.1 功能定位\n"
"\n"
"热点助手是基于热点数据的进阶热点分析工具，面向编辑、策划与运营人员。模块由左上角并列的 4 个页面组成：热点搜索、热点预测、国内热榜、国际热榜。其中「国内热榜 / 国际热榜」两个页面即由原独立的「热点榜单」模块合并而来（原 /hotspot 路径重定向至 /hotspot-assistant?tab=domestic / ?tab=international），提供国内外多平台实时热搜聚合；热点搜索、热点预测与事件详情则在聚合之上提供检索、预测与深度分析能力，强调“查热点、跟热点、分析热点”，支持按行业、地域、时间等维度检索与预测热点事件。\n"
)

# replace old head44 first subsection (4.3.1 功能定位 ... up to 4.3.2 页面结构) with new_intro + 4.3.2
pos_432 = head44.find("4.3.2 页面结构")
head44_new = "4.3 热点助手\n\n" + new_intro + head44[pos_432:]

# ---- former 热点榜单 content -> new 4.3.5 国内热榜与国际热榜 ----
# take from "4.3.2 内容结构" to end of block43, renumber labels
pos_cstruct = block43.find("4.3.2 内容结构")
hotlist_body = block43[pos_cstruct:]
hotlist_body = hotlist_body.replace("4.3.2 内容结构", "内容结构").replace("4.3.3 交互说明", "交互说明")
# terminology normalization: 导航单元统一称"页面"
hotlist_body = hotlist_body.replace("\nTab\n\n覆盖平台", "\n页面\n\n覆盖平台")
hotlist_body = hotlist_body.replace("国内/国际为两个独立 Tab", "国内/国际为两个独立页面")
# drop the stray "4.3.3 交互说明" duplicate if any & trim leading blank
hotlist_section = (
"4.3.5 国内热榜与国际热榜\n"
"\n"
"国内热榜与国际热榜两个页面即由原独立的「热点榜单」模块合并而来，面向编辑与运营人员提供国内外多平台实时热搜聚合，支撑选题发现与热点跟进。\n"
"\n"
) + hotlist_body.strip() + "\n"

# ---- block45 (OLD 生产工具) -> 4.4 ----
block45_new = block45.replace("4.5 生产工具", "4.4 生产工具")
block45_new = renum_block(block45_new, {"4.5.5":"4.4.5","4.5.4":"4.4.4","4.5.3":"4.4.3","4.5.2":"4.4.2","4.5.1":"4.4.1"})

# ---- block46 (OLD 技能市场) -> 4.5 ----
block46_new = block46.replace("4.6 技能市场（SkillHub）", "4.5 技能市场（SkillHub）")
block46_new = renum_block(block46_new, {"4.6.3":"4.5.3","4.6.2":"4.5.2","4.6.1":"4.5.1"})

# ---- block47 (OLD 历史记录) -> 4.1.5 ----
block47_new = block47.replace("4.7 历史记录（原\"历史对话\"）", "4.1.5 历史记录（原\"历史对话\"）")
block47_new = block47_new.replace("4.7.1 功能结构", "功能结构").replace("4.7.2 交互说明", "交互说明")

# ---- block48 (OLD 个人中心) -> 4.6 ----
block48_new = block48.replace("4.8 个人中心", "4.6 个人中心")

# ---- assemble chapter 4 ----
chapter4 = (
"4. 功能模块详细说明\n\n"
+ block41.rstrip() + "\n\n"
+ block47_new.rstrip() + "\n\n"        # 4.1.5 历史记录 (folded into 新建任务)
+ block42.rstrip() + "\n\n"             # 4.2 知识库
+ head44_new.rstrip() + "\n\n"
+ hotlist_section.rstrip() + "\n\n"
+ tail44.rstrip() + "\n\n"              # 4.3.6 事件详情
+ block45_new.rstrip() + "\n\n"         # 4.4 生产工具
+ block46_new.rstrip() + "\n\n"         # 4.5 技能市场
+ block48_new.rstrip() + "\n\n"         # 4.6 个人中心
+ post.lstrip()
)

# ---- global fixes (version, 功能架构表, 8.1, 5.1, 6.x) ----
full = pre.rstrip() + "\n\n" + chapter4

# version
full = full.replace("文档版本：V2.6", "文档版本：V2.7", 1)

# 功能架构表: remove standalone 热点榜单 row + update 热点助手 row
hotlist_row = ("热点榜单（已并入热点助手）\n\n国内热榜 / 国际热榜\n\n"
               "多平台热搜聚合、Top10展示、刷新（作为热点助手「国内热榜/国际热榜」页面）\n\n"
               "百度、腾讯、澎湃、抖音等（国内16）+ 参考消息/华尔街见闻/卫星通讯社（国际5）\n\n")
if hotlist_row in full:
    full = full.replace(hotlist_row, "")
else:
    # fallback: remove just the heading line block
    full = full.replace("热点榜单（已并入热点助手）\n\n国内热榜 / 国际热榜\n\n", "")
full = full.replace(
    "左上角 4 个页面导航，原热点榜单现为其「国内热榜/国际热榜」页面",
    "左上角 4 个页面导航；其中国内热榜/国际热榜即原独立的「热点榜单」模块（国内16平台+国际5平台）")

# 8.1: remove 热点榜单 row + relabel 热点助手
hl81 = ("热点榜单（已并入热点助手，/hotspot 重定向）\n\n"
        "https://x.sztv.com.cn/hotspot-assistant?tab=domestic（国内热榜） / ?tab=international（国际热榜）\n\n")
full = full.replace(hl81, "")
full = full.replace("热点助手（含左上角 4 个页面）",
                    "热点助手（4 个页面：热点搜索/热点预测/国内热榜/国际热榜）")

# 5.1 flow
full = full.replace("切换「国内热榜 / 国际热榜」Tab", "切换「国内热榜 / 国际热榜」页面")

# 6. 非功能需求
full = full.replace("热点榜单、历史记录列表", "国内热榜/国际热榜、历史记录列表")

# revision row V2.7
v27 = ("V2.7\n\n2026-08-11\n\n"
       "按重新登录扫描的真实左侧导航架构重写第4章：移除独立的「热点榜单」模块章节，将其内容合并为热点助手的「国内热榜/国际热榜」两个页面；热点助手调整为第4.3章并由左上角4个页面（热点搜索/热点预测/国内热榜/国际热榜）组成；「历史记录」由独立章节归并至4.1新建任务的4.1.5；后续章节重新编号（生产工具→4.4、技能市场→4.5、个人中心→4.6）；同步修订功能架构表、8.1附录、业务流程与非功能需求中的章节编号引用。\n\n阿爪\n\n")
sec2 = "2. 产品概述"
if sec2 in full:
    full = full.replace(sec2, v27 + sec2, 1)

with io.open(DST, "w", encoding="utf-8") as f:
    f.write(full)

print("V2.7.md written ->", DST)
