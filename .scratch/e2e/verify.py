#!/usr/bin/env python3
"""e2e 步骤 6：三份成品读回 + 跨组件一致性核对。"""
import json
import subprocess

ROOT = "/Users/Python/Office Agent Bridge"
WB = "钙钛矿各家企业现状.xlsx"
PPT = "e2e-钙钛矿产业现状与展望.pptx"
DOC = "e2e-钙钛矿产业现状说明.docx"


def call(tool, args):
    out = subprocess.run(
        ["python3", "skills/office-agent-bridge/scripts/bridge_client.py", tool, json.dumps(args, ensure_ascii=False)],
        cwd=ROOT, capture_output=True, text=True,
    )
    text = out.stdout.strip() or out.stderr.strip()
    try:
        return json.loads(text).get("data")
    except Exception:
        return {"__error__": text[:300]}


# ---- 1. 表格：读回自己的数据集表 ----
xl = call("wps_read_range", {"workbookName": WB, "sheetName": "e2e_钙钛矿数据集",
                             "address": "A1:D21", "includeFormulas": False})
sheet_rows = xl["values"]
sheet = {}
for r in sheet_rows:
    if r and r[0] and r[1] is not None:
        sheet[str(r[0])] = r[1]

# ---- 2. 演示文稿：把全部文字拼起来 ----
ppt = call("wps_ppt_read_presentation", {"presentationName": PPT, "maxSlides": 20})
ppt_text = "\n".join(
    " ".join((s.get("textSnippets") or []) + ([s.get("title")] if s.get("title") else []))
    for s in ppt["slides"])

# ---- 3. 文档：表格 + 段落文字 ----
tb = call("wps_word_manage_table", {"documentName": DOC, "action": "inspect", "tableIndex": 1})
doc_rows = (tb.get("data") or [])[1:]
DOC_ORDER = ["企业总数", "中国", "海外", "专业钙钛矿", "传统光伏", "设备", "材料", "跨界龙头",
             "上游资源", "太空光伏", "江苏省", "广东省", "产能_2025", "产能_2030", "产能_2035",
             "市场需求_2030", "市场需求_2035", "叠层组件最高效率"]
# 文档表按行序与数据集前 18 项对齐（文档表第 20 行为成本目标，单独在下方核对）
doc_tbl = {DOC_ORDER[i]: str(doc_rows[i][1]) for i in range(min(len(DOC_ORDER), len(doc_rows)))}
doc = call("wps_word_read_document", {"documentName": DOC, "scope": "full", "maxParagraphs": 130,
                                     "includeTables": False, "includeFormatting": False})
doc_text = "\n".join((p.get("text") or "") for p in (doc.get("paragraphDetails") or []))

CHECKS = DOC_ORDER  # 18 项，按同一顺序核对

print(f"{'指标':<20}{'表格(源)':<12}{'文档表格':<12}{'演示文稿文字':<14}判定")
bad = 0
for key in CHECKS:
    v_sheet = sheet.get(key)
    v_doc = doc_tbl.get(key)
    v_ppt = "命中" if str(v_sheet) in ppt_text else "缺失"
    ok = str(v_sheet) == str(v_doc) and v_ppt == "命中"
    if not ok:
        bad += 1
    print(f"{key:<20}{str(v_sheet):<12}{str(v_doc):<12}{v_ppt:<14}{'一致' if ok else '★不一致'}")

tot = sum(int(sheet[k]) for k in ["专业钙钛矿", "传统光伏", "设备", "材料", "跨界龙头", "上游资源", "太空光伏"])
print(f"\n类型分项合计 = {tot}（应等于企业总数 {sheet.get('企业总数')}）-> {'OK' if tot == int(sheet['企业总数']) else '★不符'}")
print(f"国家分项合计 = {int(sheet['中国']) + int(sheet['海外'])} -> {'OK' if int(sheet['中国']) + int(sheet['海外']) == int(sheet['企业总数']) else '★不符'}")

print("\n--- 演示文稿结构 ---")
print(f"页数={ppt['slideCount']}  各页标题:")
for s in ppt["slides"]:
    snip = (s.get("textSnippets") or [])
    print(f"  P{s['index']} shapes={s['shapeCount']} :: {(snip[0] if snip else '(无文字)')[:60]}")

print("\n--- 文档结构 ---")
print("大纲:", [f"L{o['level']}:{o['text']}" for o in doc.get("outline") or []])
print(f"段落数={doc.get('paragraphCount')} 表格={doc.get('tableCount')} 字数={doc.get('wordCount')}")
print(f"\n不一致项 = {bad}")
