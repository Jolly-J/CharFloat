#!/usr/bin/env python3
"""e2e 步骤 7：失败恢复实测 —— 先记录指纹，再打 4 类失败，最后核对状态与磁盘。"""
import hashlib
import json
import os
import subprocess

ROOT = "/Users/Python/Office Agent Bridge"
WB = "钙钛矿各家企业现状.xlsx"
PPT = "e2e-钙钛矿产业现状与展望.pptx"
DOC = "e2e-钙钛矿产业现状说明.docx"
FILES = [os.path.expanduser("~/Downloads/钙钛矿各家企业现状.xlsx"),
         os.path.expanduser("~/Downloads/" + PPT),
         os.path.expanduser("~/Downloads/" + DOC)]


def call(tool, args):
    out = subprocess.run(
        ["python3", "skills/office-agent-bridge/scripts/bridge_client.py", tool, json.dumps(args, ensure_ascii=False)],
        cwd=ROOT, capture_output=True, text=True,
    )
    return (out.stdout.strip() or out.stderr.strip())


def fp():
    r = {}
    for f in FILES:
        h = hashlib.sha256(open(f, "rb").read()).hexdigest()[:12]
        r[os.path.basename(f)] = f"{os.path.getsize(f)}B {h}"
    return r


before = fp()
print("=== 失败注入 ===")
CASES = [
    ("不存在的演示文稿", "wps_ppt_read_presentation", {"presentationName": "e2e-根本不存在.pptx"}),
    ("越界页码删除", "wps_ppt_manage_slides", {"presentationName": PPT, "action": "delete", "slideIndex": 99}),
    ("不存在的 Word 文档写入", "wps_word_write_content",
     {"documentName": "e2e-不存在.docx", "content": "失败探针", "type": "paragraph"}),
    ("不存在的工作表写入（走审计路径）", "wps_patch_cells",
     {"workbookName": WB, "sheetName": "e2e_不存在的表", "address": "A1", "values": [["x"]], "reason": "e2e 失败探针"}),
    ("不存在的 PPT 表格页", "wps_ppt_manage_table",
     {"presentationName": PPT, "slideIndex": 99, "action": "create_table", "rows": 2, "columns": 2}),
]
for label, tool, args in CASES:
    raw = call(tool, args)
    first = raw.replace("\n", " ")[:220]
    print(f"\n[{label}] {tool}\n  {first}")

print("\n=== 失败后内存状态 ===")
st_ppt = json.loads(call("wps_ppt_read_presentation", {"presentationName": PPT, "maxSlides": 20}))["data"]
print(f"演示文稿页数={st_ppt['slideCount']}")
st_doc = json.loads(call("wps_word_read_document", {"documentName": DOC, "scope": "outline"}))["data"]
print(f"文档大纲={[o['text'] for o in st_doc['outline']]}")
st_xl = json.loads(call("wps_read_range", {"workbookName": WB, "sheetName": "e2e_钙钛矿数据集",
                                           "address": "A1:D21", "includeFormulas": False}))["data"]
print(f"数据集行数={len(st_xl['values'])} 首行首列={st_xl['values'][2][:2]}")

print("\n=== 安全重试（先建一页临时页，再删掉）===")
print(" ", call("wps_ppt_manage_slides", {"presentationName": PPT, "action": "add", "layoutIndex": 12})[:160].replace("\n", " "))
print(" ", call("wps_ppt_manage_slides", {"presentationName": PPT, "action": "delete", "slideIndex": 10})[:160].replace("\n", " "))
st_ppt2 = json.loads(call("wps_ppt_read_presentation", {"presentationName": PPT, "maxSlides": 20}))["data"]
print(f"重试后页数={st_ppt2['slideCount']}（应为 9）")

after = fp()
print("\n=== 磁盘指纹 ===")
for k in before:
    print(f"  {k}\n    before={before[k]}\n    after ={after[k]}  {'未变' if before[k]==after[k] else '★变了'}")
