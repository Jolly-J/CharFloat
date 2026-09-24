#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
片 A（无插件）道具生成器：同一批作品的"云端 AI 返回版"。

片 B 里 AI 在 WPS 里就地做出三份精修成品；片 A 里同一件事交给云端 AI，
拿回来的是劣化产物。两条片子做的是同一件事，对照才成立。

本脚本从 promo/成品基线/ 的脱敏成品出发，生成：

  02-上传给AI的文件/   真加密的成品（AI 读不到字节）
  03-AI返回的文件/     劣化版 + 版本堆积

劣化规则（模拟"云端重新生成一份"）：
  xlsx：241 条公式**全部算成死值**、填色/边框/字体/合并全丢、多出空表
  pptx：原生图表换成图片、卡片错位、字号不统一
  docx：表格边框与底纹丢失、标题层级退化成普通段落

全部数据沿用脱敏成品里的虚构值，不含任何真实数据。
"""
from __future__ import annotations

import copy
import io
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

PROMO = Path(__file__).resolve().parent.parent
FINALS = PROMO / "成品基线"
ENV_A = PROMO / "env-A-无插件"
DEMO_PASSWORD = "Demo#2026"      # 演示口令，非真实凭据

SRC_XLSX = FINALS / "作品一_成本测算评审_脱敏.xlsx"
SRC_PPTX = FINALS / "作品二_五年业务规划_脱敏.pptx"
SRC_DOCX = FINALS / "作品三_经营周报_脱敏.docx"

CELL_RE = re.compile(r"(?:'[^']+'|[A-Za-z_\u4e00-\u9fff][\w.\u4e00-\u9fff]*)!\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}\$?\d+")
SUM_RE = re.compile(r"SUM\(([^()]+)\)", re.I)


# ---------------------------------------------------------------------------
# 一、公式求值器：把这 12 种形态算成数字（用于"公式变死值"）
# ---------------------------------------------------------------------------
class Evaluator:
    def __init__(self, wb):
        self.wb = wb
        self.cache: dict[tuple[str, str], float] = {}

    def _num(self, cell) -> float:
        v = cell.value
        if v is None:
            return 0.0
        if isinstance(v, bool):
            return float(v)
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str) and v.startswith("="):
            return self.sheet_value(cell.parent, v[1:])
        return 0.0

    def _range_sum(self, ws, ref: str) -> float:
        if "!" in ref:
            sheet, ref = ref.split("!", 1)
            ws = self.wb[sheet.strip("'")]
        ref = ref.replace("$", "")
        total = 0.0
        for row in ws[ref]:
            if isinstance(row, tuple):
                total += sum(self._num(c) for c in row)
            else:
                total += self._num(row)
        return total

    def sheet_value(self, ws, formula: str, seen=None) -> float:
        seen = set() if seen is None else seen
        head, _sep, _tail = formula.partition("(")
        expr = formula

        def repl_sum(m):
            return f"({self._range_sum(ws, m.group(1))!r})"

        for _ in range(8):                      # SUM 可以嵌套在别的函数里，多跑几轮
            new = SUM_RE.sub(repl_sum, expr)
            if new == expr:
                break
            expr = new

        def repl_ref(m):
            ref = m.group(0)
            target_ws = ws
            if "!" in ref:
                sheet, ref = ref.split("!", 1)
                target_ws = self.wb[sheet.strip("'")]
            key = (target_ws.title, ref.replace("$", ""))
            if key in seen:
                raise ValueError(f"循环引用 {key}")
            if key in self.cache:
                return repr(self.cache[key])
            cell = target_ws[ref.replace("$", "")]
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                seen.add(key)
                val = self.sheet_value(target_ws, v[1:], seen)
                seen.discard(key)
                self.cache[key] = val
                return repr(val)
            return repr(self._num(cell))

        expr = CELL_RE.sub(repl_ref, expr)
        if not re.fullmatch(r"[0-9eE+\-*/(). ]+", expr):
            raise ValueError(f"求值器不支持的公式：{formula}")
        return eval(expr, {"__builtins__": {}}, {})          # noqa: S307 —— 已用白名单正则限死字符集


# ---------------------------------------------------------------------------
# 二、作品一劣化版
# ---------------------------------------------------------------------------
def degrade_xlsx(dst: Path, *, variant: str = "v1", junk_sheet: bool = False,
                 ai_palette: bool = False, title_suffix: str = "") -> dict:
    src = openpyxl.load_workbook(SRC_XLSX)
    ev = Evaluator(src)

    out = openpyxl.Workbook()
    out.remove(out.active)
    stats = {"cells": 0, "formulas": 0, "dropped_merges": 0}

    for ws in src.worksheets:
        nws = out.create_sheet(ws.title)
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    # 关键劣化点：公式 -> 死数字
                    try:
                        v = round(ev.sheet_value(ws, v[1:]), 6)
                    except Exception as e:                      # noqa: BLE001
                        print(f"      ! 求值失败 {ws.title}!{c.coordinate} {c.value}: {e}")
                        v = 0
                    stats["formulas"] += 1
                nws.cell(row=c.row, column=c.column, value=v)
                stats["cells"] += 1
        # 只保留"AI 自己上的色"，你的填色/边框/字体/合并全部丢失
        if ai_palette:
            for c in nws[1]:
                if c.value:
                    c.font = Font(bold=True, color="FFFFFF")
                    c.fill = PatternFill("solid", fgColor="ED7D31")
        stats["dropped_merges"] += len(ws.merged_cells.ranges)
        nws.column_dimensions["A"].width = 21.0

    if title_suffix:
        first = out.worksheets[0]
        first["A1"] = f"{first['A1'].value}{title_suffix}"

    if junk_sheet:
        out.create_sheet("Sheet1")          # 云端重造常见的"顺手多一个空表"

    out.save(dst)
    return stats


# ---------------------------------------------------------------------------
# 三、作品二劣化版：原生图表换图片 + 卡片错位
# ---------------------------------------------------------------------------
def _chart_png(path: Path, idx: int) -> None:
    from PIL import Image, ImageDraw, ImageFont
    font_path = None
    for p in ("/System/Library/Fonts/STHeiti Light.ttc", "/System/Library/Fonts/Hiragino Sans GB.ttc"):
        if Path(p).exists():
            font_path = p
            break

    def F(size):
        return ImageFont.truetype(font_path, size) if font_path else ImageFont.load_default()

    img = Image.new("RGB", (900, 420), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 900, 60], fill="#ED7D31")
    d.text((24, 16), f"业务目标趋势图 {idx}", font=F(28), fill="white")
    base_y = 360
    d.line([(80, base_y), (860, base_y)], fill="#999999", width=2)
    vals = [0.35, 0.55, 0.72, 0.9]
    for i, v in enumerate(vals):
        x = 120 + i * 190
        h = int(v * 250)
        d.rectangle([x, base_y - h, x + 90, base_y], fill="#ED7D31")
        d.text((x + 10, base_y + 10), f"{2026 + i}", font=F(22), fill="#555555")
    d.text((720, 70), "图片\n不可编辑", font=F(20), fill="#BBBBBB")
    img.save(path, "PNG")


def degrade_pptx(dst: Path, tmp: Path) -> dict:
    """把原生图表换成图片。用 lxml 改 XML，别用字符串切片——切歪了就是非法 XML。"""
    from lxml import etree

    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    for i in range(1, 5):
        _chart_png(tmp / f"chart_{i}.png", i)

    with zipfile.ZipFile(SRC_PPTX) as z:
        part_map = {n: z.read(n) for n in z.namelist()}

    # 1) 删掉原生图表部件、其 rels，以及内嵌数据源
    killed = 0
    for n in list(part_map):
        if (re.fullmatch(r"ppt/charts/chart\d+\.xml", n)
                or re.fullmatch(r"ppt/charts/_rels/chart\d+\.xml\.rels", n)
                or n.startswith("ppt/embeddings/")):
            del part_map[n]
            killed += 1

    # 2) 幻灯片 rels：把 chart 关系就地改成 image，rId 保持不变，pic 才能直接复用
    rels_name = "ppt/slides/_rels/slide1.xml.rels"
    rels = etree.fromstring(part_map[rels_name])
    chart_rids = []
    for rel in rels:
        if str(rel.get("Type", "")).endswith("/chart"):
            chart_rids.append(rel.get("Id"))
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            rel.set("Target", f"../media/ai_chart_{len(chart_rids)}.png")
    part_map[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    for i in range(1, len(chart_rids) + 1):
        part_map[f"ppt/media/ai_chart_{i}.png"] = (tmp / f"chart_{i}.png").read_bytes()

    # 3) 幻灯片：graphicFrame(图表) 整块换成 pic
    slide = etree.fromstring(part_map["ppt/slides/slide1.xml"])
    replaced = 0
    for gf in slide.iter(f"{P}graphicFrame"):
        chart = gf.find(f".//{C}chart")
        if chart is None:
            continue
        rid = chart.get(f"{R}id")
        xfrm = gf.find(f".//{A}xfrm")
        x = y = "60"
        cx, cy = "4000000", "2200000"
        if xfrm is not None:
            off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
            if off is not None:
                x, y = off.get("x"), off.get("y")
            if ext is not None:
                cx, cy = ext.get("cx"), ext.get("cy")
        replaced += 1
        pic = etree.SubElement(gf.getparent(), f"{P}pic")
        nv = etree.SubElement(pic, f"{P}nvPicPr")
        cNvPr = etree.SubElement(nv, f"{P}cNvPr")
        cNvPr.set("id", str(900 + replaced))
        cNvPr.set("name", f"AI图表{replaced}.png")
        etree.SubElement(nv, f"{P}cNvPicPr")
        etree.SubElement(nv, f"{P}nvPr")
        blipFill = etree.SubElement(pic, f"{P}blipFill")
        blip = etree.SubElement(blipFill, f"{A}blip")
        blip.set(f"{R}embed", rid or f"rId{replaced}")
        stretch = etree.SubElement(blipFill, f"{A}stretch")
        etree.SubElement(stretch, f"{A}fillRect")
        spPr = etree.SubElement(pic, f"{P}spPr")
        xf = etree.SubElement(spPr, f"{A}xfrm")
        off = etree.SubElement(xf, f"{A}off"); off.set("x", x); off.set("y", y)
        ext = etree.SubElement(xf, f"{A}ext"); ext.set("cx", cx); ext.set("cy", cy)
        prst = etree.SubElement(spPr, f"{A}prstGeom"); prst.set("prst", "rect")
        etree.SubElement(prst, f"{A}avLst")
        gf.getparent().remove(gf)
    part_map["ppt/slides/slide1.xml"] = etree.tostring(slide, xml_declaration=True,
                                                       encoding="UTF-8", standalone=True)

    # 4) 卡片位置打散：每个 <a:off> 加一点固定抖动
    slide2 = etree.fromstring(part_map["ppt/slides/slide1.xml"])
    jittered = 0
    for off in slide2.iter(f"{A}off"):
        jittered += 1
        try:
            off.set("x", str(int(off.get("x")) + 13 * (jittered % 5 - 2)))
            off.set("y", str(int(off.get("y")) + 9 * ((jittered + 2) % 4 - 1)))
        except (TypeError, ValueError):
            pass
    part_map["ppt/slides/slide1.xml"] = etree.tostring(slide2, xml_declaration=True,
                                                       encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, raw in part_map.items():
            zout.writestr(n, raw)
    return {"charts_replaced": killed, "pics_inserted": replaced, "shapes_jittered": jittered}


# ---------------------------------------------------------------------------
# 四、作品三劣化版：表格边框/底纹丢失、标题层级退化
# ---------------------------------------------------------------------------
def degrade_docx(dst: Path) -> dict:
    stats = {"tables_stripped": 0, "styles_flattened": 0}
    with zipfile.ZipFile(SRC_DOCX) as z:
        part_map = {n: z.read(n) for n in z.namelist()}

    doc = part_map["word/document.xml"].decode("utf-8")
    # 表格边框与底纹全部去掉（云端重造最常见的丢失）
    stats["tables_stripped"] = len(re.findall(r"<w:tblBorders>", doc))
    doc = re.sub(r"<w:tblBorders>.*?</w:tblBorders>", "", doc, flags=re.S)
    doc = re.sub(r"<w:shd [^/>]*/>", "", doc)
    # 这份文档的排版是"直接格式"（w:rPr / w:pPr，没有 pStyle）：
    # 按 pStyle 找是找不到的——要把直接格式剥掉，段落才会退回默认样子。
    stats["styles_flattened"] = len(re.findall(r"<w:pPr[ >/]", doc))
    doc = re.sub(r"<w:pPr\s*/>", "", doc)
    doc = re.sub(r"<w:pPr[ >].*?</w:pPr>", "", doc, flags=re.S)
    doc = re.sub(r"<w:rPr\s*/>", "", doc)
    doc = re.sub(r"<w:rPr[ >].*?</w:rPr>", "", doc, flags=re.S)
    # 表格列宽也一并去掉：列宽塌成默认，一眼看出"不是原来那份"
    doc = re.sub(r"<w:tcW [^/>]*/>", "", doc)
    doc = re.sub(r"<w:tblW [^/>]*/>", "", doc)
    part_map["word/document.xml"] = doc.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, raw in part_map.items():
            zout.writestr(n, raw)
    return stats


# ---------------------------------------------------------------------------
# 五、加密道具：同一个任务，文件加密，云端读不到
# ---------------------------------------------------------------------------
def build_encrypted(stage: Path) -> dict:
    from msoffcrypto.format.ooxml import OOXMLFile

    plain = stage / "成本测算评审_源数据.xlsx"
    shutil.copyfile(SRC_XLSX, plain)

    out_dir = ENV_A / "02-上传给AI的文件"
    out_dir.mkdir(parents=True, exist_ok=True)
    enc = out_dir / "成本测算评审_机密.xlsx"
    with open(plain, "rb") as fin, open(enc, "wb") as fout:
        OOXMLFile(fin).encrypt(DEMO_PASSWORD, fout)

    zpath = out_dir / "成本测算评审_机密.zip"
    if zpath.exists():
        zpath.unlink()
    subprocess.run(["zip", "-q", "-j", "-P", DEMO_PASSWORD, str(zpath), str(plain)], check=True)
    return {"加密xlsx": enc.name, "加密zip": zpath.name}


# ---------------------------------------------------------------------------
def main() -> int:
    tmp = ENV_A / ".build"
    tmp.mkdir(parents=True, exist_ok=True)
    ret = ENV_A / "03-AI返回的文件"
    ret.mkdir(parents=True, exist_ok=True)
    for old in ret.glob("*"):
        old.unlink()

    print("== 作品一 · 成本测算评审（公式变死值）==")
    variants = [
        ("成本测算评审_AI分析版.xlsx", dict(variant="v1")),
        ("成本测算评审_v2.xlsx", dict(variant="v2", ai_palette=True)),
        ("成本测算评审_最终版.xlsx", dict(variant="v3", junk_sheet=True)),
        ("成本测算评审_最终版_真的最终版.xlsx", dict(variant="v4", ai_palette=True,
                                                  junk_sheet=True, title_suffix="（AI 生成）")),
    ]
    for name, kw in variants:
        st = degrade_xlsx(ret / name, **kw)
        print(f"  {name:<38} 公式→死值 {st['formulas']:>3} 条 · 丢弃合并 {st['dropped_merges']:>2} 处")

    print("== 作品二 · 五年业务规划（原生图表→图片）==")
    st = degrade_pptx(ret / "五年业务规划_AI版.pptx", tmp)
    print(f"  图表替换 {st['charts_replaced']} 个部件 · 错位形状 {st['shapes_jittered']} 处")

    print("== 作品三 · 经营周报（表格样式丢失）==")
    st = degrade_docx(ret / "经营周报_AI版.docx")
    print(f"  去掉表格边框 {st['tables_stripped']} 处 · 退化段落样式 {st['styles_flattened']} 处")

    print("== 加密道具（同一个任务，文件加密）==")
    info = build_encrypted(tmp)
    for k, v in info.items():
        print(f"  {k}: {v}")

    shutil.rmtree(tmp, ignore_errors=True)
    print("\n完成。片 A 道具全部基于作品一/二/三生成，数据沿用脱敏值。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
