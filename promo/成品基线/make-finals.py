#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
三个成品的基线文件（提前定稿，供审样与片尾镜头）。

为什么要先有成品：录制时才不会有临场发挥。施工脚本 plan.json 的每一步，都是为了让
WPS 里的文件一步一步长成这三份基线文件的样子。数据口径与 plan.json、官网双轨动画完全一致。

产物（promo/成品基线/）：
  月度经营分析报表_成品.xlsx      作品一 · 表格
  2026年三季度经营汇报_成品.pptx  作品二 · PPT
  2026年三季度经营分析报告_成品.docx 作品三 · 文档

数据全部为虚构演示数据。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt as PptPt

# ---------------------------------------------------------------------------
# 统一口径
# ---------------------------------------------------------------------------
MONTHS = ["7 月", "8 月", "9 月"]
REVENUE = [86, 94, 108]
COST = [66, 72, 80]
PROFIT = [20, 22, 28]
MARGIN = ["23.3%", "23.4%", "25.9%"]
TITLE = "2026 年 7–9 月经营分析（内部）"
CONCLUSION = "结论：9 月营收环比 +14.9%，利润率升至 25.9%；运营成本连续三月上升，采购与渠道费用需复核。"

BLUE_DEEP = "1F3864"
BLUE_HEAD = "2F5597"
BLUE_MAX = "2563EB"
BLUE_MIN = "DBEAFE"
AMBER_MAX = "F59E0B"
AMBER_MIN = "FEF9C3"
GREY_BG = "F1F5F9"
BORDER = "D6DCE4"
PPT_BLUE = PptRGB(0x0F, 0x4C, 0x81)
PPT_INK = PptRGB(0x1F, 0x29, 0x37)
PPT_GREY = PptRGB(0x64, 0x74, 0x8B)
PPT_AMBER = PptRGB(0xF5, 0x9E, 0x0B)

OUT = Path(__file__).resolve().parent


def thin(color=BORDER) -> Border:
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


# ---------------------------------------------------------------------------
# 作品一 · 表格
# ---------------------------------------------------------------------------
def build_sheet(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "月度经营分析"

    ws.merge_cells("A1:D1")
    ws["A1"] = TITLE
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE_DEEP)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32

    for i, text in enumerate(["经营指标 / 万元", *MONTHS], start=1):
        c = ws.cell(row=2, column=i, value=text)
        c.font = Font(bold=True, size=11, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE_HEAD)
        c.alignment = Alignment(horizontal="center" if i > 1 else "left", vertical="center")
    ws.row_dimensions[2].height = 22

    labels = ["营业收入", "运营成本", "营业利润", "利润率"]
    for r, name in enumerate(labels, start=3):
        c = ws.cell(row=r, column=1, value=name)
        c.font = Font(bold=name in ("营业利润", "利润率"))

    for col in range(2, 5):
        L = get_column_letter(col)
        if col == 2 or col == 3:
            ws.cell(row=3, column=col, value=REVENUE[col - 2]).number_format = "#,##0"
            ws.cell(row=4, column=col, value=COST[col - 2]).number_format = "#,##0"
        else:
            # 9 月列与施工脚本 sheet-03 完全一致：跨表动态汇总，不是死数字
            ws.cell(row=3, column=col, value="=SUM('9月明细'!B5:B8)").number_format = "#,##0"
            ws.cell(row=4, column=col, value="=SUM('9月明细'!B12:B15)").number_format = "#,##0"
        ws.cell(row=5, column=col, value=f"={L}3-{L}4").number_format = "#,##0"
        ws.cell(row=6, column=col, value=f"={L}5/{L}3").number_format = "0.0%"

    for r in range(2, 7):
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = thin()
        ws.row_dimensions[r].height = 20
    ws["C3"].comment = Comment("8 月数据已与财务对账一致，勿改。", "财务-林")

    ws.conditional_formatting.add(
        "B3:D3", ColorScaleRule(start_type="min", start_color=BLUE_MIN, end_type="max", end_color=BLUE_MAX)
    )
    ws.conditional_formatting.add(
        "B4:D4", ColorScaleRule(start_type="min", start_color=AMBER_MIN, end_type="max", end_color=AMBER_MAX)
    )

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "营收与成本趋势（万元）"
    chart.y_axis.title = None
    chart.x_axis.title = None
    data = Reference(ws, min_col=1, max_col=4, min_row=3, max_row=4)
    chart.add_data(data, titles_from_data=True, from_rows=True)
    chart.set_categories(Reference(ws, min_col=2, max_col=4, min_row=2))
    chart.width, chart.height = 15.2, 8.8      # cm，对应施工脚本里的 430pt × 250pt
    chart.legend.position = "b"
    chart.dLbls = None
    ws.add_chart(chart, "F2")

    ws["A9"] = CONCLUSION
    ws.merge_cells("A9:D9")
    ws["A9"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws["A9"].fill = PatternFill("solid", fgColor=GREY_BG)
    ws["A9"].font = Font(size=10, color="0F172A")
    ws.row_dimensions[9].height = 26

    ws.freeze_panes = "A3"
    for col, w in {"A": 16.5, "B": 13.2, "C": 12.6, "D": 14.0, "E": 2.0, "F": 12.0}.items():
        ws.column_dimensions[col].width = w
    wb.defined_names["经营数据"] = DefinedName("经营数据", attr_text="'月度经营分析'!$B$3:$D$6")

    det = wb.create_sheet("9月明细")
    det["A1"] = "9 月明细（分项，已并入分析表）"
    det["A1"].font = Font(bold=True, size=12)
    det["A4"], det["B4"] = "收入分项", "金额（万元）"
    for i, (name, amt) in enumerate([("华东区", 42), ("华南区", 31), ("华北区", 22), ("海外", 13)], start=5):
        det.cell(row=i, column=1, value=name)
        det.cell(row=i, column=2, value=amt).number_format = "#,##0"
    det["A9"], det["B9"] = "合计", "=SUM(B5:B8)"
    det["A9"].font = Font(bold=True)
    det["B9"].number_format = "#,##0"
    det["A11"], det["B11"] = "成本分项", "金额（万元）"
    for i, (name, amt) in enumerate([("采购", 46), ("人力", 21), ("渠道", 9), ("其他", 4)], start=12):
        det.cell(row=i, column=1, value=name)
        det.cell(row=i, column=2, value=amt).number_format = "#,##0"
    det["A16"], det["B16"] = "合计", "=SUM(B12:B15)"
    det["A16"].font = Font(bold=True)
    det["B16"].number_format = "#,##0"
    det.column_dimensions["A"].width = 18
    det.column_dimensions["B"].width = 14

    wb.save(path)


# ---------------------------------------------------------------------------
# 作品二 · PPT
# ---------------------------------------------------------------------------
def _slide_blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _textbox(slide, left, top, width, height, text, size, color, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = PptPt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "微软雅黑"
    return box


def build_ppt(path: Path) -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)   # 16:9
    W = prs.slide_width

    # 1 封面
    s = _slide_blank(prs)
    bar = s.shapes.add_shape(1, 0, 0, W, Inches(0.35))
    bar.fill.solid(); bar.fill.fore_color.rgb = PPT_BLUE; bar.line.fill.background()
    _textbox(s, Inches(1.0), Inches(2.4), W - Inches(2.0), Inches(1.4),
             "2026 年三季度经营汇报", 44, PPT_BLUE, bold=True)
    _textbox(s, Inches(1.0), Inches(3.9), W - Inches(2.0), Inches(0.6),
             "数据来源：月度经营分析 · 2026-09 · 单位：万元", 16, PPT_GREY)

    # 2 三张指标卡
    s = _slide_blank(prs)
    _textbox(s, Inches(0.9), Inches(0.6), Inches(8), Inches(0.8), "核心指标", 32, PPT_INK, bold=True)
    cards = [("REVENUE", "108 万元", "9 月营业收入，环比 +14.9%"),
             ("PROFIT", "28 万元", "9 月营业利润，环比 +27.3%"),
             ("MARGIN", "25.9%", "利润率较 7 月提升 2.6 个百分点")]
    cw, gap = Inches(3.6), Inches(0.45)
    for i, (tag, big, desc) in enumerate(cards):
        left = Inches(0.9) + i * (cw + gap)
        card = s.shapes.add_shape(1, left, Inches(2.1), cw, Inches(2.6))
        card.fill.solid(); card.fill.fore_color.rgb = PptRGB(0xF8, 0xFA, 0xFC)
        card.line.color.rgb = PptRGB(0xE2, 0xE8, 0xF0)
        _textbox(s, left + Inches(0.3), Inches(2.35), cw - Inches(0.6), Inches(0.4), tag, 12, PPT_BLUE, bold=True)
        _textbox(s, left + Inches(0.3), Inches(2.9), cw - Inches(0.6), Inches(0.9), big, 30, PPT_INK, bold=True)
        _textbox(s, left + Inches(0.3), Inches(3.9), cw - Inches(0.6), Inches(0.6), desc, 13, PPT_GREY)

    # 3 原生图表页
    s = _slide_blank(prs)
    _textbox(s, Inches(0.9), Inches(0.6), Inches(9), Inches(0.8), "营收与成本趋势", 32, PPT_INK, bold=True)
    cd = CategoryChartData()
    cd.categories = MONTHS
    cd.add_series("营业收入", tuple(REVENUE))
    cd.add_series("运营成本", tuple(COST))
    gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.9), Inches(1.8),
                                Inches(11.5), Inches(4.8), cd)
    chart = gframe.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "单位：万元"
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    # 4 结论与行动
    s = _slide_blank(prs)
    _textbox(s, Inches(0.9), Inches(0.6), Inches(9), Inches(0.8), "结论与行动", 32, PPT_INK, bold=True)
    bullets = [
        "营收连续三月增长，9 月环比 +14.9%，海外占比升至 12.0%",
        "利润率逐月改善至 25.9%，结构优化有效",
        "运营成本环比 +11.1%，采购与渠道费用需专项复核",
        "下月动作：完成成本结构复盘，评估海外渠道复制",
    ]
    box = s.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.2))
    tf = box.text_frame; tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "·  " + b
        p.font.size = PptPt(19); p.font.color.rgb = PPT_INK; p.font.name = "微软雅黑"
        p.space_after = PptPt(16)

    # 5 结束页
    s = _slide_blank(prs)
    _textbox(s, Inches(1.0), Inches(2.9), W - Inches(2.0), Inches(1.2), "谢谢", 48, PPT_BLUE, bold=True)
    _textbox(s, Inches(1.0), Inches(4.2), W - Inches(2.0), Inches(0.6),
             "字浮 CharFloat · 让 AI 直接在你的桌面上把活干完", 16, PPT_GREY)

    prs.save(path)


# ---------------------------------------------------------------------------
# 作品三 · 文档
# ---------------------------------------------------------------------------
def _set_cjk(doc, font_name="微软雅黑"):
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def _three_line_table(table) -> None:
    """麦肯锡三线表：顶底粗线 + 表头下细线，无内部竖线。"""
    def edges(cell, top=None, bottom=None):
        tcPr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement("w:tcBorders")
        for name, val in (("top", top), ("bottom", bottom)):
            el = OxmlElement(f"w:{name}")
            if val is None:
                el.set(qn("w:val"), "nil")
            else:
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(val))
                el.set(qn("w:color"), "1F2937")
            borders.append(el)
        for name in ("left", "right"):
            el = OxmlElement(f"w:{name}")
            el.set(qn("w:val"), "nil")
            borders.append(el)
        tcPr.append(borders)

    for r, row in enumerate(table.rows):
        for cell in row.cells:
            if r == 0:
                edges(cell, top=18, bottom=8)
            elif r == len(table.rows) - 1:
                edges(cell, bottom=18)
            else:
                edges(cell)


def build_doc(path: Path) -> None:
    doc = Document()
    _set_cjk(doc)
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Mm(25.4)
    sec.left_margin = sec.right_margin = Mm(28)

    h = doc.add_heading("2026 年三季度经营分析报告", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("一、总体表现", level=2)
    doc.add_paragraph(
        "2026 年 7 至 9 月，公司营业收入由 86 万元增长至 108 万元，累计增幅 25.6%；"
        "营业利润由 20 万元增长至 28 万元；利润率由 23.3% 提升至 25.9%。"
        "运营成本由 66 万元上升至 80 万元，环比增速高于营收增速，是本期需要重点关注的科目。"
    )

    table = doc.add_table(rows=5, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    rows = [
        ["经营指标 / 万元", *MONTHS],
        ["营业收入", *[str(v) for v in REVENUE]],
        ["运营成本", *[str(v) for v in COST]],
        ["营业利润", *[str(v) for v in PROFIT]],
        ["利润率", *MARGIN],
    ]
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.size = Pt(10)
            run.font.bold = (r == 0) or (c == 0 and r > 0)
    _three_line_table(table)

    doc.add_heading("二、结论与行动建议", level=2)
    for item in [
        "成本专项：9 月运营成本环比增长 11.1%，建议对采购价格与渠道费用逐项复核。",
        "结构优化：利润率连续三月改善，维持在 25% 以上，当前收入结构可继续执行。",
        "增量探索：海外收入占 9 月营收 12.0%，建议评估该渠道的可复制性。",
        "数据口径：本报告数据源为《月度经营.xlsx》，与演示文稿使用同一份口径。",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.save(path)


# ---------------------------------------------------------------------------
def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    made = []

    xlsx = OUT / "月度经营分析报表_成品.xlsx"
    build_sheet(xlsx); made.append(xlsx)

    pptx = OUT / "2026年三季度经营汇报_成品.pptx"
    build_ppt(pptx); made.append(pptx)

    docx = OUT / "2026年三季度经营分析报告_成品.docx"
    build_doc(docx); made.append(docx)

    for p in made:
        print(f"  {p.name}  ({p.stat().st_size / 1024:.1f} KB)")
    print("\n三个成品已生成。数据全部为虚构演示数据。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
