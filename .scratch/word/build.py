#!/usr/bin/env python3
"""构建 `测试文字文稿.docx` 成品：分阶段调用 MCP 工具，全程记录请求/响应。

用法: python3 .scratch/word/build.py <phase>
阶段: probe | content | tables | toc | layout | review | verify | save
"""
import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SCRATCH = ROOT / '.scratch' / 'word'
LOG = SCRATCH / 'run.log'
DOC = '测试文字文稿.docx'
TOKEN = (Path.home() / '.wps-bridge' / 'token').read_text().strip()
URL = 'http://127.0.0.1:19890/api/v1/tool/call'
SESSION = 'scratch-word-author'


def log(kind, payload):
    line = '[%s] %s %s\n' % (time.strftime('%H:%M:%S'), kind, json.dumps(payload, ensure_ascii=False))
    with LOG.open('a', encoding='utf-8') as fh:
        fh.write(line)


def call(tool, args, expect=True):
    body = {'name': tool, 'arguments': args, 'clientName': 'WPS Testdoc Author', 'sessionId': SESSION}
    log('CALL', {'tool': tool, 'args': args})
    try:
        req = urllib.request.Request(URL, headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + TOKEN}, data=json.dumps(body).encode())
        with urllib.request.urlopen(req, timeout=180) as r:
            result = json.load(r)
    except urllib.error.HTTPError as e:
        result = {'httpError': e.code, 'body': e.read().decode()[:4000]}
    except Exception as e:  # noqa: BLE001
        result = {'transportError': repr(e)}
    log('RESULT', result)
    if expect and not result.get('success'):
        print('!! 失败: %s -> %s' % (tool, json.dumps(result, ensure_ascii=False)[:1500]))
        raise SystemExit(1)
    print('-- %s: %s' % (tool, json.dumps(result.get('data', result), ensure_ascii=False)[:400]))
    return result.get('data', result)


def script(code, params=None):
    return call('wps_execute_script', {'documentName': DOC, 'component': 'word', 'code': code, 'params': params or {}})


def wc(lines, ptype=None, fmt=None, location='end'):
    args = {'documentName': DOC, 'content': lines, 'location': location}
    if ptype:
        args['type'] = ptype
    if fmt:
        args['formatting'] = fmt
    return call('wps_word_write_content', args)


BODY_FMT = {'fontSizePt': 12, 'fontName': '宋体', 'firstLineIndentChars': 2, 'lineSpacingPt': 24, 'alignment': 3, 'spaceBeforePt': 0, 'spaceAfterPt': 3}
H1_FMT = {'fontSizePt': 16, 'fontName': '黑体', 'bold': True, 'lineSpacingPt': 28, 'spaceBeforePt': 12, 'spaceAfterPt': 8, 'alignment': 0}
H2_FMT = {'fontSizePt': 14, 'fontName': '黑体', 'bold': True, 'lineSpacingPt': 26, 'spaceBeforePt': 10, 'spaceAfterPt': 6, 'alignment': 0}
H3_FMT = {'fontSizePt': 12, 'fontName': '黑体', 'bold': True, 'lineSpacingPt': 24, 'spaceBeforePt': 8, 'spaceAfterPt': 4, 'alignment': 0}
BULLET_FMT = {'fontSizePt': 12, 'fontName': '宋体', 'lineSpacingPt': 24, 'alignment': 0, 'spaceAfterPt': 2}
NOTE_FMT = {'fontSizePt': 9, 'fontName': '楷体', 'lineSpacingPt': 18, 'alignment': 0, 'spaceBeforePt': 4, 'spaceAfterPt': 10}


def phase_reset():
    call('wps_word_manage_content', {'documentName': DOC, 'action': 'clear_all'})
    data = call('wps_word_read_document', {'documentName': DOC, 'scope': 'full', 'maxParagraphs': 5}, expect=False)
    print('== 重置后 paragraphs=%s tables=%s' % (data.get('paragraphCount'), data.get('tableCount')))


def phase_cover():
    # ---- 封面 ----
    wc(['Office Agent Bridge'], 'paragraph', {'fontSizePt': 30, 'fontName': '微软雅黑', 'bold': True, 'alignment': 1, 'spaceBeforePt': 90, 'spaceAfterPt': 10, 'lineSpacingPt': 40})
    wc(['2.1 版本发布与验收说明'], 'paragraph', {'fontSizePt': 18, 'fontName': '微软雅黑', 'alignment': 1, 'spaceAfterPt': 60, 'lineSpacingPt': 30})
    wc(['文档编号：OAB-REL-2025-021　　密级：内部'], 'paragraph', {'fontSizePt': 12, 'fontName': '宋体', 'alignment': 1, 'lineSpacingPt': 24, 'spaceAfterPt': 4})
    wc(['版本号：{{VERSION}}　　编制日期：{{DATE}}'], 'paragraph', {'fontSizePt': 12, 'fontName': '宋体', 'alignment': 1, 'lineSpacingPt': 24, 'spaceAfterPt': 4})
    wc(['编制单位：信息化建设部'], 'paragraph', {'fontSizePt': 12, 'fontName': '宋体', 'alignment': 1, 'lineSpacingPt': 24, 'spaceAfterPt': 4})
    # ---- 目录标题（正文首段另起一页，靠脚本设 PageBreakBefore）----
    wc(['目　录'], 'paragraph', {'fontSizePt': 20, 'fontName': '黑体', 'bold': True, 'alignment': 1, 'spaceAfterPt': 14, 'lineSpacingPt': 30})


def phase_body():
    # ---- 一、概述 ----
    wc(['一、概述'], 'heading1', H1_FMT)
    wc(['1.1 编写目的'], 'heading2', H2_FMT)
    wc(['本文档用于说明 Office Agent Bridge 2.1 版本的交付范围、验收结论与后续计划，供项目干系人评审与归档使用。文档同时记录本次验收所依据的环境、方法与证据边界，便于下一个版本直接复用。',
        '本版本以“MCP 契约稳定、双宿主可验证”为目标，重点收敛工具定义与宿主实现之间的偏差，并在真实 WPS 宿主上完成端到端回归。'], 'paragraph', BODY_FMT)
    wc(['1.2 适用范围'], 'heading2', H2_FMT)
    wc(['本文档适用于以下角色与场景：'], 'paragraph', BODY_FMT)
    wc(['项目组开发与测试人员：了解本版本交付内容、验证范围与遗留风险；',
        '业务部门最终用户：确认办公文档自动化能力是否满足日常使用需要；',
        '运维与支持人员：掌握部署要求、排障入口与回滚边界。'], 'bullet_list', BULLET_FMT)

    # ---- 二、版本交付内容 ----
    wc(['二、版本交付内容'], 'heading1', H1_FMT)
    wc(['2.1 主要能力清单'], 'heading2', H2_FMT)
    wc(['2.1 版本按组件交付结构化操作入口，覆盖表格、文字与演示三类文档，并按宿主能力差异分别标注实现状态。',
        '所有工具的入参校验、审计记录与失败分类集中在网关层统一处理，宿主加载项只负责对象模型操作，避免同一能力在不同组件中重复实现。'], 'paragraph', BODY_FMT)
    wc(['2.2 模块交付状态'], 'heading2', H2_FMT)
    wc(['各模块的交付范围与验收结论如表 1 所示。'], 'paragraph', BODY_FMT)


def phase_tables():
    call('wps_word_manage_table', {
        'documentName': DOC, 'action': 'insert', 'stylePreset': 'mckinsey_three_line', 'repeatHeader': True,
        'data': [
            ['模块', '交付状态', '主要变更', '责任方', '验收结论'],
            ['MCP 网关契约', '已交付', '工具定义按类拆分并固化发布顺序', '平台组', '通过'],
            ['WPS 加载项', '已交付', '按表格/文字/演示拆分宿主执行层', '宿主组', '通过'],
            ['Microsoft 通道', '部分交付', '保留 Office.js 通道，原生驱动待实机验收', '互操作组', '有条件通过'],
            ['审计与回滚', '部分交付', '覆盖单元格值与公式，样式与结构不回滚', '平台组', '有条件通过'],
        ]})
    wc(['注：表中“有条件通过”表示主体能力可用，但仍有待实机补充验证的分支。'], 'paragraph', NOTE_FMT)
    wc(['三、验收测试结果'], 'heading1', H1_FMT)
    wc(['3.1 测试范围与方法'], 'heading2', H2_FMT)
    wc(['本次验收以真实 WPS 宿主上的实机操作为主，静态检查、构建通过与模拟通道仅作为前置门槛，不替代实机结论。检查项包括：'], 'paragraph', BODY_FMT)
    wc(['契约一致性：核对工具定义、网关分支与宿主实现三者的对应关系；',
        '文件级回归：在专用测试文档上完成写入、排版、保存与导出；',
        '边界场景：断线、超时与宿主未实现方法的失败分类与提示。'], 'bullet_list', BULLET_FMT)
    wc(['3.2 关键指标'], 'heading2', H2_FMT)
    wc(['本次验收的关键指标实测结果如表 2 所示。'], 'paragraph', BODY_FMT)
    call('wps_word_manage_table', {
        'documentName': DOC, 'action': 'insert', 'stylePreset': 'mckinsey_three_line', 'repeatHeader': True,
        'data': [
            ['指标项', '目标值', '实测值', '结论'],
            ['工具定义与网关分支一致率', '100%', '100%', '达标'],
            ['加载项入口与源码一致性', '无漂移', '无漂移', '达标'],
            ['实机写入成功率', '≥ 99%', '99.4%', '达标'],
            ['宿主预览导出成功率', '≥ 95%', '96.8%', '达标'],
        ]})
    wc(['注：指标取值为本次验收批次内的统计结果，不作为长期服务承诺。'], 'paragraph', NOTE_FMT)
    wc(['3.3 遗留问题'], 'heading3', H3_FMT)
    wc(['原生驱动分支尚未完成实机验收，相关结论需在 Windows 环境补充验证后才能确认。',
        '部分只读能力在个别宿主上仍缺少等价实现，跨宿主行为差异需在下一个版本中继续收敛。'], 'paragraph', BODY_FMT)
    wc(['四、结论与后续计划'], 'heading1', H1_FMT)
    wc(['4.1 验收结论'], 'heading2', H2_FMT)
    wc(['经上述检查，{{VERSION}} 版本的核心能力达到发布标准，同意按“内部试用”范围发布，并在下一个版本中继续收敛遗留问题。'], 'paragraph', BODY_FMT)
    wc(['4.2 后续计划'], 'heading2', H2_FMT)
    wc(['补齐 Windows 原生驱动分支的实机验收用例；',
        '为样式、结构与文字/演示操作补充可回滚的快照机制；',
        '把验收证据按候选标识归档，形成下一个会话可直接使用的恢复材料。'], 'bullet_list', BULLET_FMT)
    wc(['编制单位：信息化建设部'], 'paragraph', {'fontSizePt': 12, 'fontName': '宋体', 'alignment': 2, 'lineSpacingPt': 24, 'spaceBeforePt': 16})
    wc(['审核：质量管理组'], 'paragraph', {'fontSizePt': 12, 'fontName': '宋体', 'alignment': 2, 'lineSpacingPt': 24})
    wc(['日期：{{DATE}}'], 'paragraph', {'fontSizePt': 12, 'fontName': '宋体', 'alignment': 2, 'lineSpacingPt': 24})


def phase_toc():
    # 正文写完后插目录：把选区放到「一、概述」段首，走 selection 插入路径
    probe = script(r"""
const ps = doc.Paragraphs;
let tocTitleIdx = 0, firstBodyIdx = 0;
for (let i = 1; i <= ps.Count; i++) {
  const t = (ps.Item(i).Range.Text || "").replace(/[\r\n\x07]/g, "");
  if (!tocTitleIdx && t === "目　录") tocTitleIdx = i;
  if (!firstBodyIdx && t === "一、概述") firstBodyIdx = i;
}
const anchor = ps.Item(firstBodyIdx).Range.Start;
app.Selection.SetRange(anchor, anchor);
return { tocTitleIdx: tocTitleIdx, firstBodyIdx: firstBodyIdx, anchor: anchor, applied: app.Selection.Range.Start, paragraphs: ps.Count };
""")
    print('-- 锚点: %s' % json.dumps(probe.get('returnValue'), ensure_ascii=False))
    call('wps_word_insert_table_of_contents', {'documentName': DOC, 'upperHeadingLevel': 1, 'lowerHeadingLevel': 3, 'insertLocation': 'selection', 'includePageNumbers': True})


def phase_polish():
    # 封面/目录单独成页；更新目录域与页码域；清掉目录段误继承的分页
    out = script(r"""
const ps = doc.Paragraphs;
let tocTitleIdx = 0, firstBodyIdx = 0;
for (let i = 1; i <= ps.Count; i++) {
  const t = (ps.Item(i).Range.Text || "").replace(/[\r\n\x07]/g, "");
  if (!tocTitleIdx && t === "目　录") tocTitleIdx = i;
  if (!firstBodyIdx && t === "一、概述") firstBodyIdx = i;
}
if (tocTitleIdx) ps.Item(tocTitleIdx).Format.PageBreakBefore = true;
for (let i = tocTitleIdx + 1; i < firstBodyIdx; i++) ps.Item(i).Format.PageBreakBefore = false;
if (firstBodyIdx) ps.Item(firstBodyIdx).Format.PageBreakBefore = true;
let tocUpdated = null, fieldCount = 0;
try { doc.TablesOfContents.Item(1).Update(); tocUpdated = "ok"; } catch (e) { tocUpdated = e.message; }
try { fieldCount = doc.Fields.Count; doc.Fields.Update(); } catch (e) {}
const tocParas = [];
for (let i = tocTitleIdx + 1; i < firstBodyIdx; i++) tocParas.push((ps.Item(i).Range.Text || "").replace(/[\r\n\x07\t]/g, " ").trim().slice(0, 60));
return { tocTitleIdx: tocTitleIdx, firstBodyIdx: firstBodyIdx, tocCount: doc.TablesOfContents.Count, tocUpdated: tocUpdated, fieldCount: fieldCount, tocParas: tocParas };
""")
    print('-- 目录域: %s' % json.dumps(out.get('returnValue'), ensure_ascii=False))


def phase_layout():
    call('wps_word_format_document', {'documentName': DOC, 'preset': 'custom', 'margins': {'topMm': 30, 'bottomMm': 28, 'leftMm': 30, 'rightMm': 28}})
    # 专用工具的页眉页脚与水印（观察其真实落点）
    call('wps_word_page_layout_and_watermark', {'documentName': DOC, 'headerText': 'Office Agent Bridge 2.1 版本发布与验收说明', 'watermarkText': '内部资料 请勿外传'})
    call('wps_word_page_layout_and_watermark', {'documentName': DOC, 'differentFirstPage': True, 'pageNumberFormat': '第 X 页 / 共 Y 页'})


def phase_layout_fix():
    """用脚本补齐专用工具做不到的部分：整册水印（页眉层）+ 页脚页码域。"""
    out = script(r"""
const sec = doc.Sections.Item(1);
const report = { steps: [], watermarkAnchorPage: null, bodyShapesBefore: doc.Shapes.Count };

// 1) 记录专用工具水印的真实锚点页，然后清掉正文层水印
try {
  if (doc.Shapes.Count > 0) {
    const s = doc.Shapes.Item(1);
    report.watermarkAnchorPage = s.Anchor.Information(3);
    report.watermarkShape = { name: s.Name, left: Math.round(s.Left), top: Math.round(s.Top), width: Math.round(s.Width), rotation: s.Rotation };
  }
  for (let i = doc.Shapes.Count; i >= 1; i--) doc.Shapes.Item(i).Delete();
  report.steps.push("body-watermark-removed:" + doc.Shapes.Count);
} catch (e) { report.steps.push("cleanup-failed:" + e.message); }

// 2) 页眉层水印：页眉内容每页重复，因此水印覆盖全册
const headers = [sec.Headers.Item(1)];
try { if (doc.PageSetup.DifferentFirstPageHeaderFooter) headers.push(sec.Headers.Item(2)); } catch (e) {}
headers.forEach(function (h) {
  try { for (let i = h.Shapes.Count; i >= 1; i--) h.Shapes.Item(i).Delete(); } catch (de) {}
});
const pw = doc.PageSetup.PageWidth, ph = doc.PageSetup.PageHeight;
headers.forEach(function (h) {
  try {
    const shp = h.Shapes.AddTextEffect(0, "内部资料 请勿外传", "Microsoft YaHei", 44, false, false, 0, 0);
    shp.Rotation = 315;
    shp.Fill.Transparency = 0.82;
    shp.Line.Visible = false;
    shp.WrapFormat.Type = 3;
    shp.Left = (pw - shp.Width) / 2;
    shp.Top = (ph - shp.Height) / 2;
    try { shp.ZOrder(1); } catch (ze) {}
    report.steps.push("header-watermark:" + Math.round(shp.Width) + "x" + Math.round(shp.Height) + "@" + Math.round(shp.Left) + "," + Math.round(shp.Top));
  } catch (e) { report.steps.push("header-watermark-failed:" + e.message); }
});

// 3) 页脚页码：用 PAGE / NUMPAGES 域，专用工具没有该能力
try {
  const f = sec.Footers.Item(1);
  f.Range.Delete();
  const r = f.Range;
  r.InsertAfter("第 @@P@@ 页 / 共 @@N@@ 页");
  r.ParagraphFormat.Alignment = 1;
  f.Range.Font.Size = 10.5;
  f.Range.Font.NameFarEast = "宋体";
  function putField(token, code) {
    const rg = f.Range;
    const fnd = rg.Find;
    fnd.ClearFormatting();
    fnd.Text = token;
    fnd.Forward = true;
    fnd.Wrap = 0;
    if (fnd.Execute()) { doc.Fields.Add(fnd.Parent, -1, code, false); return true; }
    return false;
  }
  report.pField = putField("@@P@@", "PAGE");
  report.nField = putField("@@N@@", "NUMPAGES");
  f.Range.Fields.Update();
  report.footerText = (f.Range.Text || "").replace(/[\r\n\x13-\x15]/g, "").trim();
  report.footerFieldCount = f.Range.Fields.Count;
} catch (e) { report.steps.push("footer-failed:" + e.message); }

try { doc.TablesOfContents.Item(1).Update(); report.steps.push("toc-updated"); } catch (e) { report.steps.push("toc-update-failed:" + e.message); }
report.totalPages = doc.ComputeStatistics(2);
return report;
""")
    print('-- 版式补齐: %s' % json.dumps(out.get('returnValue'), ensure_ascii=False))


def phase_review():
    # 1) 占位符查找替换
    for q, r in (('{{VERSION}}', '2.1.0'), ('{{DATE}}', '2026年9月22日')):
        call('wps_word_find_and_replace', {'documentName': DOC, 'searchQuery': q, 'replaceText': r, 'matchCase': False, 'matchWholeWord': False, 'useWildcards': False, 'scope': 'full', 'replaceFormatting': {'bold': False, 'fontName': '宋体', 'fontSizePt': 12}})
    left = script(r"""
const hit = [];
const ps = doc.Paragraphs;
for (let i = 1; i <= ps.Count; i++) {
  const t = (ps.Item(i).Range.Text || "");
  if (t.indexOf("{{") >= 0) hit.push(i);
}
return { placeholdersLeft: hit.length, hitParagraphs: hit, versionHits: (doc.Content.Text.match(/2\.1\.0/g) || []).length, dateHits: (doc.Content.Text.match(/2026年9月22日/g) || []).length };
""")
    print('-- 替换核对: %s' % json.dumps(left.get('returnValue'), ensure_ascii=False))

    # 2) 修订痕迹：开启修订后改一处措辞，再关闭修订模式保留痕迹
    call('wps_word_review_and_comments', {'documentName': DOC, 'action': 'enable_track_changes'})
    call('wps_word_find_and_replace', {'documentName': DOC, 'searchQuery': '内部试用', 'replaceText': '内测与灰度试用', 'scope': 'full'})
    call('wps_word_review_and_comments', {'documentName': DOC, 'action': 'disable_track_changes'})

    # 3) 批注：先把选区定位到指代文本，再用专用工具加批注
    sel = script(r"""
const target = "原生驱动分支尚未完成实机验收";
const rng = doc.Content;
const f = rng.Find;
f.ClearFormatting(); f.Text = target; f.Forward = true; f.Wrap = 0;
if (f.Execute()) { app.Selection.SetRange(f.Parent.Start, f.Parent.End); }
return { found: !!f.Parent, selText: (app.Selection.Range.Text || "").slice(0, 40), selStart: app.Selection.Range.Start };
""")
    print('-- 批注锚点: %s' % json.dumps(sel.get('returnValue'), ensure_ascii=False))
    call('wps_word_review_and_comments', {'documentName': DOC, 'action': 'add_comment', 'commentText': '该结论依赖 Windows 实机验证，发布前请补齐证据后再定稿。', 'author': '质量管理组'})
    data = call('wps_word_review_and_comments', {'documentName': DOC, 'action': 'list_comments'})
    print('-- 批注列表: %s' % json.dumps(data, ensure_ascii=False)[:400])
    rev = script(r"""
const rv = [];
const r = doc.Revisions;
for (let i = 1; i <= r.Count; i++) {
  const x = r.Item(i);
  rv.push({ type: x.Type, author: x.Author, text: (x.Range.Text || "").slice(0, 30) });
}
return { revisionCount: r.Count, revisions: rv, trackRevisions: doc.TrackRevisions, comments: doc.Comments.Count };
""")
    print('-- 修订核对: %s' % json.dumps(rev.get('returnValue'), ensure_ascii=False))


def phase_save():
    out_dir = SCRATCH / 'export'
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = out_dir / '测试文字文稿.pdf'
    if pdf.exists():
        pdf.unlink()
    call('wps_word_save_document', {'documentName': DOC})
    call('wps_word_save_document', {'documentName': DOC, 'filePath': str(pdf), 'format': 'pdf'})
    print('-- PDF 存在: %s (%s bytes)' % (pdf.exists(), pdf.stat().st_size if pdf.exists() else -1))


def phase_structure():
    out = script(r"""
const ps = doc.Paragraphs;
const rows = [];
for (let i = 1; i <= ps.Count; i++) {
  const p = ps.Item(i);
  const t = (p.Range.Text || "").replace(/[\r\n\x07]/g, "");
  rows.push({ i: i, style: p.Style.NameLocal, text: t.slice(0, 30), page: p.Range.Information(3) });
}
const toc = [];
try { toc.push(doc.TablesOfContents.Item(1).Range.Text.replace(/[\r\n\t]/g, " | ")); } catch (e) { toc.push("err:" + e.message); }
return { totalPages: doc.ComputeStatistics(2), pageNumbersInFooter: (doc.Sections.Item(1).Footers.Item(1).Range.Text || "").replace(/[\r\n]/g, " ").trim(), tocText: toc[0], sample: rows.filter(function (r) { return r.style !== "正文" || r.i < 25; }) };
""")
    val = out.get('returnValue') or {}
    print(json.dumps(val, ensure_ascii=False, indent=1)[:3000])


if __name__ == '__main__':
    phase = sys.argv[1] if len(sys.argv) > 1 else 'probe'
    globals()['phase_' + phase]()
