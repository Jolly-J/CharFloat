#!/usr/bin/env python3
"""e2e 步骤 5：用 MCP 专用工具把配套说明写进新建 Word 文档。"""
import json
import subprocess
import sys

ROOT = "/Users/Python/Office Agent Bridge"
DOC = "文字文稿3"


def call(tool, args):
    out = subprocess.run(
        ["python3", "skills/office-agent-bridge/scripts/bridge_client.py", tool, json.dumps(args, ensure_ascii=False)],
        cwd=ROOT, capture_output=True, text=True,
    )
    text = out.stdout.strip() or out.stderr.strip()
    try:
        parsed = json.loads(text)
        ok = parsed.get("success", True)
        brief = {k: v for k, v in (parsed.get("data") or {}).items() if k in ("message", "success", "tableIndex", "rows", "columns")}
        print(f"[{'OK ' if ok else 'ERR'}] {tool} -> {json.dumps(brief, ensure_ascii=False)[:160]}")
        return ok
    except Exception:
        print(f"[ERR] {tool} -> {text[:200]}")
        return False


STEPS = [
    ("wps_word_write_content", {"documentName": DOC, "content": "钙钛矿光伏产业链现状说明", "type": "heading1"}),
    ("wps_word_write_content", {"documentName": DOC, "type": "paragraph", "content":
        "本文档为 e2e 端到端交付场景的配套说明，全部数据取自工作簿《钙钛矿各家企业现状.xlsx》中已存在的"
        "「企业现状总览」「市场预测」「行业概览」三张工作表，未引入外部数据。汇总口径已写入新增工作表"
        "「e2e_钙钛矿数据集」，便于逐项回溯。"}),
    ("wps_word_write_content", {"documentName": DOC, "content": "一、数据来源与口径", "type": "heading2"}),
    ("wps_word_write_content", {"documentName": DOC, "type": "bullet_list", "content": [
        "企业样本：企业现状总览表第 3–39 行，共 37 家企业，字段含国家、省份、类型、技术路线、最新进展。",
        "市场预测：市场预测表 2025 / 2030 / 2035 三列的产能、产量、需求数据。",
        "技术与成本指标：行业概览表 18 行指标，含理论效率、实测效率、成本目标与瓶颈。",
        "类型归并：按「类型」字段分组统计，专业钙钛矿 8 + 传统光伏 8 + 跨界龙头 3 = 电池与组件 19 家。"]}),
    ("wps_word_write_content", {"documentName": DOC, "content": "二、核心数据集", "type": "heading2"}),
    ("wps_word_write_content", {"documentName": DOC, "type": "paragraph",
        "content": "下表与工作簿「e2e_钙钛矿数据集」工作表逐行对应，演示文稿中的图表数值同样取自该数据集。"}),
    ("wps_word_manage_table", {"documentName": DOC, "action": "insert", "rows": 21, "columns": 4,
        "stylePreset": "mckinsey_three_line", "data": [
            ["指标", "数值", "单位", "来源工作表"],
            ["企业总数", "37", "家", "企业现状总览"],
            ["中国", "36", "家", "企业现状总览"],
            ["海外", "1", "家", "企业现状总览"],
            ["专业钙钛矿", "8", "家", "企业现状总览"],
            ["传统光伏", "8", "家", "企业现状总览"],
            ["设备", "9", "家", "企业现状总览"],
            ["材料", "6", "家", "企业现状总览"],
            ["跨界龙头", "3", "家", "企业现状总览"],
            ["上游资源", "2", "家", "企业现状总览"],
            ["太空光伏", "1", "家", "企业现状总览"],
            ["江苏省", "12", "家", "企业现状总览"],
            ["广东省", "6", "家", "企业现状总览"],
            ["产能（2025 年）", "4", "GW", "市场预测"],
            ["产能（2030 年）", "108", "GW", "市场预测"],
            ["产能（2035 年）", "300", "GW", "市场预测"],
            ["市场需求（2030 年）", "45", "GW", "市场预测"],
            ["市场需求（2035 年）", "130", "GW", "市场预测"],
            ["叠层组件最高效率", "29.5", "%", "行业概览"],
            ["组件成本目标", "降至 0.5 元/W 以下", "—", "行业概览"]]}),
    ("wps_word_write_content", {"documentName": DOC, "content": "三、市场空间与产业化节奏", "type": "heading2"}),
    ("wps_word_write_content", {"documentName": DOC, "type": "bullet_list", "content": [
        "产能规划：2025 年 4GW → 2030 年 108GW → 2035 年 300GW，2030 年产能为 2025 年的 27 倍。",
        "需求端：整体市场需求 2030 年 45GW、2035 年 130GW。",
        "效率：单结理论效率 33%、叠层理论效率 45%，均高于晶硅理论极限 29.4%；叠层组件实测最高效率 29.5%（2048 cm²）。",
        "节奏：2025 年为量产元年，2026 年多条 GW 级产线投产，当年全球组件量产规模预测 15–20GW。"]}),
    ("wps_word_write_content", {"documentName": DOC, "content": "四、结论与建议", "type": "heading2"}),
    ("wps_word_write_content", {"documentName": DOC, "type": "bullet_list", "content": [
        "产业链已成形但尚未出清：37 家中 19 家做电池组件、9 家做设备，江苏省以 12 家居首。",
        "叠层是效率与量产的交汇点，单结路线更适合柔性、BIPV、太空光伏等差异化场景。",
        "两道硬门槛：稳定寿命 3–5 年（晶硅 20 年以上）与组件成本降至 0.5 元/W 以下。",
        "建议优先跟踪 GW 级产线良率与实证电站数据，未跨过门槛前不宜按 2035 年 300GW 线性外推。"]}),
]

if __name__ == "__main__":
    fails = 0
    for tool, args in STEPS:
        if not call(tool, args):
            fails += 1
    print(f"steps={len(STEPS)} fails={fails}")
    sys.exit(1 if fails else 0)
