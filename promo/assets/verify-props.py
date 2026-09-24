#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
道具验收：把"这些道具到底是不是真的"写成一份可核对的报告。

检查项：
  1. 加密 xlsx 是真加密（file 判定 CDFV2 Encrypted、openpyxl 拒读；装有 msoffcrypto 时再验口令能解开）
  2. 加密 zip 需要口令（unzip 无口令失败、带口令成功）
  3. 演示工作簿的关键特征（合并标题、批注、冻结、命名区域、动态公式、9 月留空）
  4. 5 个"AI 返回"版本道具确实丢了格式/公式（这是片 A 对比拍的依据）
  5. 三个成品基线文件存在且结构正确

用法：
  python3 promo/assets/verify-props.py                    # 打印 + 写报告
  python3 promo/assets/verify-props.py --no-report        # 只打印
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
from docx import Document
from pptx import Presentation

PROMO = Path(__file__).resolve().parent.parent
ENV_A = PROMO / "env-A-无插件"
ENV_B = PROMO / "env-B-有插件"
FINALS = PROMO / "成品基线"
ENC_DIR = ENV_A / "03-加密文件"
REPORT = ENC_DIR / "加密校验报告.txt"

DEMO_PASSWORD = "Demo#2026"   # 演示口令，非真实凭据

lines: list[str] = []
failures: list[str] = []


def say(text: str = "") -> None:
    print(text)
    lines.append(text)


def check(label: str, ok: bool, detail: str = "") -> None:
    mark = "通过" if ok else "不通过"
    say(f"  [{mark}] {label}" + (f" —— {detail}" if detail else ""))
    if not ok:
        failures.append(label)


# ---------------------------------------------------------------------------
def check_encrypted() -> None:
    say("一、加密道具是不是真加密")
    enc_xlsx = ENC_DIR / "2026年度集团经营数据_机密.xlsx"
    enc_zip = ENC_DIR / "2026年度集团经营数据_机密.zip"

    if not enc_xlsx.exists():
        check("加密 xlsx 存在", False, str(enc_xlsx))
        return
    check("加密 xlsx 存在", True, f"{enc_xlsx.stat().st_size} 字节")

    file_out = subprocess.run(["file", str(enc_xlsx)], capture_output=True, text=True).stdout.strip()
    is_cfb = "CDFV2 Encrypted" in file_out or "Composite Document File" in file_out
    say(f"        file 判定：{file_out.split(': ', 1)[-1]}")
    check("宿主层面是真加密容器（不是普通 xlsx）", is_cfb)

    try:
        openpyxl.load_workbook(enc_xlsx)
        opened = True
    except Exception as e:                                  # noqa: BLE001
        opened = False
        say(f"        openpyxl 读取结果：{type(e).__name__} —— 云端解析器撞的就是这堵墙")
    check("常规解析器打不开（云端 AI 拿不到内容）", not opened)

    try:
        import msoffcrypto                                  # noqa: PLC0415
        import io
        with open(enc_xlsx, "rb") as f:
            office = msoffcrypto.OfficeFile(f)
            keyed = office.is_encrypted()
            office.load_key(password=DEMO_PASSWORD)
            buf = io.BytesIO()
            office.decrypt(buf)
        wb = openpyxl.load_workbook(buf)
        check("带口令可以正常解开（说明只是加密，不是坏文件）", True,
              f"解密后工作表：{wb.sheetnames}")
    except ImportError:
        say("        [跳过] 未安装 msoffcrypto-tool，无法做解密校验（不影响'打不开'这一结论）")
    except Exception as e:                                  # noqa: BLE001
        check("带口令可以正常解开", False, f"{type(e).__name__}: {e}")

    if enc_zip.exists():
        check("加密 zip 存在", True, f"{enc_zip.stat().st_size} 字节")
        # 用空口令模拟"没有密码就想解压"；解不出内容会往 stdout 吐原始字节，必须容错解码
        no_pw = subprocess.run(["unzip", "-P", "", "-p", str(enc_zip)],
                               capture_output=True, text=True, errors="replace")
        check("无口令解不开 zip",
              no_pw.returncode != 0 or "incorrect password" in (no_pw.stderr or "").lower(),
              (no_pw.stderr or "").strip().splitlines()[-1] if no_pw.stderr else "")
        with_pw = subprocess.run(["unzip", "-P", DEMO_PASSWORD, "-p", str(enc_zip)],
                                 capture_output=True, text=True, errors="replace")
        check("带口令能解开 zip", with_pw.returncode == 0 and len(with_pw.stdout) > 100,
              f"解出 {len(with_pw.stdout)} 字节")
        try:
            with zipfile.ZipFile(enc_zip) as z:
                say(f"        zip 内文件：{z.namelist()}")
        except Exception:                                   # noqa: BLE001
            say("        (zip 用传统加密，zipfile 需口令才能列目录，属正常)")
    else:
        check("加密 zip 存在", False, str(enc_zip))


# ---------------------------------------------------------------------------
def check_demo_workbook() -> None:
    say("")
    say("二、片 B 的演示工作簿（要留给 AI 去改的那份）")
    p = ENV_B / "月度经营.xlsx"
    if not p.exists():
        check("演示工作簿存在", False, str(p))
        return
    wb = openpyxl.load_workbook(p)
    ws = wb["月度经营分析"]
    check("工作表齐备", wb.sheetnames == ["月度经营分析", "9月明细"], str(wb.sheetnames))
    check("标题已合并（AI 不许破坏它）", "A1:D1" in [str(r) for r in ws.merged_cells.ranges])
    check("表头冻结在 A3", ws.freeze_panes == "A3", str(ws.freeze_panes))
    check("存在命名区域「经营数据」", "经营数据" in list(wb.defined_names))
    check("C3 有对账批注（AI 不许弄丢）", ws["C3"].comment is not None)
    check("7/8 月利润与利润率是动态公式",
          ws["B5"].value == "=B3-B4" and ws["B6"].value == "=B5/B3",
          f"B5={ws['B5'].value}  B6={ws['B6'].value}")
    check("9 月列留空（等 AI 来补）",
          all(ws.cell(r, 4).value is None for r in range(3, 7)),
          str([ws.cell(r, 4).value for r in range(3, 7)]))
    det = wb["9月明细"]
    check("9 月明细的合计格是空的（等 AI 来算）",
          det["B9"].value is None and det["B16"].value is None)
    income = sum(det.cell(r, 2).value for r in range(5, 9))
    cost = sum(det.cell(r, 2).value for r in range(12, 16))
    check("分项加总等于演示口径 108 / 80", income == 108 and cost == 80, f"{income} / {cost}")


# ---------------------------------------------------------------------------
def check_ai_versions() -> None:
    say("")
    say("三、片 A 的 5 个「AI 返回」版本道具（摆拍：用于稳定复现格式流失）")
    d = ENV_A / "02-AI返回的文件"
    v1 = d / "月度经营_分析报表.xlsx"
    if not v1.exists():
        check("版本道具存在", False, str(d))
        return
    files = sorted(d.glob("*.xlsx"))
    check("5 个版本都在", len(files) == 5, "、".join(f.name for f in files))

    wb = openpyxl.load_workbook(v1)
    ws = wb.active
    check("v1 丢了标题合并（对比源文件的 A1:D1）",
          len(list(ws.merged_cells.ranges)) == 0)
    check("v1 丢了冻结窗格（对比源文件的 A3）", ws.freeze_panes is None)
    check("v1 丢了批注", not any(c.comment for row in ws.iter_rows() for c in row))
    check("v1 的利润变成了死数字（对比源文件的 =B3-B4）",
          isinstance(ws["B5"].value, (int, float)), f"B5={ws['B5'].value!r}")
    check("v1 的利润率丢了百分比格式", ws["B6"].number_format == "General",
          f"B6 格式={ws['B6'].number_format}")

    v2 = d / "月度经营_分析报表_v2.xlsx"
    if v2.exists():
        ws2 = openpyxl.load_workbook(v2).active
        check("v2 的图表是「插进来的图片」而不是原生图表",
              len(ws2._images) >= 1 and len(ws2._charts) == 0,
              f"图片 {len(ws2._images)} 个 / 原生图表 {len(ws2._charts)} 个")


# ---------------------------------------------------------------------------
def check_finals() -> None:
    say("")
    say("四、三个成品基线")
    xs = FINALS / "月度经营分析报表_成品.xlsx"
    pp = FINALS / "2026年三季度经营汇报_成品.pptx"
    dc = FINALS / "2026年三季度经营分析报告_成品.docx"

    for f in (xs, pp, dc):
        check(f"{f.name} 存在", f.exists(), f"{f.stat().st_size / 1024:.1f} KB" if f.exists() else "")

    if xs.exists():
        wb = openpyxl.load_workbook(xs)
        ws = wb["月度经营分析"]
        check("成品表格：9 月用跨表 SUM 公式（与施工 sheet-03 一致）",
              str(ws["D3"].value).startswith("=SUM("), str(ws["D3"].value))
        check("成品表格：条件格式两处（营收标蓝 / 成本标黄）",
              len(list(ws.conditional_formatting)) == 2)
        check("成品表格：原生图表 1 张", len(ws._charts) == 1)
        check("成品表格：结论整行合并", "A9:D9" in [str(r) for r in ws.merged_cells.ranges])
    if pp.exists():
        prs = Presentation(pp)
        has_chart = any(sh.has_chart for s in prs.slides for sh in s.shapes)
        check("成品 PPT：5 页", len(prs.slides) == 5, f"{len(prs.slides)} 页")
        check("成品 PPT：含原生图表", has_chart)
    if dc.exists():
        doc = Document(dc)
        check("成品文档：含 5×4 三线表", len(doc.tables) == 1 and len(doc.tables[0].rows) == 5)
        check("成品文档：含两级标题",
              sum(1 for p in doc.paragraphs if p.style.name.startswith("Heading")) >= 3)


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-report", action="store_true", help="不写报告文件")
    args = ap.parse_args()

    say("字浮 CharFloat · 宣传片道具验收报告")
    say("=" * 62)
    say("说明：全部数据为 CLI 合成的虚构演示数据，不含任何真实业务数据。")
    say("")

    check_encrypted()
    check_demo_workbook()
    check_ai_versions()
    check_finals()

    say("")
    say("=" * 62)
    if failures:
        say(f"结论：{len(failures)} 项不通过 ——")
        for f in failures:
            say(f"  - {f}")
    else:
        say("结论：全部通过。道具可用于拍摄。")

    if not args.no_report:
        ENC_DIR.mkdir(parents=True, exist_ok=True)
        REPORT.write_text("\n".join(lines) + "\n", "utf-8")
        print(f"\n报告已写入：{REPORT.relative_to(PROMO.parent)}")

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
