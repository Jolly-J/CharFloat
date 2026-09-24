#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
施工驱动器：按 plan.json 的分步顺序，逐步调用本机 WPS 桥，让文件在屏幕上一点点长出来。

设计要点（都是为了录屏可控）：
  1. **单步可控**：--gate 每步等回车，导演可以按镜头节奏走；--only 只重录某一步，不用重头来。
  2. **跑之前先自检**：--dry-run 只用 plan.json 对 tool-schema.json 做校验（工具是否存在、
     参数名是否在 schema 里、必填是否齐），不连 WPS、不改文件。
  3. **不打印凭据**：token 从 ~/.wps-bridge/token 读，只用于 Authorization 头，绝不回显。
  4. **留证据**：每一步的原始响应落到 录制产出/responses/<step-id>.json。

用法：
  python3 promo/施工/build.py --list                       # 列出全部步骤
  python3 promo/施工/build.py --dry-run                    # 自检参数契约（不连 WPS）
  python3 promo/施工/build.py --group sheet --gate         # 录表格：每步等回车
  python3 promo/施工/build.py --group sheet --pace 2       # 录表格：每步停 2 秒
  python3 promo/施工/build.py --only sheet-09              # 只重录第 9 步（插图表）
  python3 promo/施工/build.py --all --pace 2 --out 录制产出
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROMO = HERE.parent
PLAN_PATH = HERE / "plan.json"
SCHEMA_PATH = PROMO / "assets" / "tool-schema.json"

DEFAULT_BASE_URL = os.environ.get("WPS_BRIDGE_URL", "http://127.0.0.1:19890")
BRIDGE_HOME = Path(os.environ.get("WPS_BRIDGE_HOME", Path.home() / ".wps-bridge"))
CLIENT_NAME = "CharFloat-Promo-Recorder"

# 参数自检时用来占位的假值（真跑时由 --workbook / 上一步 capture / --out 提供）
PLACEHOLDER_SAMPLES = {"presentationName": "演示文稿.pptx", "documentName": "报告.docx"}


# ---------------------------------------------------------------------------
# 基础
# ---------------------------------------------------------------------------
def die(msg: str, code: int = 1) -> None:
    print(f"\n[中止] {msg}", file=sys.stderr)
    sys.exit(code)


def load_plan() -> dict:
    if not PLAN_PATH.exists():
        die(f"找不到施工计划 {PLAN_PATH}")
    return json.loads(PLAN_PATH.read_text("utf-8"))


def load_schema() -> dict[str, dict]:
    if not SCHEMA_PATH.exists():
        die(f"找不到工具契约 {SCHEMA_PATH}\n先跑：node --import tsx promo/assets/dump-tools.ts")
    tools = json.loads(SCHEMA_PATH.read_text("utf-8"))["tools"]
    return {t["name"]: t for t in tools}


def read_token() -> str:
    f = BRIDGE_HOME / "token"
    if not f.exists():
        die(
            f"读不到本机凭据 {f}\n"
            "说明：桥接服务要求本机安装凭据。请先启动桌面端（npm start）让服务自己生成凭据，或确认 WPS_BRIDGE_HOME 是否正确。"
        )
    return f.read_text("utf-8").strip()


def call_tool(tool: str, args: dict, token: str, timeout: int = 180) -> dict:
    payload = json.dumps(
        {"name": tool, "arguments": args, "clientName": CLIENT_NAME}, ensure_ascii=False
    ).encode("utf-8")
    req = urllib.request.Request(
        f"{DEFAULT_BASE_URL}/api/v1/tool/call",
        data=payload,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": f"Bearer {token}",   # 不回显、不落盘
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        if e.code == 401:
            die("桥接返回 401：凭据不匹配。运行目录可能不是当前这份安装，检查 WPS_BRIDGE_HOME 与运行中的服务。")
        # 宿主层拒绝（422/500 等）也要走统一失败分支：先归档原始响应，再中止。
        # 早期版本在这里直接 die()，结果失败响应没落盘，而提示语却让人去看那个文件。
        try:
            parsed = json.loads(body)
            if not isinstance(parsed, dict):
                parsed = {"success": False, "error": body[:800]}
        except json.JSONDecodeError:
            parsed = {"success": False, "error": body[:800]}
        parsed["_httpStatus"] = e.code
        return parsed
    except urllib.error.URLError as e:
        die(f"连不上桥接服务 {DEFAULT_BASE_URL}：{e}\n请确认桌面端/后台服务在跑（node dist/bridge/cli.cjs --status）。")
    return {}


# ---------------------------------------------------------------------------
# 占位符与捕获
# ---------------------------------------------------------------------------
def resolve(node, variables: dict):
    """把 {{var}} 替换成实际值；未解析的变量直接报错，避免把 '{{x}}' 当参数发给宿主。"""
    if isinstance(node, dict):
        return {k: resolve(v, variables) for k, v in node.items()}
    if isinstance(node, list):
        return [resolve(v, variables) for v in node]
    if isinstance(node, str):
        out, i = "", 0
        while True:
            start = node.find("{{", i)
            if start < 0:
                out += node[i:]
                break
            end = node.find("}}", start)
            if end < 0:
                out += node[i:]
                break
            out += node[i:start]
            key = node[start + 2:end].strip()
            if key not in variables:
                raise KeyError(key)
            out += str(variables[key])
            i = end + 2
        return out
    return node


def deep_find(obj, keys: list[str]):
    """在响应里递归找第一个命中的键——不同宿主返回的包装层不一样，别硬编码路径。"""
    if isinstance(obj, dict):
        for k in keys:
            v = obj.get(k)
            if isinstance(v, str) and v.strip():
                return v
        for v in obj.values():
            found = deep_find(v, keys)
            if found:
                return found
    elif isinstance(obj, list):
        for v in obj:
            found = deep_find(v, keys)
            if found:
                return found
    return None


# ---------------------------------------------------------------------------
# 参数契约自检
# ---------------------------------------------------------------------------
def validate_step(step: dict, schema: dict[str, dict]) -> list[str]:
    problems: list[str] = []
    tool = step.get("tool")
    if tool not in schema:
        problems.append(f"工具不存在：{tool}")
        return problems
    props = schema[tool]["inputSchema"].get("properties", {}) or {}
    required = schema[tool]["inputSchema"].get("required", []) or []
    args = step.get("args", {}) or {}
    for key in args:
        if key not in props:
            problems.append(f"[{tool}] 参数不在契约里：{key}")
    for key in required:
        if key not in args:
            problems.append(f"[{tool}] 缺少必填参数：{key}")
    return problems


# ---------------------------------------------------------------------------
def iter_steps(plan: dict, groups: list[str]):
    for g in plan["groups"]:
        if groups and g["id"] not in groups:
            continue
        for s in g["steps"]:
            yield g, s


def summarise_response(resp: dict, limit: int = 260) -> str:
    text = json.dumps(resp, ensure_ascii=False)
    text = text.replace("\\n", " ")
    return text if len(text) <= limit else text[:limit] + f" …(共 {len(text)} 字符)"


def main() -> int:
    ap = argparse.ArgumentParser(description="按施工计划逐步驱动本机 WPS")
    ap.add_argument("--list", action="store_true", help="只列出步骤，不执行")
    ap.add_argument("--dry-run", action="store_true", help="只做参数契约自检，不连 WPS")
    ap.add_argument("--probe", action="store_true",
                    help="链路自检：读凭据并调用只读诊断工具 bridge_get_capabilities，不改任何文档")
    ap.add_argument("--group", action="append", default=[], choices=["sheet", "ppt", "doc"],
                    help="只跑某一组，可重复传")
    ap.add_argument("--all", action="store_true", help="三组依次跑")
    ap.add_argument("--only", action="append", default=[], help="只跑指定 step id，可重复传")
    ap.add_argument("--gate", action="store_true", help="每步等回车再走（录屏推荐）")
    ap.add_argument("--pace", type=float, default=0.0, help="每步之间停顿秒数")
    ap.add_argument("--skip-optional", action="store_true", help="跳过 optional 步骤")
    ap.add_argument("--workbook", default=None, help="表格作品的目标工作簿名（默认取 plan 里的值）")
    ap.add_argument("--var", action="append", default=[],
                    help="手工指定变量，形如 name=value；用于单步重录（例如 --var presentationName=演示文稿3）")
    ap.add_argument("--out", default=str(PROMO / "录制产出"), help="文件落盘目录与响应留档目录")
    ap.add_argument("--continue-on-error", action="store_true", help="某步失败仍继续（默认停下）")
    ap.add_argument("--show-args", action="store_true", help="打印每一步的完整参数")
    args = ap.parse_args()

    plan = load_plan()
    schema = load_schema()
    out_dir = Path(args.out).resolve()
    variables = {
        "workbookName": args.workbook or plan["meta"]["vars"]["workbookName"],
        "outDir": str(out_dir),
    }
    for item in args.var:
        if "=" not in item:
            die(f"--var 需要 name=value 形式，收到：{item}")
        key, value = item.split("=", 1)
        variables[key.strip()] = value

    groups = args.group or ([] if args.all or args.only else ["sheet"])
    if args.all:
        groups = ["sheet", "ppt", "doc"]
    steps = list(iter_steps(plan, groups))
    if args.only:
        steps = [(g, s) for g, s in steps if s["id"] in args.only]
        if not steps:
            die(f"--only 没有匹配到任何步骤：{args.only}")
    if args.skip_optional:
        steps = [(g, s) for g, s in steps if not s.get("optional")]

    # ---------------- 列表模式 ----------------
    if args.list:
        print(f"施工计划 v{plan['meta']['version']} · 共 {len(steps)} 步\n")
        last_group = None
        for g, s in steps:
            if g["id"] != last_group:
                print(f"\n== {g['title']} ==")
                last_group = g["id"]
            flag = "（选拍）" if s.get("optional") else ""
            print(f"  {s['id']:<11} {s['title']}{flag}")
        return 0

    # ---------------- 自检模式 ----------------
    if args.dry_run:
        print(f"参数契约自检：{len(steps)} 步\n")
        bad = 0
        for g, s in steps:
            problems = validate_step(s, schema)
            if problems:
                bad += 1
                print(f"  ✗ {s['id']} {s['title']}")
                for p in problems:
                    print(f"      - {p}")
        if bad:
            print(f"\n自检失败：{bad} 步有问题，先修 plan.json 再跑。")
            return 1
        print("  全部通过：工具名与参数都符合运行时契约。")
        print(f"  说明：本模式只校验参数名与必填项，不校验语义（例如区域地址是否落在表内）。")
        return 0

    # ---------------- 链路自检模式 ----------------
    if args.probe:
        token = read_token()
        print(f"链路自检：{DEFAULT_BASE_URL}（只读诊断工具 bridge_get_capabilities，不改任何文档）\n")
        resp = call_tool("bridge_get_capabilities", {}, token)
        ok = bool(resp.get("success"))
        data = resp.get("data") if isinstance(resp, dict) else None
        print(f"  鉴权与连通：{'通过' if ok else '失败'}")
        print(f"  响应顶层字段：{list(resp.keys()) if isinstance(resp, dict) else type(resp).__name__}")
        if isinstance(data, dict):
            for key in ("platform", "connected", "host", "hostName", "version", "tools", "rollback"):
                if key in data:
                    value = data[key]
                    text = json.dumps(value, ensure_ascii=False)
                    print(f"  {key}: {text if len(text) <= 120 else text[:120] + ' …'}")
        print(f"\n  原始响应（截断）：{summarise_response(resp, 300)}")
        if not ok:
            print("\n  链路不通时：确认桌面端在跑（npm start）、加载项已装、WPS 已打开。")
        return 0 if ok else 1

    # ---------------- 执行模式 ----------------
    token = read_token()
    out_dir.mkdir(parents=True, exist_ok=True)
    resp_dir = out_dir / "responses"
    resp_dir.mkdir(parents=True, exist_ok=True)

    total = len(steps)
    print(f"施工计划 v{plan['meta']['version']} · 本次执行 {total} 步")
    print(f"目标：{DEFAULT_BASE_URL}   落盘目录：{out_dir}")
    print("=" * 78)

    done = 0
    for idx, (g, s) in enumerate(steps, start=1):
        head = f"[{idx}/{total}] {s['id']} · {s['title']}"
        print(f"\n{head}\n{'-' * len(head)}")
        print(f"  工具：{s['tool']}")
        print(f"  镜头：{s.get('camera', '')}")

        if args.gate:
            try:
                input("  ▶ 按回车执行这一步（Ctrl+C 中止）… ")
            except (EOFError, KeyboardInterrupt):
                print("\n已中止。")
                return 130

        try:
            resolved = resolve(s.get("args", {}) or {}, variables)
        except KeyError as e:
            msg = (
                f"参数里的变量 {{{{{e.args[0]}}}}} 还没被解析出来。\n"
                f"  通常是因为上一步没跑（例如 {{presentationName}} 由 ppt-01 捕获）。\n"
                f"  解决：连带上一步一起跑，或先单独跑到捕获那一步。"
            )
            if args.continue_on_error:
                print(f"  ✗ {msg}")
                continue
            die(msg)

        if args.show_args:
            print(f"  参数：{json.dumps(resolved, ensure_ascii=False)}")

        resp = call_tool(s["tool"], resolved, token)
        (resp_dir / f"{s['id']}.json").write_text(
            json.dumps({"step": s["id"], "tool": s["tool"], "args": resolved, "response": resp},
                       ensure_ascii=False, indent=2),
            "utf-8",
        )

        ok = bool(resp.get("success")) if isinstance(resp, dict) else False
        print(f"  {'✓ 完成' if ok else '✗ 失败'}：{summarise_response(resp)}")

        if not ok:
            if args.continue_on_error:
                continue
            die(
                f"第 {idx} 步（{s['id']}）失败，已停下。\n"
                f"  原始响应已存：{resp_dir / (s['id'] + '.json')}\n"
                "  常见原因：WPS 里没打开目标文件 / 文件名与 plan 里不一致 / 加载项被阻断。\n"
                "  先跑 node dist/bridge/cli.cjs --status 和 --doctor 复核。"
            )

        cap = s.get("capture")
        if cap:
            for var, keys in cap.items():
                found = deep_find(resp, keys)
                if found:
                    variables[var] = found
                    print(f"  已捕获 {var} = {found}")
                else:
                    print(f"  ! 没能在响应里找到 {var}（候选键：{keys}）")
                    print(f"    后续步骤若引用 {{{{{var}}}}} 会停下。可先看 {resp_dir / (s['id'] + '.json')}，")
                    print(f"    再用 --only 配合手工值重跑后续步骤。")

        done += 1
        if args.pace and idx < total:
            time.sleep(args.pace)

    print("\n" + "=" * 78)
    print(f"完成 {done}/{total} 步。原始响应留档：{resp_dir}")
    print("重录某一步：python3 promo/施工/build.py --only <step-id> --gate")
    return 0


if __name__ == "__main__":
    sys.exit(main())
