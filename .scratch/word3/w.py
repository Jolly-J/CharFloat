#!/usr/bin/env python3
"""Word 高级能力实测调用器：走本地 HTTP /api/v1/tool/call。"""
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

PORT = 19890
TOKEN = (Path.home() / ".wps-bridge" / "token").read_text().strip()


def call(name, arguments, timeout=90):
    body = json.dumps({"name": name, "arguments": arguments,
                       "clientName": "word3-sweep", "sessionId": "word3-sweep-session"}).encode()
    req = urllib.request.Request(
        "http://127.0.0.1:%d/api/v1/tool/call" % PORT,
        headers={"Content-Type": "application/json", "Authorization": "Bearer " + TOKEN},
        data=body)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return {"success": False, "httpError": e.code, "raw": e.read().decode()[:2000]}
    except Exception as e:  # noqa: BLE001
        return {"success": False, "transportError": repr(e)}


def script(code, params=None, timeout=90, doc="测试文字文稿.docx"):
    return call("wps_execute_script",
                {"code": code, "component": "word", "documentName": doc, "params": params or {}},
                timeout=timeout)


def run(code, params=None, doc="测试文字文稿.docx", timeout=90):
    """执行脚本并返回 returnValue（自动解包双层 data）。"""
    r = script(code, params, timeout, doc)
    if not r.get("success"):
        return {"__bridgeError__": r}
    d = r.get("data", {})
    if not d.get("success"):
        return {"__scriptError__": d}
    return d.get("returnValue")


def show(obj, limit=8000):
    text = json.dumps(obj, ensure_ascii=False, indent=2)
    print(text[:limit])
    if len(text) > limit:
        print("...[truncated %d]" % (len(text) - limit))


if __name__ == "__main__":
    show(call(sys.argv[1], json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}))
