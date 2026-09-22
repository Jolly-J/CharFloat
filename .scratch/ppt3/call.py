#!/usr/bin/env python3
"""极简 bridge HTTP 客户端：给 .scratch/ppt3 下的实验脚本复用。"""
import json
import sys
import urllib.request
from pathlib import Path

TOKEN = Path.home().joinpath(".wps-bridge/token").read_text().strip()
URL = "http://127.0.0.1:19890/api/v1/tool/call"


def call(name, arguments=None, timeout=120):
    body = json.dumps({"name": name, "arguments": arguments or {}}).encode("utf-8")
    req = urllib.request.Request(
        URL,
        data=body,
        headers={"Authorization": "Bearer " + TOKEN, "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return json.loads(e.read().decode("utf-8"))


def main():
    name = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    out = call(name, args)
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
