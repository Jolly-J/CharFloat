#!/usr/bin/env python3
"""临时调用助手：把 MCP 工具调用结果打到 stdout。用法:
  python3 .scratch/word/wps.py <tool> '<json-args>' [--raw]
"""
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

TOKEN = (Path.home() / '.wps-bridge' / 'token').read_text().strip()
URL = 'http://127.0.0.1:19890/api/v1/tool/call'


def call(name, arguments):
    body = {'name': name, 'arguments': arguments, 'clientName': 'WPS Testdoc Author', 'sessionId': 'scratch-word-author'}
    req = urllib.request.Request(URL, headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + TOKEN}, data=json.dumps(body).encode())
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return {'httpError': e.code, 'body': e.read().decode()[:4000]}
    except Exception as e:  # noqa: BLE001
        return {'transportError': repr(e)}


def main():
    tool = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    result = call(tool, args)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
