#!/usr/bin/env python3
"""apimap 只读探测客户端。禁止打印凭据。用法：
  python3 probe.py inspect '<expression>' [component]
  python3 probe.py script <file.js> [component]
"""
import json
import sys
import urllib.request
import urllib.error
from pathlib import Path

TOKEN = (Path.home() / '.wps-bridge' / 'token').read_text().strip()
URL = 'http://127.0.0.1:19890/api/v1/tool/call'


def call(name, arguments):
    body = json.dumps({'name': name, 'arguments': arguments,
                       'clientName': 'apimap', 'sessionId': 'apimap-session'}).encode()
    req = urllib.request.Request(URL, headers={
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + TOKEN}, data=body)
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return {'success': False, 'error': 'HTTP %s: %s' % (e.code, e.read().decode()[:800])}
    except Exception as e:
        return {'success': False, 'error': '%s: %s' % (type(e).__name__, e)}


def main():
    mode = sys.argv[1]
    if mode == 'inspect':
        args = {'expression': sys.argv[2]}
        if len(sys.argv) > 3:
            args['component'] = sys.argv[3]
        out = call('wps_inspect_api', args)
    else:
        code = Path(sys.argv[2]).read_text()
        args = {'code': code}
        if len(sys.argv) > 3:
            args['component'] = sys.argv[3]
        out = call('wps_execute_script', args)
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
