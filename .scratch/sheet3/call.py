#!/usr/bin/env python3
"""Call a bridge tool, print raw JSON. Usage: call.py <tool> '<json-args>'"""
import json, sys, os
from pathlib import Path
import urllib.request, urllib.error

port = int(os.environ.get('WPS_BRIDGE_PORT', '19890'))
token = (Path.home() / '.wps-bridge' / 'token').read_text().strip()

def call(tool, args):
    body = json.dumps({'name': tool, 'arguments': args}).encode()
    req = urllib.request.Request('http://127.0.0.1:%d/api/v1/tool/call' % port,
        headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token}, data=body)
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return {'success': False, '_http': e.code, 'error': e.read().decode()[:4000]}

if __name__ == '__main__':
    tool = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    out = call(tool, args)
    print(json.dumps(out, ensure_ascii=False, indent=2))
