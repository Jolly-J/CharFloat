#!/usr/bin/env python3
"""Thin wrapper around the bridge HTTP API for the PPT job. Usage: call.py <tool> '<json>'"""
import json
import urllib.error
import sys
import urllib.request
import uuid
from pathlib import Path

TOKEN = Path.home().joinpath('.wps-bridge/token').read_text().strip()


def call(name, arguments=None):
    body = {
        'name': name,
        'arguments': arguments or {},
        'clientName': 'PPT-Validator',
        'sessionId': str(uuid.uuid4()),
    }
    req = urllib.request.Request(
        'http://127.0.0.1:19890/api/v1/tool/call',
        headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + TOKEN},
        data=json.dumps(body).encode(),
    )
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        raise RuntimeError('HTTP %s: %s' % (e.code, e.read().decode()[:1500])) from None


if __name__ == '__main__':
    tool = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    out = call(tool, args)
    print(json.dumps(out, ensure_ascii=False, indent=2))
