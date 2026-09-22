#!/usr/bin/env python3
"""故障实测用 MCP 客户端：显式会话，原样打印服务端返回。"""
import json, os, sys, urllib.request, urllib.error
from pathlib import Path

TOKEN = (Path.home() / '.wps-bridge' / 'token').read_text().strip()
PORT = int(os.environ.get('WPS_BRIDGE_PORT', '19890'))
URL = 'http://127.0.0.1:%d/mcp' % PORT


def _post(body, sid=None):
    h = {'Content-Type': 'application/json', 'Authorization': 'Bearer ' + TOKEN,
         'Accept': 'application/json, text/event-stream'}
    if sid:
        h['mcp-session-id'] = sid
    req = urllib.request.Request(URL, headers=h, data=json.dumps(body).encode())
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            return r.headers, r.read().decode()
    except urllib.error.HTTPError as e:
        return e.headers, 'HTTP %s: %s' % (e.code, e.read().decode()[:2000])


def _parse(raw):
    raw = raw.strip()
    if raw.startswith('event:') or raw.startswith('data:'):
        for line in raw.splitlines():
            if line.startswith('data:'):
                raw = line[5:].strip()
    try:
        return json.loads(raw)
    except Exception:
        return raw


class Session:
    def __init__(self, client='fault-probe'):
        self.client = client
        h, b = _post({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
            "protocolVersion": "2024-11-05", "capabilities": {},
            "clientInfo": {"name": client, "version": "1.0"}}})
        self.sid = h.get('mcp-session-id')
        _post({"jsonrpc": "2.0", "method": "notifications/initialized"}, self.sid)
        self._id = 10

    def call(self, name, arguments=None):
        self._id += 1
        h, b = _post({"jsonrpc": "2.0", "id": self._id, "method": "tools/call",
                      "params": {"name": name, "arguments": arguments or {}}}, self.sid)
        return _parse(b)


def one(name, arguments=None, client='fault-probe'):
    s = Session(client)
    return s.call(name, arguments)


if __name__ == '__main__':
    name = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    out = one(name, args, os.environ.get('FAULT_CLIENT', 'fault-probe'))
    print(json.dumps(out, ensure_ascii=False, indent=1))
