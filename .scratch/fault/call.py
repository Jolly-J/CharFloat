#!/usr/bin/env python3
"""精简调用：call.py <tool> <json> [client] -> 打印 isError 与 text 原文"""
import json, sys
import mcp
tool = sys.argv[1]
args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
client = sys.argv[3] if len(sys.argv) > 3 else 'fault-probe'
r = mcp.one(tool, args, client)
res = r.get('result', r)
err = res.get('isError')
txt = ''
try:
    txt = res['content'][0]['text']
except Exception:
    txt = json.dumps(res, ensure_ascii=False)
print('[tool] %s' % tool)
print('[isError] %s' % err)
print('[raw] %s' % txt)
