#!/usr/bin/env python3
"""apimap 批量反射驱动。
用法：
  python3 reflect.py <component> '<expr1>' '<expr2>' ...
  python3 reflect.py <component> --limit 40 '<expr>' ...
输出：把 "name|type|count" 解析成可读表格。
"""
import json
import sys
import urllib.request
import urllib.error
from pathlib import Path

TOKEN = (Path.home() / '.wps-bridge' / 'token').read_text().strip()
URL = 'http://127.0.0.1:19890/api/v1/tool/call'
HERE = Path(__file__).parent


def call(name, arguments, timeout=180):
    body = json.dumps({'name': name, 'arguments': arguments,
                       'clientName': 'apimap', 'sessionId': 'apimap-session'}).encode()
    req = urllib.request.Request(URL, headers={
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + TOKEN}, data=body)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return {'success': False, 'error': 'HTTP %s: %s' % (e.code, e.read().decode()[:800])}
    except Exception as e:
        return {'success': False, 'error': '%s: %s' % (type(e).__name__, e)}


def reflect(component, exprs, opts=None, extra=None):
    args = {'code': (HERE / 'reflect.js').read_text(), 'component': component,
            'params': {'exprs': exprs, 'opts': opts or {}}}
    if extra:
        args.update(extra)
    return call('wps_execute_script', args)


def run_script(component, code, params=None, extra=None):
    args = {'code': code, 'component': component}
    if params is not None:
        args['params'] = params
    if extra:
        args.update(extra)
    return call('wps_execute_script', args)


def render(resp):
    if not resp.get('success'):
        print('!! CALL FAILED: %s' % resp.get('error'))
        return
    d = resp.get('data') or {}
    if not d.get('success'):
        print('!! SCRIPT FAILED: %s' % json.dumps(d, ensure_ascii=False)[:600])
        return
    rv = d.get('returnValue')
    if isinstance(rv, dict) and 'results' in rv:
        for r in rv['results']:
            print('=' * 70)
            print('EXPR: %s' % r.get('e'))
            if r.get('err'):
                print('  !! ERROR: %s' % r['err'])
                continue
            print('  type=%s members=%s' % (r.get('t'), r.get('n')))
            if r.get('fe'):
                print('  forInErr: %s' % r['fe'])
            if r.get('re'):
                print('  readErrs: %s' % r['re'])
            for m in (r.get('m') or '').split(','):
                if not m:
                    continue
                parts = m.split('|')
                n = parts[0]
                t = parts[1] if len(parts) > 1 else '?'
                c = parts[2] if len(parts) > 2 else ''
                mark = '()' if t == 'function' else ''
                print('    - %s%s [%s]%s' % (n, mark, t, (' count=%s' % c) if c else ''))
    else:
        print(json.dumps(rv, ensure_ascii=False, indent=2)[:6000])
    print('--- execMs=%s' % d.get('executionTimeMs'))


if __name__ == '__main__':
    comp = sys.argv[1]
    rest = sys.argv[2:]
    opts = {}
    if rest and rest[0] == '--limit':
        opts['limit'] = int(rest[1])
        rest = rest[2:]
    render(reflect(comp, rest, opts))
