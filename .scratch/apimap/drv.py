#!/usr/bin/env python3
"""apimap 探测驱动（统一入口，方便重复调用）。"""
import json
import sys
import urllib.request
import urllib.error
from pathlib import Path

TOKEN = (Path.home() / '.wps-bridge' / 'token').read_text().strip()
URL = 'http://127.0.0.1:19890/api/v1/tool/call'
HERE = Path(__file__).parent


def call(name, arguments, timeout=240):
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


XL = 'apimap-scratch.xlsx'
WD = 'apimap-scratch.docx'
PP = 'apimap-scratch.pptx'


def _target(component):
    if component == 'excel':
        return {'workbookName': XL}
    if component == 'word':
        return {'documentName': WD}
    if component == 'ppt':
        return {'presentationName': PP}
    return {}


def script(component, code, params=None, **extra):
    args = {'code': code, 'component': component}
    if params is not None:
        args['params'] = params
    args.update(_target(component))
    args.update(extra)
    return call('wps_execute_script', args)


def reflect(component, exprs, opts=None):
    return script(component, (HERE / 'reflect.js').read_text(),
                  {'exprs': exprs, 'opts': opts or {}})


def evals(component, pairs, **extra):
    return script(component, (HERE / 'evalx.js').read_text(),
                  {'items': [{'k': k, 'e': e} for k, e in pairs]}, **extra)


def show(resp, raw=False):
    if not resp.get('success'):
        print('!! CALL FAILED: %s' % resp.get('error'))
        return None
    d = resp.get('data') or {}
    if not d.get('success'):
        print('!! SCRIPT FAILED: %s' % json.dumps(d, ensure_ascii=False)[:1500])
        return None
    rv = d.get('returnValue')
    if raw:
        print(json.dumps(rv, ensure_ascii=False, indent=2))
    elif isinstance(rv, list):
        for line in rv:
            print(line)
    elif isinstance(rv, dict) and 'results' in rv:
        for r in rv['results']:
            print('=' * 72)
            print('EXPR: %s' % r.get('e'))
            if r.get('err'):
                print('  !! ERROR: %s' % r['err'])
                continue
            print('  type=%s members=%s' % (r.get('t'), r.get('n')))
            if r.get('fe'):
                print('  forInErr: %s' % r['fe'])
            if r.get('re'):
                print('  readErrs: %s' % r['re'])
            names = [m for m in (r.get('m') or '').split(',') if m]
            # 按类型分组输出，便于阅读
            fns = [m.split('|')[0] for m in names if m.split('|')[1:2] == ['function']]
            props = [m for m in names if m.split('|')[1:2] != ['function']]
            print('  -- 方法(%d): %s' % (len(fns), ', '.join(fns)))
            print('  -- 属性(%d): %s' % (len(props), ', '.join(props)))
    else:
        print(json.dumps(rv, ensure_ascii=False, indent=2)[:8000])
    print('--- execMs=%s' % d.get('executionTimeMs'))
    return rv


def refl(component, exprs, opts=None):
    return show(reflect(component, exprs, opts))


def ev(component, pairs, **extra):
    return show(evals(component, pairs, **extra))
