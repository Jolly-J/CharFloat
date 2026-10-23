#!/usr/bin/env python3
"""Local Bridge client. Never print connection credentials."""
import argparse
import base64
import socket
import json
import os
from pathlib import Path
import subprocess
import sys
import urllib.request
import urllib.error
import uuid

def home_dir():
    return Path(os.environ.get('WPS_BRIDGE_HOME') or (str(Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'WPSBridge') if sys.platform == 'win32' else str(Path.home() / '.wps-bridge')))

class BridgeFailure(RuntimeError):
    def __init__(self, detail):
        self.detail = detail
        super().__init__(str(detail.get('error', 'Bridge call failed')))


def failure_detail(result):
    detail = dict(result) if isinstance(result, dict) else {'error': str(result)}
    detail['success'] = False
    if isinstance(detail.get('errorDetails'), dict):
        for key in ['kind', 'channel', 'method', 'executed', 'unsafe']:
            if key in detail['errorDetails']:
                detail[key] = detail['errorDetails'][key]
    detail.setdefault('executed', 'unknown')
    if detail['executed'] != 'no':
        detail['recovery'] = '先读回目标确认是否已执行；不要自动重放写操作。'
    return detail


def request(route, body=None, auth=True):
    port = int(os.environ.get('WPS_BRIDGE_PORT', '19890'))
    if not 1024 <= port <= 65535:
        raise ValueError('Invalid Bridge port')
    headers = {'Content-Type': 'application/json'}
    if auth:
        headers['Authorization'] = 'Bearer ' + (home_dir() / 'token').read_text().strip()
    req = urllib.request.Request('http://127.0.0.1:%d%s' % (port, route), headers=headers, data=None if body is None else json.dumps(body).encode())
    try:
        with urllib.request.urlopen(req, timeout=35) as r:
            result = json.load(r)
    except urllib.error.HTTPError as e:
        raw = e.read().decode('utf-8', errors='replace')
        try:
            detail = json.loads(raw)
        except ValueError:
            detail = {'error': 'Bridge HTTP %s: %s' % (e.code, raw[:600])}
        raise BridgeFailure(failure_detail(detail)) from None
    except json.JSONDecodeError as e:
        raise BridgeFailure(failure_detail({'error': 'Bridge 返回非 JSON 响应', 'kind': 'transport', 'executed': 'unknown'})) from None
    except (TimeoutError, socket.timeout, urllib.error.URLError) as e:
        raise BridgeFailure(failure_detail({'error': str(e), 'kind': 'transport', 'executed': 'unknown'})) from None
    if isinstance(result, dict) and result.get('success') is False:
        raise BridgeFailure(failure_detail(result))
    return result

def prepare_result(result, preview_output=None):
    """Legacy response is unchanged; explicit preview output opts into compact metadata."""
    if not isinstance(result, dict):
        return result
    if not preview_output:
        return result
    out = dict(result)
    if isinstance(out.get('data'), dict):
        out['data'] = prepare_result(out['data'], preview_output)
        return out
    encoded = out.pop('imageBase64', None)
    if encoded is not None:
        if preview_output:
            raw = base64.b64decode(encoded, validate=True)
            if not raw.startswith(b'\x89PNG\r\n\x1a\n'):
                raise ValueError('预览不是 PNG；未写入图片文件')
            target = Path(preview_output)
            with target.open('xb') as f:
                f.write(raw)
            out['previewFile'] = str(target)
            out['previewBytes'] = len(raw)
            if len(raw) >= 24 and raw[12:16] == b'IHDR':
                out['previewWidth'] = int.from_bytes(raw[16:20], 'big')
                out['previewHeight'] = int.from_bytes(raw[20:24], 'big')
    elif preview_output:
        raise ValueError('响应没有 imageBase64；未生成预览文件，请检查 imageUnavailable/warnings')
    return out


def emit_result(result, output=None, stream=None):
    text = json.dumps(result, ensure_ascii=False, indent=2) + '\n'
    if output:
        with Path(output).open('x', encoding='utf-8') as f:
            f.write(text)
    else:
        print(text, end='', file=stream or sys.stdout)


def emit_local_output_failure(result, args, exc, host_succeeded):
    """Output errors never discard the execution uncertainty or trigger another call."""
    data = result.get('data', result) if isinstance(result, dict) else {}
    summary = {key: data[key] for key in ['slideIndex', 'slideId', 'shapeCount', 'verifiedSavedFlag']
               if isinstance(data, dict) and isinstance(data.get(key), (int, float, bool))}
    detail = {
        'success': False, 'phase': 'local_output', 'hostCallReturnedSuccess': host_succeeded,
        'executed': 'unknown', 'doNotReplay': True,
        'error': '宿主调用已返回' + ('成功' if host_succeeded else '失败') + '，随后本地输出失败；结果文件或预览未完成。',
        'outputErrorType': type(exc).__name__,
        'outputPath': args.output, 'previewPath': args.preview_output,
        'resultSummary': summary,
        'recovery': '不要盲目重放工具；保留现有文件，先读回目标。'
    }
    if not host_succeeded and isinstance(result, dict):
        if isinstance(result.get('error'), str):
            detail['hostError'] = result['error'][:600]
        source = result.get('errorDetails')
        if isinstance(source, dict):
            detail['errorDetails'] = {k: source[k] for k in ['kind', 'channel', 'method', 'executed', 'unsafe', 'message']
                                      if isinstance(source.get(k), (str, bool))}
        detail['executed'] = result.get('executed') if result.get('executed') in ['no', 'unknown'] else 'unknown'
        partial = result.get('partial')
        if isinstance(partial, dict):
            detail['partial'] = {k: partial[k] for k in ['slideId', 'shapeId', 'completedCount', 'failedIndex']
                                 if isinstance(partial.get(k), (int, float))}
            for key in ['completedShapeIds', 'completedSlideIds']:
                if isinstance(partial.get(key), list):
                    detail['partial'][key] = [x for x in partial[key][:2000] if isinstance(x, (int, float))]
    if not host_succeeded and isinstance(result, dict) and result.get('diagnosticType') == 'ppt_layout':
        fields = ['slideWidth', 'slideHeight', 'slideSize', 'slideOrientation', 'designCount', 'layoutCount', 'slideCount']
        for key in ['before', 'after', 'applied']:
            source = result.get(key)
            if isinstance(source, dict):
                detail[key] = {k: source[k] for k in fields if k in source and (source[k] is None or isinstance(source[k], (int, float)))}
                if key == 'applied' and source.get('templateApplied') is True:
                    detail[key]['templateApplied'] = True
                if key == 'applied' and isinstance(source.get('layout'), dict) and isinstance(source['layout'].get('appliedPages'), (int, float)):
                    detail[key]['layout'] = {'appliedPages': source['layout']['appliedPages']}
        categories = {'write_failed', 'readback_mismatch', 'apply_failed', 'invalid_argument', 'unclassified_warning'}
        reasons = {'host_rejected', 'unsupported', 'not_found', 'invalid_argument', 'host_failure'}
        field_names = {'slideWidth', 'slideHeight', 'slideOrientation', 'template', 'layout', 'preset', 'unknown'}
        detail['warningDetails'] = [{'field': x['field'], 'code': x['code'], 'reason': x['reason']}
                                    for x in result.get('warningDetails', [])[:200]
                                    if isinstance(x, dict) and x.get('field') in field_names and x.get('code') in categories and x.get('reason') in reasons]
    try:
        emit_result(detail, stream=sys.stderr)
    except OSError:
        # stderr itself unavailable: still fail without retrying a host operation.
        pass


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--doctor', action='store_true')
    parser.add_argument('--start', action='store_true')
    parser.add_argument('--repair-addon', action='store_true')
    parser.add_argument('tool', nargs='?')
    parser.add_argument('arguments', nargs='?')
    query = parser.add_mutually_exclusive_group()
    query.add_argument('--list-tools', nargs='?', const='', metavar='PREFIX', help='只列工具名/描述，可按前缀过滤')
    query.add_argument('--schema', metavar='TOOL', help='查询单个工具的真实 schema')
    parser.add_argument('--session-id', help='同一任务复用的非敏感会话标识；不同任务应隔离，默认每次生成')
    parser.add_argument('--args-file', help='UTF-8 JSON 对象参数文件，与位置参数 JSON 二选一')
    parser.add_argument('--output', help='结果 JSON 文件（拒绝覆盖）；原调用成功结构保持兼容，显式预览输出才省略 base64')
    parser.add_argument('--preview-output', help='PPT 预览 PNG 文件（拒绝覆盖），仅用于 wps_ppt_capture_slide_preview')
    args = parser.parse_args(argv)
    if args.session_id is not None and (not args.session_id.strip() or len(args.session_id) > 128):
        parser.error('--session-id 需为1至128字符的非空任务标识')
    querying = args.list_tools is not None or args.schema is not None
    if querying and (args.tool or args.args_file or args.preview_output or args.doctor or args.start or args.repair_addon):
        parser.error('工具查询不能与执行或服务选项混用')
    if args.args_file and args.arguments is not None:
        parser.error('--args-file 与位置参数 JSON 不能同时传入')
    if args.args_file and not args.tool:
        parser.error('--args-file 需要工具名')
    if args.preview_output and args.tool != 'wps_ppt_capture_slide_preview':
        parser.error('--preview-output 仅用于 wps_ppt_capture_slide_preview')
    if args.preview_output and args.output and Path(args.preview_output).resolve() == Path(args.output).resolve():
        parser.error('图片与结果文件不能使用同一路径')
    for filename in [args.output, args.preview_output]:
        if filename and Path(filename).exists():
            parser.error('输出文件已存在，拒绝覆盖: ' + filename)
    if querying:
        tools = request('/api/v1/mcp-tools')
        if args.schema is not None:
            tool = next((t for t in tools if t.get('name') == args.schema), None)
            if tool is None:
                raise ValueError('未知工具: ' + args.schema)
            emit_result(tool, args.output)
        else:
            emit_result([{'name': t['name'], 'description': t.get('description', '')} for t in tools if t['name'].startswith(args.list_tools)], args.output)
        return
    if args.start or args.repair_addon:
        file = home_dir() / 'installation.json'
        if not file.exists():
            raise RuntimeError('Missing installation record. Open the installed 字浮 CharFloat app once, or follow the source README.')
        installation = json.loads(file.read_text())
        command, entry = installation['executable'], installation['cli']
        if not Path(command).is_file() or not Path(entry).is_file():
            raise RuntimeError('Installed executable or CLI is missing. Repair the Bridge application first.')
        env = dict(os.environ, ELECTRON_RUN_AS_NODE='1', WPS_BRIDGE_RESOURCES=installation['resources'])
        subprocess.run([command, entry, '--repair-addon' if args.repair_addon else '--start'], env=env, check=True, timeout=45)
        return
    if args.doctor:
        result = {'installationRecord': (home_dir() / 'installation.json').exists(), 'credentialsPresent': (home_dir() / 'token').exists()}
        try:
            health = request('/health', auth=False)
            result['service'] = health
            if health.get('service') != 'wps-bridge' or health.get('protocol') != 2:
                raise RuntimeError('Service identity or version mismatch; do not terminate an unknown process.')
            result['diagnosis'] = request('/api/v1/tool/call', {'name': 'bridge_diagnose', 'arguments': {}})['data']
        except Exception as e:
            result['error'] = str(e)
        emit_result(result, args.output)
    elif args.tool:
        arguments = json.loads(Path(args.args_file).read_text(encoding='utf-8') if args.args_file else (args.arguments or '{}'))
        if not isinstance(arguments, dict):
            raise ValueError('工具参数必须是 JSON 对象')
        try:
            result = request('/api/v1/tool/call', {'name': args.tool, 'arguments': arguments, 'clientName': 'Bridge Skill', 'sessionId': args.session_id or str(uuid.uuid4())})
            if isinstance(result, dict) and isinstance(result.get('data'), dict) and result['data'].get('success') is False:
                raise BridgeFailure(failure_detail(result['data']))
        except BridgeFailure as e:
            try:
                emit_result(e.detail, args.output, sys.stderr)
            except Exception as exc:
                emit_local_output_failure(e.detail, args, exc, False)
            return 1
        try:
            emit_result(prepare_result(result, args.preview_output), args.output)
        except Exception as exc:
            emit_local_output_failure(result, args, exc, True)
            return 1
    else:
        parser.print_help()
if __name__ == '__main__':
    try:
        sys.exit(main() or 0)
    except BridgeFailure as exc:
        emit_result(exc.detail, stream=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)
