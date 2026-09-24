#!/usr/bin/env python3
"""Local Bridge client. Never print connection credentials."""
import argparse
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
        raise RuntimeError('Bridge HTTP %s: %s' % (e.code, e.read().decode()[:600])) from None
    if isinstance(result, dict) and result.get('success') is False:
        raise RuntimeError(result.get('error', 'Bridge call failed'))
    return result

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--doctor', action='store_true')
    parser.add_argument('--start', action='store_true')
    parser.add_argument('--repair-addon', action='store_true')
    parser.add_argument('tool', nargs='?')
    parser.add_argument('arguments', nargs='?', default='{}')
    args = parser.parse_args()
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
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.tool:
        print(json.dumps(request('/api/v1/tool/call', {'name': args.tool, 'arguments': json.loads(args.arguments), 'clientName': 'Bridge Skill', 'sessionId': str(uuid.uuid4())}), ensure_ascii=False, indent=2))
    else:
        parser.print_help()
if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)
