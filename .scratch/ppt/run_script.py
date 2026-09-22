#!/usr/bin/env python3
"""Run a JS file against the WPS host through wps_execute_script.

Usage: run_script.py js/<file>.js [params.json]
The files are edited on disk instead of inlined into shell quoting.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from call import call

HERE = Path(__file__).parent
js_file = HERE / sys.argv[1]
params = json.loads((HERE / sys.argv[2]).read_text()) if len(sys.argv) > 2 else {}

res = call('wps_execute_script', {
    'component': 'ppt',
    'presentationName': params.get('presentationName', '测试演示文稿.pptx'),
    'code': js_file.read_text(),
    'params': params,
})
data = res.get('data', {})
print(json.dumps(data.get('returnValue', data), ensure_ascii=False, indent=1))
