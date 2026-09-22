#!/usr/bin/env python3
"""Export slide previews to .scratch/ppt/preview/slide-NN.png (no base64 in stdout)."""
import base64
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from call import call

OUT = Path(__file__).parent / 'preview'
OUT.mkdir(exist_ok=True)

indices = [int(x) for x in sys.argv[1:]]
for idx in indices:
    res = call('wps_ppt_capture_slide_preview', {'presentationName': '测试演示文稿.pptx', 'slideIndex': idx})
    data = res.get('data', {})
    b64 = data.get('imageBase64')
    if not b64:
        print(json.dumps({'slide': idx, 'ERROR': data}, ensure_ascii=False))
        continue
    raw = base64.b64decode(b64)
    path = OUT / ('slide-%02d.png' % idx)
    path.write_bytes(raw)
    print(json.dumps({'slide': idx, 'path': str(path), 'bytes': len(raw),
                      'imagePath': data.get('imagePath')}, ensure_ascii=False))
