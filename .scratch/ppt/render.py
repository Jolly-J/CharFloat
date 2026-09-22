#!/usr/bin/env python3
"""Render the whole deck to PNG through the WPS host and copy images into .scratch/ppt/preview/.

Why: wps_ppt_capture_slide_preview fails on this host ("PPT 未生成预览") because
Slide.Export silently no-ops here. Presentation.SaveCopyAs(dir, 18) (ppSaveAsPNG)
does produce one PNG per slide, so we use that as the render path.
"""
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from call import call

HERE = Path(__file__).parent
OUT = HERE / 'preview'
OUT.mkdir(exist_ok=True)
STAMP = time.strftime('%H%M%S')
# NOTE: WPS derives the export format from the extension and strips it to build the
# folder name, so the path must end with ".png" (folder "render-X" is then created).
REMOTE = '/Users/jolin/.wps-bridge/previews/render-%s.png' % STAMP

pres = sys.argv[1] if len(sys.argv) > 1 else '测试演示文稿.pptx'
code = (
    'const dir = params.dir;'
    'try { pres.SaveCopyAs(dir, 18); } catch (e) { return {ok:false, err:String(e && e.message || e)}; }'
    'return {ok:true, dir:dir, slides:pres.Slides.Count};'
)
res = call('wps_execute_script', {'component': 'ppt', 'presentationName': pres, 'code': code,
                                  'params': {'dir': REMOTE}})
rv = res.get('data', {}).get('returnValue', res.get('data'))
print('render ->', json.dumps(rv, ensure_ascii=False))

src = Path(REMOTE[:-4] if REMOTE.endswith('.png') else REMOTE)
if not src.is_dir():
    sys.exit('render directory missing: %s' % REMOTE)

copied = []
for i in range(1, 200):
    f = src / ('幻灯片%d.png' % i)
    if not f.exists():
        break
    dst = OUT / ('slide-%02d.png' % i)
    shutil.copyfile(f, dst)
    copied.append((i, dst.stat().st_size))
print('copied', len(copied), 'images to', OUT)
for i, size in copied:
    print('  slide %02d  %6d bytes' % (i, size))
shutil.rmtree(src, ignore_errors=True)
