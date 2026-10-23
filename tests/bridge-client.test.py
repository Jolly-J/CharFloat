"""Fake HTTP responses only; never connects to a service or reads real credentials."""
import sys
sys.dont_write_bytecode = True
import base64
import contextlib
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
import urllib.error
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('bridge_client', 'skills/charfloat/scripts/bridge_client.py')
b = importlib.util.module_from_spec(spec)
spec.loader.exec_module(b)
PNG = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=')

class ClientTests(unittest.TestCase):
    def run_cli(self, args, response):
        calls = []
        def request(route, body=None, **kw):
            calls.append((route, body))
            return response
        out = io.StringIO()
        with patch.object(b, 'request', request), contextlib.redirect_stdout(out):
            code = b.main(args)
        return code, out.getvalue(), calls

    def test_legacy_success_preserves_base64_and_shape(self):
        original = {'success': True, 'data': {'success': True, 'imageBase64': base64.b64encode(PNG).decode(), 'extra': 7}}
        _, out, calls = self.run_cli(['wps_ppt_capture_slide_preview', '{"presentationName":"示例.pptx"}'], original)
        self.assertEqual(json.loads(out), original)
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][1]['arguments']['presentationName'], '示例.pptx')

    def test_utf8_file_output_and_explicit_session(self):
        with tempfile.TemporaryDirectory() as directory:
            args = Path(directory)/'参数.json'; out = Path(directory)/'结果.json'
            args.write_text('{"presentationName":"示例.pptx","slideIndex":2}', encoding='utf-8')
            _, stdout, calls = self.run_cli(['wps_ppt_get_slide_shapes','--args-file',str(args),'--output',str(out),'--session-id','task-test'], {'success': True, 'data': {'slideId': 8}})
            self.assertEqual(stdout, '')
            self.assertEqual(json.loads(out.read_text())['data']['slideId'], 8)
            self.assertEqual(calls[0][1]['arguments']['presentationName'], '示例.pptx')
            self.assertEqual(calls[0][1]['sessionId'], 'task-test')

    def test_default_sessions_independent(self):
        _, _, a = self.run_cli(['tool'], {})
        _, _, c = self.run_cli(['tool'], {})
        self.assertNotEqual(a[0][1]['sessionId'], c[0][1]['sessionId'])

    def test_queries_use_existing_endpoint_and_filter(self):
        tools = [{'name':'wps_ppt_read_presentation','description':'读取','inputSchema':{'type':'object'}},{'name':'other','description':'其他'}]
        _, out, calls = self.run_cli(['--schema','wps_ppt_read_presentation','--session-id','task-test'], tools)
        self.assertEqual(json.loads(out), tools[0]); self.assertEqual(calls, [('/api/v1/mcp-tools', None)])
        _, out, _ = self.run_cli(['--list-tools','wps_ppt_'], tools)
        self.assertEqual([x['name'] for x in json.loads(out)], ['wps_ppt_read_presentation'])

    def test_preview_explicit_png_output(self):
        with tempfile.TemporaryDirectory() as directory:
            file = Path(directory)/'slide.png'
            response={'success':True,'data':{'imageBase64':base64.b64encode(PNG).decode(),'success':True}}
            _, out, _ = self.run_cli(['wps_ppt_capture_slide_preview','--preview-output',str(file)], response)
            self.assertEqual(file.read_bytes(),PNG)
            data=json.loads(out)['data'];self.assertNotIn('imageBase64',data);self.assertEqual(data['previewBytes'],len(PNG))
            self.assertIn('imageBase64',response['data'])
            with self.assertRaises(SystemExit):
                self.run_cli(['wps_ppt_capture_slide_preview','--preview-output',str(file)],response)

    def test_invalid_preview_does_not_write(self):
        with tempfile.TemporaryDirectory() as directory:
            file=Path(directory)/'slide.png'
            with self.assertRaises(ValueError):
                b.prepare_result({'imageBase64':base64.b64encode(b'not-png').decode()},str(file))
            self.assertFalse(file.exists())

    def test_http_failure_fields_and_timeout_single_attempt(self):
        for source in [
            {'success':False,'error':'宿主失败','errorDetails':{'kind':'failed','channel':'wps-addon','executed':'unknown'},'partial':{'completedShapeIds':[7]}},
            {'success':False,'error':'拒绝','errorDetails':{'kind':'rejected','executed':'no'}}
        ]:
            with tempfile.TemporaryDirectory() as directory:
                home=Path(directory);(home/'token').write_text('fake-test-only')
                error=urllib.error.HTTPError('http://127.0.0.1',422,'failed',{},io.BytesIO(json.dumps(source).encode()))
                with patch.object(b,'home_dir',lambda:home), patch.object(b.urllib.request,'urlopen',side_effect=error) as open_mock:
                    with self.assertRaises(b.BridgeFailure) as captured:b.request('/api/v1/tool/call',{'name':'tool'})
                    self.assertEqual(open_mock.call_count,1)
                    self.assertEqual(captured.exception.detail['executed'],source['errorDetails']['executed'])
                    self.assertEqual(captured.exception.detail.get('partial'),source.get('partial'))
                    if source['errorDetails']['executed']=='unknown':self.assertIn('读回',captured.exception.detail['recovery'])
        with tempfile.TemporaryDirectory() as directory:
            home=Path(directory);(home/'token').write_text('fake-test-only')
            with patch.object(b,'home_dir',lambda:home),patch.object(b.urllib.request,'urlopen',side_effect=TimeoutError('timeout')) as open_mock:
                with self.assertRaises(b.BridgeFailure) as captured:b.request('/api/v1/tool/call',{'name':'tool'})
                self.assertEqual(open_mock.call_count,1);self.assertEqual(captured.exception.detail['executed'],'unknown')

    def test_nested_failure_and_error_result_file_not_replayed(self):
        with tempfile.TemporaryDirectory() as directory:
            output=Path(directory)/'error.json'
            code, stdout, calls=self.run_cli(['tool','--output',str(output)],{'success':True,'data':{'success':False,'error':'partial','partial':{'completedShapeIds':[4]}}})
            self.assertEqual(code,1);self.assertEqual(len(calls),1);self.assertEqual(stdout,'')
            result=json.loads(output.read_text());self.assertEqual(result['executed'],'unknown');self.assertEqual(result['partial']['completedShapeIds'],[4])

    def test_output_race_after_host_success_never_replays(self):
        for preview in [False, True]:
            with tempfile.TemporaryDirectory() as directory:
                target=Path(directory)/('preview.png' if preview else 'result.json')
                calls=[]
                def request(route, body=None, **kw):
                    calls.append(route)
                    target.write_bytes(b'created-by-other-writer')
                    return {'success':True,'data':{'slideIndex':1,'imageBase64':base64.b64encode(PNG).decode()}}
                stderr=io.StringIO()
                args=['wps_ppt_capture_slide_preview','--preview-output' if preview else '--output',str(target)]
                with patch.object(b,'request',request),contextlib.redirect_stderr(stderr):
                    code=b.main(args)
                detail=json.loads(stderr.getvalue())
                self.assertEqual(code,1);self.assertEqual(len(calls),1)
                self.assertEqual(detail['phase'],'local_output');self.assertTrue(detail['hostCallReturnedSuccess']);self.assertTrue(detail['doNotReplay'])
                self.assertNotIn('imageBase64',stderr.getvalue());self.assertEqual(target.read_bytes(),b'created-by-other-writer')

    def test_output_permission_failure_after_success_never_replays(self):
        with tempfile.TemporaryDirectory() as directory:
            target=Path(directory)/'result.json'
            calls=[]
            def request(route, body=None, **kw):
                calls.append(route);return {'success':True,'data':{'slideId':101}}
            stderr=io.StringIO()
            with patch.object(b,'request',request),patch.object(b.Path,'open',side_effect=PermissionError('simulated local write denied')),contextlib.redirect_stderr(stderr):
                code=b.main(['tool','--output',str(target)])
            detail=json.loads(stderr.getvalue());self.assertEqual(code,1);self.assertEqual(len(calls),1)
            self.assertEqual(detail['resultSummary'],{'slideId':101});self.assertEqual(detail['outputErrorType'],'PermissionError');self.assertTrue(detail['doNotReplay'])

    def test_failed_host_then_output_failure_preserves_state_and_partial_on_stderr(self):
        for race in [False, True]:
            with tempfile.TemporaryDirectory() as directory:
                target=Path(directory)/'error.json'
                failure=b.failure_detail({'success':False,'error':'synthetic timeout','errorDetails':{'kind':'unknown','executed':'unknown'},'partial':{'completedShapeIds':[7],'private':'must-not-leak'},'diagnosticType':'ppt_layout','before':{'slideWidth':900},'after':{'slideWidth':1000,'private':'must-not-leak'},'applied':{'slideWidth':1000},'warningDetails':[{'field':'slideHeight','code':'write_failed','reason':'host_rejected','private':'must-not-leak'}]})
                calls=[]
                def request(route, body=None, **kw):
                    calls.append(route)
                    if race:target.write_text('other-writer')
                    raise b.BridgeFailure(failure)
                stderr=io.StringIO()
                if race:
                    with patch.object(b,'request',request),contextlib.redirect_stderr(stderr):code=b.main(['tool','--output',str(target)])
                else:
                    with patch.object(b,'request',request),patch.object(b.Path,'open',side_effect=PermissionError('synthetic denied')),contextlib.redirect_stderr(stderr):code=b.main(['tool','--output',str(target)])
                result=json.loads(stderr.getvalue())
                self.assertEqual(code,1);self.assertEqual(len(calls),1);self.assertFalse(result['hostCallReturnedSuccess'])
                self.assertEqual(result['executed'],'unknown');self.assertEqual(result['errorDetails']['executed'],'unknown')
                self.assertEqual(result['partial'],{'completedShapeIds':[7]});self.assertTrue(result['doNotReplay']);self.assertEqual(result['hostError'],'synthetic timeout')
                self.assertNotIn('must-not-leak',stderr.getvalue())
                self.assertEqual(result['after'],{'slideWidth':1000});self.assertEqual(result['applied'],{'slideWidth':1000})
                self.assertEqual(result['warningDetails'],[{'field':'slideHeight','code':'write_failed','reason':'host_rejected'}])

if __name__=='__main__':unittest.main()
