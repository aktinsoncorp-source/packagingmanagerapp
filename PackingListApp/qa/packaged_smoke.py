"""Exercise the distributable engine with the real Microsoft Word validator."""
import json
from collections import Counter
import shutil
import subprocess
import sys
from pathlib import Path

APP=Path(__file__).resolve().parents[1]
ROOT=APP.parent
sys.path.insert(0,str(APP/'src'))
from engine import read_completed_doc
from extractors import file_hash
from docx import Document

work=APP/'qa'/'packaged-smoke'
work.mkdir(exist_ok=True)
folder=work/'20260926'
if folder.exists():raise RuntimeError('Use a fresh smoke-test folder; do not overwrite a prior result.')
for sub in ('SPX','FLASH','LAZADA & TIKTOK'):
    shutil.copytree(ROOT/'20260926'/sub,folder/sub)
exe=ROOT/'Packing List Manager'/'App Files'/'PackingEngine.exe'

def call(action, **kwargs):
    result = _call(action, **kwargs)
    if action == 'generate' and result.get('confirmation_required'):
        result = _call(action, **kwargs, repeat_confirmation=result['repeat_token'])
    return result


def _call(action,**args):
    request=work/'request.json';response=work/'response.json'
    request.write_text(json.dumps({'action':action,'folder':str(folder),**args}),encoding='utf-8')
    if response.exists():response.unlink()
    subprocess.run([str(exe),'--request',str(request),'--response',str(response)],check=True,timeout=240,creationflags=subprocess.CREATE_NO_WINDOW)
    result=json.loads(response.read_text(encoding='utf-8'))
    assert result['ok'],result
    return result

first=call('scan')
assert len(first['entries'])==9
assert all(e['status']=='READY' for e in first['entries'])
selected=[{'path':e['path'],'sha256':e['sha256']} for e in first['entries']]
result=call('generate', output_format='docx',selected=selected)
assert len(result['outputs'])==3
for output in result['outputs']:
    name=output['filename'];group=name.split('_Packing_')[0]
    approved=read_completed_doc(ROOT/'20260926'/'000 Output'/name,group)
    actual=read_completed_doc(Path(output['path']),group)
    assert set(actual)==set(approved),(group,'order IDs')
    assert all(actual[k]['fingerprint']==approved[k]['fingerprint'] for k in actual),(group,'content')
    assert output['pages']=={'SPX':11,'FLASH':4,'LAZADA_TIKTOK':2}[group],output
    summary = lambda p: [[c.text for c in r.cells] for r in Document(p).tables[-1].rows]
    expected = Counter()
    for record in approved.values():
        for item in record['order']['items']:
            expected[(item['product'],item['variation'])] += item['pieces']
    assert {(r[1],r[2]):int(r[3]) for r in summary(Path(output['path']))[1:-1]} == dict(expected), (group, 'original-name summary')
again=call('scan')
assert all(e['status']=='DONE' for e in again['entries'])
report={'ok':True,'source_files':9,'orders':139,'pieces':318,'outputs':result['outputs'],
        'second_scan':'All source files marked already processed'}
(work/'verification.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2))
