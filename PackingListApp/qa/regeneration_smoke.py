"""Verify the installed app can regenerate and recover deleted Word outputs."""
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

APP = Path(__file__).resolve().parents[1]
ROOT = APP.parent
sys.path.insert(0, str(APP / 'src'))
from engine import read_completed_doc
from extractors import parse_pdf

work = Path(tempfile.mkdtemp(prefix='regeneration-smoke-', dir=APP / 'qa'))
folder = work / '20260926'
(folder / 'SPX').mkdir(parents=True)
source = ROOT / '20260926/SPX/20260926_SPX_PL_4.pdf'
shutil.copyfile(source, folder / 'SPX' / source.name)
batch = parse_pdf(source)
exe = ROOT / 'Packing List Manager/App Files/PackingEngine.exe'


def call(action, **kwargs):
    request, response = work / 'request.json', work / 'response.json'
    request.write_text(json.dumps({'action': action, 'folder': str(folder), **kwargs}), encoding='utf-8')
    response.unlink(missing_ok=True)
    subprocess.run([str(exe), '--request', str(request), '--response', str(response)], check=True,
                   timeout=240, creationflags=subprocess.CREATE_NO_WINDOW)
    result = json.loads(response.read_text(encoding='utf-8'))
    assert result['ok'], result
    return result


def generate():
    e = call('scan')['entries'][0]
    result = call('generate', selected=[{'path': e['path'], 'sha256': e['sha256']}])
    output = result['outputs'][0]
    records = read_completed_doc(Path(output['path']), 'SPX')
    assert {o.key: o.fingerprint() for o in batch.orders} == {k: v['fingerprint'] for k, v in records.items()}
    return output


first = generate()
assert call('scan')['entries'][0]['status'] == 'DONE'
second = generate()
assert second['filename'].endswith('_2.docx')
assert second['mode'] == 'regenerated'
Path(first['path']).unlink()
assert call('scan')['entries'][0]['status'] == 'DONE'
Path(second['path']).unlink()
missing = call('scan')
assert missing['entries'][0]['status'] == 'MISSING'
assert missing['entries'][0]['missing_orders'] == len(batch.orders)
# Retain this response for the UI missing-output scenario.
(work / 'missing-scan.json').write_text(json.dumps(missing, indent=2), encoding='utf-8')
third = generate()
assert third['filename'].endswith('_1.docx')
assert call('scan')['entries'][0]['status'] == 'DONE'
history = json.loads((folder / '000 Output/Processing_History.json').read_text(encoding='utf-8'))
for output in history['outputs'].values():
    assert len(output['sources']) == 1
    assert output['sources'][0]['path'] == 'SPX/' + source.name
    assert len(output['sources'][0]['order_keys']) == len(batch.orders)
report = {'ok': True, 'work': str(work), 'orders': len(batch.orders), 'pieces': batch.pieces,
          'manual_regeneration': 'passed', 'surviving_copy': 'passed', 'deleted_output_recovery': 'passed',
          'next_suffix': third['filename'], 'source_manifests': 'verified', 'pages': third['pages']}
(APP / 'qa/regeneration-packaged-verification.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(json.dumps(report, indent=2))
