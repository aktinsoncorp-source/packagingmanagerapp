"""Installed engine names + document generation using a disposable PDF copy."""
from pathlib import Path
import json
import shutil
import subprocess
import sys
import tempfile

APP = Path(__file__).resolve().parents[1]
ROOT = APP.parent
sys.path.insert(0, str(APP / 'src'))
from docx import Document
from documents import verify_word

work = Path(tempfile.mkdtemp(prefix='names-packaged-', dir=APP / 'qa'))
folder = work / '20260926'
(folder / 'SPX').mkdir(parents=True)
shutil.copyfile(APP / 'qa/packaged-smoke/20260926/SPX/20260926_SPX_PL_4.pdf', folder / 'SPX/test.pdf')
catalog = work / 'ProductNames.json'
exe = ROOT / 'Packing List Manager/App Files/PackingEngine.exe'


def call(action, **kwargs):
    req, resp = work / 'request.json', work / 'response.json'
    req.write_text(json.dumps({'action': action, 'folder': str(folder), 'catalog_path': str(catalog), **kwargs}), encoding='utf-8')
    resp.unlink(missing_ok=True)
    subprocess.run([str(exe), '--request', str(req), '--response', str(resp)], check=True, timeout=240, creationflags=subprocess.CREATE_NO_WINDOW)
    result = json.loads(resp.read_text(encoding='utf-8'))
    assert result['ok'], result
    return result


scan = call('scan')
assert scan['version'] == '1.5.0'
data = call('names_list')
assert data['rows']
for row in data['rows']:
    row['alternative_product'] = 'QA Standard Inventory Item'
    row['alternative_variation'] = 'Standard / 1 pack'
call('names_save', rows=data['rows'], revision=data['revision'])
result = call('generate', selected=[{'path': e['path'], 'sha256': e['sha256']} for e in scan['entries']])
output = Path(result['outputs'][0]['path'])
doc = Document(output)
assert len(doc.tables) == 8
assert len(doc.tables[-1].rows) == 3
assert doc.tables[-1].cell(1, 1).text == 'QA Standard Inventory Item'
assert doc.tables[-1].cell(1, 3).text == '8'
assert call('scan')['entries'][0]['status'] == 'DONE'
layout = verify_word([output], export_pdf=True)
report = {'ok': True, 'version': '1.5.0', 'work': str(work), 'output': str(output), 'layout': layout,
          'orders': 7, 'pieces': 8, 'summary_rows': 1, 'original_fingerprints_preserved': True}
(APP / 'qa/names-packaged-verification.json').write_text(json.dumps(report, indent=2))
print(json.dumps(report, indent=2))
