"""Exercise the installed executable's requested numbering and contents view."""
import copy
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import uuid

APP = Path(__file__).resolve().parents[1]
ROOT = APP.parent
sys.path.insert(0, str(APP / 'src'))
from engine import read_completed_doc
from extractors import file_hash

work = Path(tempfile.mkdtemp(prefix='completed-smoke-', dir=APP / 'qa'))
folder = work / '20260926'
out = folder / '000 Output'
out.mkdir(parents=True)
(folder / 'SPX').mkdir()
for filename in ('20260926_SPX_PL_4.pdf', '20260926_SPX_PL_1.pdf'):
    shutil.copyfile(ROOT / '20260926/SPX' / filename, folder / 'SPX' / filename)
first_source = 'SPX/20260926_SPX_PL_4.pdf'
other_source = 'SPX/20260926_SPX_PL_1.pdf'
seed = Path(json.loads((APP / 'qa/original-names-packaged-verification.json').read_text())['output']['path'])
records = read_completed_doc(seed, 'SPX')
names = [f'SPX_Packing_Checklist_2026-09-26_{i}.docx' for i in (1, 2, 3)]
history = {'schema': 1, 'date': '20260926', 'sources': {}, 'orders': {}, 'outputs': {}, 'legacy': False}
for name in names:
    shutil.copyfile(seed, out / name)
    history['outputs'][name] = {'group': 'SPX', 'orders': list(records), 'mode': 'regenerated', 'created': 'QA fixture',
        'status': 'active', 'record_id': uuid.uuid4().hex,
        'order_fingerprints': {k: v['fingerprint'] for k, v in records.items()},
        'order_snapshots': {k: v['order'] for k, v in records.items()},
        'sources': [{'path': first_source, 'sha256': file_hash(folder / first_source), 'order_keys': list(records)}]}
history['orders'] = {k: {**copy.deepcopy(v), 'output': names[-1], 'source': first_source} for k, v in records.items()}
history['sources'][first_source] = {'sha256': file_hash(folder / first_source), 'platform': 'SPX', 'order_keys': list(records), 'outputs': names}
(out / 'Processing_History.json').write_text(json.dumps(history), encoding='utf-8')
(out / 'Processed_Files.txt').write_text('COMPLETED BY PACKING LIST MANAGER 1.4.0\nDisposable numbering test fixture.\n')
exe = ROOT / 'Packing List Manager/App Files/PackingEngine.exe'


def call(action, **kwargs):
    req, resp = work / 'request.json', work / 'response.json'
    req.write_text(json.dumps({'action': action, 'folder': str(folder), **kwargs}))
    resp.unlink(missing_ok=True)
    subprocess.run([str(exe), '--request', str(req), '--response', str(resp)], check=True, timeout=240,
                   creationflags=subprocess.CREATE_NO_WINDOW)
    result = json.loads(resp.read_text())
    assert result['ok'], result
    return result


def generate(source):
    scan = call('scan')
    entry = next(e for e in scan['entries'] if e['path'] == source)
    return call('generate', selected=[{'path': source, 'sha256': entry['sha256']}])['outputs'][0]


assert call('scan')['version'] == '1.4.0'
original_hash = file_hash(out / names[0])
(out / names[2]).unlink()
third = generate(other_source)
assert third['filename'] == names[2]
shown = next(o for o in call('scan')['outputs'] if o['filename'] == third['filename'])
assert [s['path'] for s in shown['included_sources']] == [other_source]
assert shown['order_count'] == 8
(out / names[1]).unlink()
fourth = generate(first_source)
assert fourth['filename'].endswith('_4.docx')
shown = call('scan')['outputs']
assert [o['sequence'] for o in shown] == [1, 3, 4]
assert file_hash(out / names[0]) == original_hash
for name in (names[0], third['filename'], fourth['filename']):
    (out / name).unlink()
assert call('scan')['outputs'] == []
restart = generate(first_source)
assert restart['filename'].endswith('_1.docx')
final = call('scan')
assert {e['path']: e['status'] for e in final['entries']}[other_source] == 'MISSING'
assert len(final['outputs']) == 1
report = {'ok': True, 'version': final['version'], 'work': str(work),
          'deleted_highest': third['filename'], 'deleted_middle': fourth['filename'], 'all_deleted': restart['filename'],
          'missing_entries_hidden': True, 'reused_filename_sources_correct': True, 'existing_file_preserved': True}
(APP / 'qa/completed-documents-packaged-verification.json').write_text(json.dumps(report, indent=2))
print(json.dumps(report, indent=2))
