"""Packaged legacy migration, hidden-state regeneration and UI verification."""
from pathlib import Path
import hashlib,json,shutil,stat,subprocess,sys,tempfile
APP=Path(__file__).resolve().parents[1]
release=Path(sys.argv[1]).resolve()
legacy=Path(sys.argv[2]).resolve()
work=Path(tempfile.mkdtemp(prefix='release111-',dir=APP/'qa'));folder=work/'20260926'
shutil.copytree(legacy,folder);out=folder/'000 Output';internal=out/'.PackingData'
names=['Processing_History.json','Processed_Files.txt','.packing.lock','Courier_Selections.json']
original={n:(out/n).read_bytes() for n in names if (out/n).is_file()}
assert 'Processing_History.json' in original and 'Processed_Files.txt' in original
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
documents={p.name:sha(p) for p in out.glob('*.pdf')}
def call(action,**kwargs):
    req=work/'request.json';res=work/'response.json'
    req.write_text(json.dumps({'action':action,'folder':str(folder),'catalog_path':str(work/'ProductNames.json'),**kwargs}),encoding='utf-8');res.unlink(missing_ok=True)
    subprocess.run([str(release/'App Files/PackingEngine.exe'),'--request',str(req),'--response',str(res)],check=True,timeout=240,creationflags=subprocess.CREATE_NO_WINDOW)
    data=json.loads(res.read_text(encoding='utf-8'));assert data['ok'],data;return data
scan=call('scan');assert scan['version']=='1.11.0'
for n,data in original.items():assert (internal/n).read_bytes()==data and not (out/n).exists(),n
assert internal.stat().st_file_attributes & stat.FILE_ATTRIBUTE_HIDDEN
assert documents=={p.name:sha(p) for p in out.glob('*.pdf')}
assert all(e['message']=='Process again' for e in scan['entries'] if e['platform']!='SHOPEE')
spx=[{'path':e['path'],'sha256':e['sha256'],'courier':e['courier'],'platform':e['platform']} for e in scan['entries'] if e['platform']=='SPX']
warning=call('generate',selected=spx);assert warning['confirmation_required']
assert call('generate',selected=spx,repeat_confirmation=warning['repeat_token'],repeat_action='skip')['outputs']==[]
for p in out.glob('SPX_Packing_Checklist_*.pdf'):p.unlink()
missing=call('scan');assert next(e for e in missing['entries'] if e['platform']=='SPX')['status']=='MISSING'
warning=call('generate',selected=spx)
made=call('generate',selected=spx,repeat_confirmation=warning['repeat_token'],repeat_action='skip')['outputs']
assert len(made)==1 and made[0]['filename'].endswith('_1.pdf') and made[0]['pages']==2
import pdfplumber
with pdfplumber.open(made[0]['path']) as pdf:
    text='\n'.join(page.extract_text() for page in pdf.pages)
    assert '7 orders' in text and '8 pieces' in text and 'TOTAL PIECES' in text
scan=call('scan');assert all(e['message']=='Process again' for e in scan['entries'] if e['platform']!='SHOPEE')
assert len(scan['outputs'])==3 and sum(o['order_count'] for o in scan['outputs'])==17 and sum(o['pieces'] for o in scan['outputs'])==30
assert set(p.name for p in out.iterdir() if p.is_file())=={o['filename'] for o in scan['outputs']}
assert not list(internal.glob('.pending-*')) and not (internal/'.packing-transaction.json').exists()
for n,digest in documents.items():
    if not n.startswith('SPX_'):assert sha(out/n)==digest
subprocess.run([str(release/'Packing List Manager.exe'),'--self-test',str(folder),str(work/'main.png'),str(work/'history.png'),str(work/'names.png'),'--check-couriers','--select-for-preview'],check=True,timeout=180,creationflags=subprocess.CREATE_NO_WINDOW)
ui=json.loads((release/'SelfTest.json').read_text(encoding='utf-8'));assert ui['ok'] and ui['selectionTotalsWork'] and ui['courierChangesWork'],ui
report={'ok':True,'version':'1.11.0','work':str(work),'migrated_byte_for_byte':list(original),'windows_hidden':True,'old_documents_preserved_during_migration':True,'no_skips_existing':True,'deleted_output_regenerated':made,'visible_files_only_pdfs':True,'orders':17,'pieces':30,'ui':ui}
(APP/'qa/release-1.11-verification.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report,indent=2))
