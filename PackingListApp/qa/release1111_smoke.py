"""Packaged PDF-only generation with No skipping existing orders, alternating rows and employee controls verification."""
from pathlib import Path
import json, os, shutil, subprocess, sys, tempfile
APP=Path(__file__).resolve().parents[1];sys.path.insert(0,str(APP/'src'))
import pdfplumber
release=Path(sys.argv[1]).resolve();fixture=APP/'qa/names-regression-fixtures/20260926'
work=Path(tempfile.mkdtemp(prefix='release1111-',dir=APP/'qa'));folder=work/'20260926'
for relative in ['SPX/20260926_SPX_PL_4.pdf','FLASH/20260926_Flash_PL_3.pdf','LAZADA & TIKTOK/20260926_Lazada_PL_1.pdf','LAZADA & TIKTOK/20260926_Tiktok_PL_1.pdf']:
    target=folder/relative;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(fixture/relative,target)
shutil.copyfile(fixture/'SPX/20260926_SPX_PL_4.pdf',folder/'packing.pdf')

def call(action,**kwargs):
    request=work/'request.json';response=work/'response.json'
    request.write_text(json.dumps({'action':action,'folder':str(folder),'catalog_path':str(work/'ProductNames.json'),**kwargs}),encoding='utf-8')
    response.unlink(missing_ok=True)
    subprocess.run([str(release/'App Files/PackingEngine.exe'),'--request',str(request),'--response',str(response)],check=True,timeout=240,creationflags=subprocess.CREATE_NO_WINDOW)
    data=json.loads(response.read_text(encoding='utf-8'));
    if not data['ok']:raise RuntimeError(data.get('error',data))
    return data
entries=call('scan')['entries'];selection=[{'path':e['path'],'sha256':e['sha256'],'platform':e['platform'],'courier':e['courier']} for e in entries if e['platform']!='SHOPEE']
assert len(selection)==4
spx=[s for s in selection if s['platform']=='SPX']
first=call('generate',selected=spx)
assert first['confirmation_required'] and first['repeat_count']==7 and first['previously_generated']==0
assert not list((folder/'000 Output').glob('*_Packing_Checklist_*'))
# No still generates an unprocessed order shared by multiple source PDFs once.
initial=call('generate',selected=spx,repeat_confirmation=first['repeat_token'],repeat_action='skip')
assert initial['skipped_repeat_orders']==0 and len(initial['outputs'])==1
# No skips the processed SPX selection but generates every other platform.
warning=call('generate',selected=selection);assert warning['previously_generated']==7
mixed=call('generate',selected=selection,repeat_confirmation=warning['repeat_token'],repeat_action='skip')
assert mixed['skipped_repeat_orders']==7 and len(mixed['outputs'])==2
pdfs=initial['outputs']+mixed['outputs']
assert len(pdfs)==3 and all(o['filename'].endswith('_1.pdf') for o in pdfs)
assert not list((folder/'000 Output').glob('*.docx'))
boxes=0;gray_rectangles=0
for output in pdfs:
    with pdfplumber.open(output['path']) as pdf:
        ids=[]
        for page in pdf.pages:
            assert abs(page.width-595.3)<1 and abs(page.height-841.9)<1
            # Word emits neutral fills using PDF DeviceGray, but RGB is valid too.
            for rect in page.rects:
                color=rect.get('non_stroking_color')
                components=color if isinstance(color,(list,tuple)) else [color]
                if components and all(isinstance(c,(int,float)) and abs(c-242/255)<.005 for c in components):
                    gray_rectangles+=1
            for char in page.chars:
                if char['text']=='[' and char['x0']<55:
                    closing=[c for c in page.chars if c['text']==']' and c['x0']<62 and abs(c['top']-char['top'])<.8 and c['x0']>char['x1']]
                    assert closing and closing[0]['x0']-char['x1']>10,(output,char,closing)
                    boxes+=1
assert boxes>25
scan=call('scan');assert all(e['status']=='DONE' and e['message']=='Process again' for e in scan['entries'] if e['platform']!='SHOPEE')
assert sum(o['order_count'] for o in scan['outputs'])==17 and sum(o['pieces'] for o in scan['outputs'])==30
assert gray_rectangles>0
warning=call('generate',selected=selection)
all_done=call('generate',selected=selection,repeat_confirmation=warning['repeat_token'],repeat_action='skip')
assert all_done['outputs']==[] and all_done['skipped_repeat_orders']==17
assert len(list((folder/'000 Output').glob('*.pdf')))==3
try:
    call('generate',selected=selection,output_format='docx')
    raise AssertionError('Word export was unexpectedly accepted')
except RuntimeError as error:
    assert 'PDF only' in str(error)
spx=[s for s in selection if s['platform']=='SPX']
repeat=call('generate',selected=spx)
assert repeat['confirmation_required'] and repeat['previously_generated']==7
reprint=call('generate',selected=spx,repeat_confirmation=repeat['repeat_token'])['outputs'][0]
assert reprint['filename'].endswith('_2.pdf')
assert not list((folder/'000 Output').glob('*.docx'))
# Every deleted output loses active completion and repeat coverage.
for output in call('scan')['outputs']:
    (folder/'000 Output'/output['filename']).unlink()
reset=call('scan')
assert reset['version']=='1.11.1' and not reset['outputs']
assert all(e['status']=='READY' for e in reset['entries'] if e['platform']!='SHOPEE')
assert all('missing' not in e['message'].lower() for e in reset['entries'])
warning=call('generate',selected=selection)
assert warning['previously_generated']==0 and warning['overlapping_orders']==7
restored=call('generate',selected=selection,repeat_confirmation=warning['repeat_token'],repeat_action='skip')
assert len(restored['outputs'])==3 and all(o['mode']=='new' for o in restored['outputs'])
pdfs=restored['outputs']
for output in pdfs:
    with pdfplumber.open(output['path']) as pdf:
        text=pdf.pages[0].extract_text()
        assert 'Total orders: '+str(next(o['order_count'] for o in call('scan')['outputs'] if o['filename']==output['filename'])) in text, (output,text)
        for line in text.splitlines():
            if line.startswith(('Batch ', 'LAZADA   ', 'TIKTOK   ')):
                assert ' orders ' not in line and ' order ' not in line, line
subprocess.run([str(release/'Packing List Manager.exe'),'--self-test',str(folder),str(work/'main.png'),str(work/'history.png'),str(work/'names-locked.png'),'--check-couriers','--select-for-preview'],check=True,timeout=180,creationflags=subprocess.CREATE_NO_WINDOW)
ui=json.loads((release/'SelfTest.json').read_text(encoding='utf-8'));assert ui['ok'] and ui['selectionTotalsWork'],ui
report={'ok':True,'version':'1.11.1','work':str(work),'pdfs':pdfs,'reprint':reprint,'gray_rectangles':gray_rectangles,'word_export_rejected':True,'orders':17,'pieces':30,'wide_boxes_checked':boxes,'ui':ui,'repeat_warning_before_first_export':True,'regeneration_warning':True,'no_generated_unprocessed_overlap':True,'no_skipped_existing_and_generated_other_platforms':True,'no_all_processed_creates_no_empty_file':True}
(APP/'qa/release-1.11.1-verification.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report,indent=2))
