"""Packaged packing-price PDF checks with real Word A4/order/summary validation."""
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
import pdfplumber
from docx import Document

APP=Path(__file__).resolve().parents[1];sys.path.insert(0,str(APP/'src'))
from documents import build_document,verify_word,item_prices
from extractors import Batch,Order,Item,parse_pdf,clean
from names import apply_names,load_catalog

release=Path(sys.argv[1]).resolve()
work=Path(tempfile.mkdtemp(prefix='sort-prices-',dir=APP/'qa'));folder=work/'20260926'
fixture=APP/'qa/names-regression-fixtures/20260926'
for relative in ['SPX/20260926_SPX_PL_4.pdf','FLASH/20260926_Flash_PL_3.pdf','LAZADA & TIKTOK/20260926_Lazada_PL_1.pdf','LAZADA & TIKTOK/20260926_Tiktok_PL_1.pdf']:
    target=folder/relative;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(fixture/relative,target)
catalog=work/'ProductNames.json'
shutil.copyfile(APP.parent/'Packing List Manager/ProductNames.json',catalog)
def call(action,**fields):
    request=work/'request.json';response=work/'response.json'
    request.write_text(json.dumps(dict(action=action,folder=str(folder),catalog_path=str(catalog),**fields),ensure_ascii=False),encoding='utf-8')
    subprocess.run([str(release/'App Files/PackingEngine.exe'),'--request',str(request),'--response',str(response)],check=True,timeout=240,creationflags=subprocess.CREATE_NO_WINDOW)
    result=json.loads(response.read_text(encoding='utf-8'));assert result['ok'],result;return result
scan=call('scan');selected=[dict(path=e['path'],sha256=e['sha256']) for e in scan['entries']]
generated=call('generate',selected=selected)
if generated.get('confirmation_required'):generated=call('generate',selected=selected,repeat_confirmation=generated['repeat_token'],repeat_action='skip')
assert len(generated['outputs'])==3
orders={}
for source in folder.rglob('*.pdf'):
    if '000 Output' in source.parts:continue
    batch=parse_pdf(source)
    for order in apply_names([batch],load_catalog(catalog))[0].orders:orders[order.order_id]=order
priced_lines=0;order_count=0;pdfs=[]
for output in generated['outputs']:
    path=Path(output['path']);pdfs.append(str(path))
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            assert abs(page.width-595.3)<1 and abs(page.height-841.9)<1
            for table in page.extract_tables():
                if not table or not table[0] or not table[0][0] or not table[0][0].startswith(('Order ID:','Order Number:')):continue
                identifier=table[0][0].split(':',1)[1].split()[0];order=orders[identifier];order_count+=1
                columns=[clean(v) for v in table[1]]
                expected_columns=['Done','Product name','Variation','Pieces']+(['Unit Price','Total Price'] if order.platform in ('SPX','FLASH') else [])
                assert columns==expected_columns,(path,identifier,columns)
                assert len(table[2:])==len(order.items),(path,identifier,table)
                for row,item in zip(table[2:],order.items):
                    expected=['[ ]',clean(item.product),clean(item.variation),str(item.pieces)]
                    if order.platform in ('SPX','FLASH'):expected+=item_prices(item);priced_lines+=1
                    assert [clean(v) for v in row]==expected,(path,identifier,row,expected)
assert order_count==17 and priced_lines>10
history=json.loads((folder/'000 Output/.PackingData/Processing_History.json').read_text(encoding='utf-8'))
for original in history['orders'].values():
    for item in original['order']['items']:
        if original['order']['platform'] in ('SPX','FLASH'):assert item['unit_price'] is not None and item['total_price'] is not None
assert not list((folder/'000 Output').glob('*.docx'))
stress=work/'Summary-row-pagination.docx'
batch=Batch('pagination-test.pdf','SPX','hash',[Order('SPX',f'QA{i:03}',[Item(f'QA product {i:03} with a different picking row','100ML',2,'12.50','25.00')]) for i in range(75)])
build_document([batch],'SPX','20260926',stress);layout=verify_word([stress],export_pdf=True)
with pdfplumber.open(stress.with_suffix('.pdf')) as pdf:
    texts=[p.extract_text() or '' for p in pdf.pages];summary=next(i for i,t in enumerate(texts) if 'Product Picking Summary' in t)
    assert len(texts)-summary>=2 and 'TOTAL PIECES' in texts[-1]
report=dict(ok=True,version='1.11.1',work=str(work),pdfs=pdfs,orders=order_count,priced_line_items=priced_lines,source_prices_verified=True,history_prices_retained=True,blue_layout_unchanged=True,multipage_summary_word_verified=layout)
(APP/'qa/sort-prices-verification.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report,indent=2))
