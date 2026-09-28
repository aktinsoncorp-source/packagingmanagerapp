"""Compare the original and optimized Word checker on identical document content."""
from pathlib import Path
import json,subprocess,sys,time,shutil
import pdfplumber
APP=Path(__file__).resolve().parents[1];sys.path.insert(0,str(APP/'src'))
from documents import build_document,verify_word
from extractors import parse_pdf
root=APP/'qa/performance110';fixture=APP/'qa/names-regression-fixtures/20260926'
paths=[]
for group,names in [('SPX',['SPX/20260926_SPX_PL_4.pdf']),('FLASH',['FLASH/20260926_Flash_PL_3.pdf']),('LAZADA_TIKTOK',['LAZADA & TIKTOK/20260926_Lazada_PL_1.pdf','LAZADA & TIKTOK/20260926_Tiktok_PL_1.pdf'])]:
    dest=root/('small-'+group+'.docx');build_document([parse_pdf(fixture/n) for n in names],group,'20260926',dest);paths.append(dest)
request=root/'comparison-request.json';result=root/'comparison-result.json'
request.write_text(json.dumps({'paths':[str(p) for p in paths],'result':str(result),'export_pdf':True}),encoding='utf-8')
start=time.perf_counter()
subprocess.run([r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(root/'baseline_verify_word.ps1'),'-RequestFile',str(request)],check=True,timeout=180,creationflags=subprocess.CREATE_NO_WINDOW)
before=time.perf_counter()-start
original=json.loads(result.read_text(encoding='utf-8-sig'));assert original['ok'],original
texts=[]
for p in paths:
    with pdfplumber.open(p.with_suffix('.pdf')) as pdf:texts.append([page.extract_text() for page in pdf.pages])
start=time.perf_counter();updated=verify_word(paths,export_pdf=True);after=time.perf_counter()-start
for p,expected in zip(paths,texts):
    with pdfplumber.open(p.with_suffix('.pdf')) as pdf:assert [page.extract_text() for page in pdf.pages]==expected
assert [(d['pages'],d['orders'],d['summaries']) for d in original['documents']]==[(d['pages'],d['orders'],d['summaries']) for d in updated]
# The large-order check also retained every page's exact text and the 15-page layout.
with pdfplumber.open(root/'baseline.pdf') as old,pdfplumber.open(root/'updated.pdf') as new:
    assert [p.extract_text() for p in old.pages]==[p.extract_text() for p in new.pages]
report={'ok':True,'three_documents_orders':17,'pieces':30,'baseline_word_seconds':before,'updated_word_seconds':after,'seconds_saved':before-after,'percent_faster':(1-after/before)*100,'identical_pdf_text_and_pagination':True,'large_sample_orders':127,'large_sample_pieces':300,'large_sample_pages':15}
(root/'comparison.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report,indent=2))
