"""Restore isolated regression data from retained QA copies and a supplied source root."""
from pathlib import Path
import argparse
import shutil
import sys

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / 'src'))
from extractors import parse_pdf


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--source-root', required=True, type=Path)
    parser.add_argument('--destination', required=True, type=Path)
    args = parser.parse_args()
    args.destination.mkdir(parents=True, exist_ok=False)
    source = APP / 'qa/packaged-smoke/20260926'
    target = args.destination / '20260926'
    paths = [p for p in source.rglob('*.pdf') if '000 Output' not in p.relative_to(source).parts]
    for path in paths:
        dest = target / path.relative_to(source)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, dest)
    out = target / '000 Output'
    out.mkdir()
    log = ['PACKING LIST PROCESSING LOG', '']
    for group in ('SPX', 'FLASH', 'LAZADA_TIKTOK'):
        filename = f'{group}_Packing_Checklist_2026-09-26_1.docx'
        shutil.copyfile(APP / 'qa/original-names-layout' / f'{group}.docx', out / filename)
        log += [filename, 'Status: Completed']
        for path in paths:
            batch = parse_pdf(path)
            if batch.group == group:
                log.append(f'  - {path.relative_to(source).as_posix()} | {len(batch.orders)} orders')
        log.append('')
    (out / 'Processed_Files.txt').write_text('\n'.join(log), encoding='utf-8')
    counts = {'20260926': len(paths)}
    for date in ('20260924', '20260925'):
        counts[date] = 0
        for path in sorted((args.source_root / date).rglob('*.pdf')):
            if '000 Output' in path.parts or 'WAYBILL' in path.name.upper():
                continue
            batch = parse_pdf(path)
            dest = args.destination / date / batch.platform / path.name
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(path, dest)
            counts[date] += 1
    print(counts)
