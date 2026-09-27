"""Seed the installed catalog from compatible PDFs in existing dated folders."""
import argparse
import json
from pathlib import Path
import re
import sys

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / 'src'))
from extractors import parse_pdf
from names import encounter


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--source-root', type=Path, required=True)
    parser.add_argument('--catalog', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    batches, errors = [], []
    for folder in sorted(args.source_root.iterdir()):
        if not folder.is_dir() or not re.fullmatch(r'\d{8}', folder.name):
            continue
        for path in sorted(folder.rglob('*')):
            if path.is_file() and path.suffix.lower() == '.pdf' and '000 Output' not in path.relative_to(folder).parts:
                try:
                    batches.append(parse_pdf(path))
                except Exception as exc:
                    errors.append({'path': str(path), 'error': str(exc)})
    catalog = encounter(args.catalog, batches)
    report = {'compatible_pdfs': len(batches), 'unique_items': len(catalog['rows']), 'errors': errors}
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))
