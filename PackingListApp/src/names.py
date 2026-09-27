"""Persistent, exact product/variation definitions and Excel round trips."""
from contextlib import contextmanager
from dataclasses import replace
from pathlib import Path
import copy
import json
import os
import shutil
import uuid

from extractors import Item, PackingError

FIELDS = ('original_product', 'original_variation', 'alternative_product', 'alternative_variation')
HEADERS = ('Original product name', 'Original variation', 'Alternative product name', 'Alternative variation', 'Entry ID')
ALLOWED_SYMBOLS = " -()/%.+,&'"


def item_key(product, variation):
    return product, variation or '—'


def validate_alternative(value, label):
    if not isinstance(value, str):
        raise PackingError(f'{label} must be text.')
    if not value:
        return value
    if not value[0].isalnum():
        raise PackingError(f'{label} must start with a letter or number, without a leading space or symbol.')
    if any(not (c.isalnum() or c in ALLOWED_SYMBOLS) for c in value):
        raise PackingError(f'{label} may contain letters, numbers, spaces, and these symbols: - ( ) / % . + , & apostrophe.')
    return value.rstrip()


def validate_rows(rows):
    if not isinstance(rows, list):
        raise PackingError('The names list is invalid.')
    result, keys, ids = [], set(), set()
    for number, row in enumerate(rows, 1):
        if not isinstance(row, dict):
            raise PackingError(f'Row {number} is invalid.')
        entry = {field: row.get(field, '') for field in FIELDS}
        for field, value in entry.items():
            if not isinstance(value, str) or any(ord(c) < 32 for c in value):
                raise PackingError(f'Row {number}: {field.replace("_", " ")} must be a single line of text.')
        if not entry['original_product'].strip():
            raise PackingError(f'Row {number}: Original product name is required.')
        entry['original_variation'] = entry['original_variation'] or '—'
        for field in FIELDS[2:]:
            entry[field] = validate_alternative(entry[field], f'Row {number}: {field.replace("_", " ")}')
        key = item_key(entry['original_product'], entry['original_variation'])
        if key in keys:
            raise PackingError(f'Row {number}: this original product and variation already appear in another row.')
        entry['id'] = row.get('id') or uuid.uuid4().hex
        if not isinstance(entry['id'], str) or entry['id'] in ids:
            raise PackingError(f'Row {number}: duplicate or invalid entry identifier. Start with a fresh export.')
        keys.add(key)
        ids.add(entry['id'])
        result.append(entry)
    return sorted(result, key=lambda r: (r['original_product'].casefold(), r['original_variation'].casefold()))


def _load(path):
    if not path.exists():
        return {'schema': 1, 'revision': 0, 'rows': []}
    try:
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        if data['schema'] != 1 or not isinstance(data['revision'], int):
            raise ValueError('unrecognized format')
        data['rows'] = validate_rows(data['rows'])
        return data
    except Exception as exc:
        raise PackingError(f'Cannot read saved product names: {exc}. Keep {path.name} and restore a valid backup.') from exc


@contextmanager
def catalog_lock(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.with_suffix('.lock').open('a+b') as lock:
        import msvcrt
        lock.seek(0)
        if not lock.read(1):
            lock.write(b'0')
            lock.flush()
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            raise PackingError('Product names are being saved in another window. Try again shortly.')
        try:
            yield
        finally:
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)


def _write(path, data):
    temp = path.with_name(path.name + '.tmp-' + uuid.uuid4().hex)
    try:
        with temp.open('w', encoding='utf-8', newline='\n') as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        if path.exists():
            backup = path.with_suffix('.json.bak')
            backup_temp = temp.with_suffix('.bak')
            shutil.copyfile(path, backup_temp)
            os.replace(backup_temp, backup)
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def load_catalog(path):
    with catalog_lock(path):
        return _load(path)


def save_catalog(path, rows, revision):
    validated = validate_rows(rows)
    with catalog_lock(path):
        current = _load(path)
        if current['revision'] != revision:
            raise PackingError('The names list changed in another operation. Close Define Name, reopen it, and repeat your edits.')
        if current['rows'] != validated:
            current = {'schema': 1, 'revision': revision + 1, 'rows': validated}
            _write(path, current)
        return current


def encounter(path, batches):
    """Remember pairs on an actual PDF scan; list/export never resurrect deletes."""
    with catalog_lock(path):
        current = _load(path)
        known = {item_key(r['original_product'], r['original_variation']) for r in current['rows']}
        rows = list(current['rows'])
        for batch in batches:
            for order in batch.orders:
                for item in order.items:
                    key = item_key(item.product, item.variation)
                    if key not in known:
                        rows.append(dict(zip(FIELDS, [key[0], key[1], '', ''])))
                        known.add(key)
        if len(rows) != len(current['rows']):
            current = {'schema': 1, 'revision': current['revision'] + 1, 'rows': validate_rows(rows)}
            _write(path, current)
        return current


def apply_names(batches, catalog):
    mapping = {item_key(r['original_product'], r['original_variation']): r for r in catalog['rows']}
    def display(item):
        row = mapping.get(item_key(item.product, item.variation), {})
        return Item(row.get('alternative_product') or item.product,
                    row.get('alternative_variation') or item.variation or '—', item.pieces)
    return [replace(b, orders=[replace(o, items=[display(i) for i in o.items]) for o in b.orders]) for b in batches]


def export_excel(path, catalog):
    # Runtime Excel support is bundled with the offline employee application.
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.comments import Comment
    if path.suffix.lower() != '.xlsx':
        raise PackingError('Export names to an .xlsx Excel file.')
    wb = Workbook()
    ws = wb.active
    ws.title = 'Names'
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='19579D')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    for cell in (ws['C1'], ws['D1']):
        cell.fill = PatternFill('solid', fgColor='F8CBAD')
        cell.font = Font(name='Calibri', bold=True, color='202020')
    ws.row_dimensions[1].height = 32
    for number, row in enumerate(catalog['rows'], 2):
        for column, value in enumerate([row[f] for f in FIELDS] + [row['id']], 1):
            cell = ws.cell(number, column, value)
            cell.data_type = 's'  # Original listing text must never become a formula.
            cell.number_format = '@'
            cell.font = Font(name='Calibri', size=11, bold=column in (3, 4), color='19579D' if column in (3, 4) else '202020')
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            cell.fill = PatternFill('solid', fgColor='FFF7E8' if column in (3, 4) else ('F2F5F9' if number % 2 == 0 else 'FFFFFF'))
        ws.row_dimensions[number].height = 60
    for column, width in zip('ABCDE', [72, 34, 48, 34, 36]):
        ws.column_dimensions[column].width = width
    ws.column_dimensions['E'].hidden = True
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:E{max(1, ws.max_row)}'
    ws['A1'].comment = Comment('You may edit all four name columns or add rows. Hidden entry IDs let edited original names update the same entry. Missing rows are retained on import; delete entries in Define Name.', 'Packing List Manager')
    for cell in ('C1', 'D1'):
        ws[cell].comment = Comment('Blank uses the original field. Start with a letter or number. Allowed symbols after the first character: - ( ) / % . + , & apostrophe. Keep this workbook OPEN and click Import in Define Name to save your edits.', 'Packing List Manager')
    temp = path.with_name(path.stem + '.tmp-' + uuid.uuid4().hex + '.xlsx')
    try:
        wb.save(temp)
        os.replace(temp, path)
    except PermissionError as exc:
        raise PackingError('Close this Excel file in Excel, then export again.') from exc
    finally:
        wb.close()
        temp.unlink(missing_ok=True)
    return {'ok': True, 'path': str(path), 'count': len(catalog['rows'])}


def import_excel(path, current_rows):
    """Return a validated merge preview. Persistence happens only on Save names."""
    from openpyxl import load_workbook
    if path.suffix.lower() != '.xlsx':
        raise PackingError('Choose an .xlsx file exported from Define Name.')
    if path.stat().st_size > 20 * 1024 * 1024:
        raise PackingError('This workbook is too large. Import only the names worksheet.')
    rows = validate_rows(copy.deepcopy(current_rows))
    by_id = {r['id']: i for i, r in enumerate(rows)}
    by_key = {item_key(r['original_product'], r['original_variation']): i for i, r in enumerate(rows)}
    updates, imported_keys, imported_ids, targets = [], set(), set(), set()
    added = updated = 0
    try:
        wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
        try:
            if 'Names' not in wb.sheetnames:
                raise PackingError('The workbook needs a worksheet named Names.')
            ws = wb['Names']
            if ws.max_row > 20001:
                raise PackingError('Import at most 20,000 names at a time.')
            if [ws.cell(1, c).value for c in range(1, 5)] != list(HEADERS[:4]):
                raise PackingError('Keep the four column headings from a Define Name export.')
            has_ids = ws.cell(1, 5).value == HEADERS[4]
            for number, cells in enumerate(ws.iter_rows(min_row=2, max_col=5), 2):
                if not any(c.value is not None and c.value != '' for c in cells[:4]):
                    continue
                values = []
                for cell in cells[:4]:
                    if cell.data_type in ('f', 'e') or isinstance(cell.value, bool):
                        raise PackingError(f'Excel row {number}: use plain text, not formulas, errors, or TRUE/FALSE values.')
                    if cell.value is not None and not isinstance(cell.value, (str, int, float)):
                        raise PackingError(f'Excel row {number}: format names as text.')
                    values.append('' if cell.value is None else str(cell.value))
                row = dict(zip(FIELDS, values))
                identifier = str(cells[4].value or '') if has_ids else ''
                key = item_key(row['original_product'], row['original_variation'])
                if key in imported_keys or (identifier and identifier in imported_ids):
                    raise PackingError(f'Excel row {number}: duplicate original pair or entry ID.')
                imported_keys.add(key)
                if identifier:
                    imported_ids.add(identifier)
                target = by_id.get(identifier) if identifier in by_id else by_key.get(key)
                if target is not None:
                    if target in targets:
                        raise PackingError(f'Excel row {number}: two rows update the same saved item.')
                    targets.add(target)
                    row['id'] = rows[target]['id']
                    updated += 1
                else:
                    row['id'] = uuid.uuid4().hex
                    added += 1
                try:
                    row = validate_rows([row])[0]
                except PackingError as exc:
                    raise PackingError(f'Excel row {number}: {exc}') from exc
                updates.append((target, row))
        finally:
            wb.close()
    except PackingError:
        raise
    except Exception as exc:
        raise PackingError(f'Could not read the names workbook: {exc}') from exc
    for target, row in updates:
        if target is None:
            rows.append(row)
        else:
            rows[target] = row
    return {'ok': True, 'rows': validate_rows(rows), 'added': added, 'updated': updated}
