"""Exercise the real installer in an isolated folder, keeping desktop settings intact."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile

APP = Path(__file__).resolve().parents[1]


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run(command, timeout=180):
    result = subprocess.run(command, capture_output=True, timeout=timeout,
                            creationflags=subprocess.CREATE_NO_WINDOW)
    if result.returncode:
        raise RuntimeError(f'Command failed ({result.returncode}): {command}\n{result.stdout!r}\n{result.stderr!r}')
    return result


def verify(setup, package, source=None):
    setup, package = Path(setup).resolve(), Path(package).resolve()
    root = Path(tempfile.mkdtemp(prefix='installer-check-', dir=APP / 'qa'))
    installed = root / 'Packing List Manager'
    manifest = json.loads((package / 'Package Manifest.json').read_text(encoding='utf-8'))
    # NSIS requires an unquoted /D= path as the LAST argument; no shell is used.
    install_command = f'"{setup}" /S /NOINTEGRATION /D={installed}'
    print('Installing into isolated QA folder...', flush=True)
    run(install_command)
    for relative, expected in manifest['files_sha256'].items():
        assert digest(installed / relative) == expected, relative
    assert (installed / 'Uninstall.exe').is_file()
    print('Every installed program and initial catalog file matches the release.', flush=True)

    ui_report = root / 'ui.json'
    run([str(installed / 'Packing List Manager.exe'), '--self-test', str(ui_report)])
    assert json.loads(ui_report.read_text())['ok']
    runtime = installed / 'App Files' / 'Runtime' / 'python.exe'
    engine = installed / 'App Files' / 'Source' / 'engine.py'
    calls = 0
    def engine_call(request):
        nonlocal calls
        calls += 1
        request_path, response_path = root / f'request-{calls}.json', root / f'response-{calls}.json'
        request_path.write_text(json.dumps(request), encoding='utf-8')
        run([str(runtime), '-B', str(engine), '--request', str(request_path), '--response', str(response_path)])
        result = json.loads(response_path.read_text(encoding='utf-8'))
        assert result['ok'], result
        return result

    catalog = engine_call({'action': 'names_list'})
    assert catalog['rows'] and all('id' in row for row in catalog['rows'])
    assert not (installed / 'Packing List Manager').exists(), 'Unexpected nested catalog path'
    result = {'ok': True, 'setup': str(setup), 'setup_sha256': digest(setup),
              'installed_files_match': True, 'packaged_ui': True, 'catalog_rows': len(catalog['rows'])}
    if source:
        source = Path(source).resolve()
        before = digest(source)
        folder = root / '20261007'
        target = folder / 'LAZADA' / source.name
        target.parent.mkdir(parents=True)
        shutil.copy2(source, target)
        scanned = engine_call({'action': 'scan', 'folder': str(folder)})
        entries = scanned['entries']
        assert len(entries) == 1 and entries[0]['status'] == 'READY', entries
        e = entries[0]
        assert (e['orders'], e['shipments'], e['pieces']) == (13, 14, 38)
        print('Installed runtime: creating and checking the real Lazada PDF...', flush=True)
        generated = engine_call({'action': 'generate', 'folder': str(folder),
                                'selected': [{'path': e['path'], 'sha256': e['sha256']}], 'output_format': 'pdf'})
        assert len(generated['outputs']) == 1
        output = generated['outputs'][0]
        assert output['pages'] == 3
        rescanned = engine_call({'action': 'scan', 'folder': str(folder)})
        assert rescanned['entries'][0]['status'] == 'DONE'
        history_path = folder / '000 Output' / '.PackingData' / 'Processing_History.json'
        history = json.loads(history_path.read_text(encoding='utf-8'))
        snapshots = history['outputs'][output['filename']]['order_snapshots']
        split = [o for o in snapshots.values() if o['order_id'] == '1128986725736007']
        assert {o['tracking_id']: sum(i['pieces'] for i in o['items']) for o in split} == {'820165464947': 4, '310004587637': 16}
        assert digest(source) == before
        result.update(pdf=output['path'], orders=13, shipments=14, pieces=38, pages=3,
                      source_unchanged=True, rescan='DONE')
        preserved_output = {p: digest(p) for p in (folder / '000 Output').rglob('*') if p.is_file()}

    # Represent real post-install edits; reinstall must preserve all of them.
    data_path = installed / 'ProductNames.json'
    saved = json.loads(data_path.read_text(encoding='utf-8'))
    saved['rows'][0]['alternative_product'] = 'QA SAVED NAME - MUST SURVIVE'
    data_path.write_text(json.dumps(saved), encoding='utf-8')
    sku_path = installed / 'SkuCatalog.json'
    sku_path.write_text(sku_path.read_text(encoding='utf-8') + '\n', encoding='utf-8')
    (installed / 'Settings.json').write_text('{"folder":"QA saved folder"}', encoding='utf-8')
    # Unknown local files are also outside the uninstaller's exact release inventory.
    (installed / 'App Files' / 'my-backup.txt').write_text('Keep this user file.', encoding='utf-8')
    keep = {installed / name: digest(installed / name) for name in
            ('ProductNames.json', 'SkuCatalog.json', 'Settings.json', 'App Files/my-backup.txt')}
    print('Reinstalling to verify saved edits are preserved...', flush=True)
    run(install_command)
    assert all(digest(path) == expected for path, expected in keep.items())
    result['reinstall_preserves_data'] = True
    print('Uninstalling program files while retaining data and outputs...', flush=True)
    # _?= runs the actual uninstaller in place and makes completion waitable.
    run(f'"{installed / "Uninstall.exe"}" /S _?={installed}')
    assert not (installed / 'Packing List Manager.exe').exists()
    assert all(digest(path) == expected for path, expected in keep.items())
    if source:
        assert all(digest(path) == expected for path, expected in preserved_output.items())
    result.update(uninstall_preserves_data=True, uninstall_preserves_outputs=True)
    (root / 'verification.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result, indent=2), flush=True)
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('setup')
    parser.add_argument('package')
    parser.add_argument('--source', type=Path)
    args = parser.parse_args()
    verify(args.setup, args.package, args.source)
