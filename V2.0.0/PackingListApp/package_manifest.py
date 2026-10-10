"""Record exact release hashes and generate a non-recursive uninstall file list."""
from pathlib import Path
import argparse
import hashlib
import json
import subprocess

PRESERVE = {'ProductNames.json', 'SkuCatalog.json', 'Settings.json', 'Startup.log'}


def prepare(application, removal_list):
    application = Path(application).resolve()
    allowed = {'App Files', 'Packing List Manager.exe', 'Packing List Manager.ico',
               'Employee Guide.txt', 'Release Notes.txt', 'ProductNames.json',
               'SkuCatalog.json', 'Package Manifest.json'}
    unexpected = {p.name for p in application.iterdir()} - allowed
    if unexpected:
        raise RuntimeError('Build from a fresh release folder; unexpected files: ' + ', '.join(sorted(unexpected)))
    files = {p.relative_to(application).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
             for p in sorted(application.rglob('*')) if p.is_file() and p.name != 'Package Manifest.json'}
    commit = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=Path(__file__).parent, text=True).strip()
    manifest = {'app': 'Packing List Manager', 'version': '2.0.0', 'source_commit': commit,
                'files_sha256': files, 'preserve_on_upgrade_and_uninstall': sorted(PRESERVE)}
    (application / 'Package Manifest.json').write_text(json.dumps(manifest, indent=2), encoding='utf-8')
    def nsis(value):
        return value.replace('$', '$$').replace('"', '$\\"').replace('/', '\\')
    lines = ['; Generated from the exact release inventory. User data is excluded.']
    for relative in [*files, 'Package Manifest.json']:
        if relative not in PRESERVE:
            lines.append(f'Delete "$INSTDIR\\{nsis(relative)}"')
    directories = sorted((p.relative_to(application) for p in application.rglob('*') if p.is_dir()),
                         key=lambda path: (-len(path.parts), str(path)))
    for relative in directories:
        lines.append(f'RMDir "$INSTDIR\\{nsis(relative.as_posix())}"')
    Path(removal_list).write_text('\n'.join(lines) + '\n', encoding='utf-8')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('application')
    parser.add_argument('removal_list')
    args = parser.parse_args()
    prepare(args.application, args.removal_list)
