"""Copy a standard CPython runtime and the engine's installed dependencies.

No frozen bootloader, executable rewriting, downloads, or security exclusions.
The ._pth file isolates imports from machine/user Python installations.
"""
from importlib import metadata
from pathlib import Path
import argparse
import json
import shutil
import sys
import zipfile

from packaging.requirements import Requirement


def package(destination):
    if sys.version_info[:2] != (3, 12) or sys.maxsize <= 2**32:
        raise RuntimeError('Build with 64-bit CPython 3.12.')
    project = Path(__file__).resolve().parent
    destination = Path(destination).resolve()
    # Require a fresh package rather than retaining files from an older build.
    destination.mkdir(parents=True, exist_ok=False)
    runtime = destination / 'Runtime'
    runtime.mkdir()
    prefix = Path(sys.base_prefix)
    for filename in ('python.exe', 'python3.dll', 'python312.dll',
                     'vcruntime140.dll', 'vcruntime140_1.dll', 'LICENSE.txt'):
        shutil.copy2(prefix / filename, runtime / filename)
    shutil.copytree(prefix / 'DLLs', runtime / 'DLLs',
                    ignore=shutil.ignore_patterns('__pycache__', '*.pyc', '*tkinter*', 'tcl*', 'tk*'))
    ignored = {'site-packages', '__pycache__', 'test', 'tests', 'idlelib', 'tkinter', 'turtledemo', 'ensurepip'}
    with zipfile.ZipFile(runtime / 'python312.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
        for path in sorted((prefix / 'Lib').rglob('*')):
            relative = path.relative_to(prefix / 'Lib')
            if path.is_file() and not ignored.intersection(relative.parts) and path.suffix != '.pyc':
                archive.write(path, relative.as_posix())
    packages = runtime / 'Lib' / 'site-packages'
    packages.mkdir(parents=True)
    pending = []
    for line in (project / 'requirements.txt').read_text(encoding='utf-8-sig').splitlines():
        line = line.strip()
        if line and not line.startswith('#'):
            pending.append(Requirement(line))
    versions = {}
    while pending:
        requirement = pending.pop()
        if requirement.marker and not requirement.marker.evaluate():
            continue
        distribution = metadata.distribution(requirement.name)
        key = distribution.metadata['Name'].lower().replace('_', '-')
        if distribution.version not in requirement.specifier:
            raise RuntimeError(f'{requirement}: installed version is {distribution.version}')
        if key in versions:
            continue
        versions[key] = distribution.version
        for entry in distribution.files or []:
            if '..' in entry.parts or '__pycache__' in entry.parts or entry.suffix == '.pyc':
                continue
            original = Path(distribution.locate_file(entry))
            target = packages / str(entry)
            if original.is_file():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(original, target)
        pending.extend(Requirement(value) for value in distribution.requires or [])
    (runtime / 'python312._pth').write_text(
        'python312.zip\nDLLs\nLib/site-packages\n../Source\n.\n', encoding='utf-8')
    source = destination / 'Source'
    source.mkdir()
    for path in (project / 'src').glob('*.py'):
        shutil.copy2(path, source / path.name)
    shutil.copytree(project / 'src' / 'assets', source / 'assets',
                    ignore=shutil.ignore_patterns('picking_aliases.json'))
    (destination / 'Runtime Manifest.json').write_text(json.dumps({
        'packaging': 'standard CPython with readable engine source',
        'python': sys.version, 'dependencies': versions,
    }, indent=2), encoding='utf-8')
    print(json.dumps({'runtime': str(runtime), 'dependencies': versions}, indent=2))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('destination')
    package(parser.parse_args().destination)
