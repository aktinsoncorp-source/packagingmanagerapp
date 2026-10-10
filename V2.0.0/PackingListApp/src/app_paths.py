"""Writable per-user data for installed apps; portable/source copies stay local."""
import configparser
import json
import os
from pathlib import Path
import shutil
import uuid

DATA_FILES = ('ProductNames.json', 'ProductNames.json.bak', 'SkuCatalog.json', 'Settings.json')


def initialize_user_data(root, local):
    root, local = Path(root), Path(local)
    destination = local / 'Packing List Manager'
    legacy = local / 'Programs' / 'Packing List Manager'
    defaults = root / 'App Files' / 'Defaults'
    # A previous V2 in this installation wins; otherwise prefer the newer V2
    # per-user installation over a V1 catalog still in Program Files.
    def is_v2(folder):
        try:
            return json.loads((folder / 'ProductNames.json').read_text(encoding='utf-8-sig')).get('schema') == 3
        except (OSError, ValueError):
            return False
    candidates = ([root, legacy] if is_v2(root) else [legacy, root]) + [defaults]
    source = next((p for p in candidates if (p / 'ProductNames.json').is_file()), defaults)
    destination.mkdir(parents=True, exist_ok=True)
    for name in DATA_FILES:
        target = destination / name
        if target.exists():
            continue
        original = source / name
        if not original.is_file() and name == 'SkuCatalog.json':
            original = defaults / name
        if not original.is_file():
            continue
        temporary = destination / (name + '.migrate-' + uuid.uuid4().hex)
        try:
            shutil.copyfile(original, temporary)
            # Windows rename fails if another startup already seeded this file.
            try:
                os.rename(temporary, target)
            except FileExistsError:
                pass
        finally:
            temporary.unlink(missing_ok=True)
    return destination


def default_data_dir(source_file=__file__):
    source = Path(source_file).resolve().parent
    if source.name != 'Source' or source.parent.name != 'App Files':
        return source.parents[1] / 'Packing List Manager'
    root = source.parents[1]
    mode = configparser.ConfigParser()
    mode.read(root / 'App Files' / 'InstallMode.ini', encoding='utf-8')
    if mode.get('Setup', 'DataMode', fallback='portable') != 'user':
        return root
    local = os.environ.get('LOCALAPPDATA')
    if not local:
        raise RuntimeError('Windows Local AppData could not be located for saved product data.')
    return initialize_user_data(root, Path(local))
