import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / 'src'))
from app_paths import default_data_dir, initialize_user_data


class AppPathTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=APP / 'qa')
        self.base = Path(self.temp.name)
        self.root = self.base / 'Program Files' / 'Packing List Manager'
        self.source = self.root / 'App Files' / 'Source' / 'app.py'
        self.source.parent.mkdir(parents=True)
        self.local = self.base / 'LocalAppData'
        self.defaults = self.root / 'App Files' / 'Defaults'
        self.seed(self.defaults, 'BUNDLED', 3)
        (self.defaults / 'SkuCatalog.json').write_text('{"seed":true}')

    def tearDown(self):
        self.temp.cleanup()

    def seed(self, path, name, schema):
        path.mkdir(parents=True, exist_ok=True)
        (path / 'ProductNames.json').write_text(json.dumps({'schema':schema,'name':name}))
        (path / 'Settings.json').write_text(json.dumps({'folder':name}))

    def test_portable_and_source_paths_remain_local(self):
        self.assertEqual(default_data_dir(self.source), self.root)
        self.assertEqual(default_data_dir(APP/'src/app.py'), APP.parent/'Packing List Manager')

    def test_installed_app_writes_user_folder_and_preserves_all_program_files(self):
        (self.root/'App Files/InstallMode.ini').write_text('[Setup]\nDataMode=user\n')
        before={p:p.read_bytes() for p in self.root.rglob('*') if p.is_file()}
        with patch.dict(os.environ, LOCALAPPDATA=str(self.local)):
            result=default_data_dir(self.source)
        self.assertEqual(result,self.local/'Packing List Manager')
        self.assertEqual(json.loads((result/'ProductNames.json').read_text())['name'],'BUNDLED')
        self.assertTrue(all(p.read_bytes()==v for p,v in before.items()))

    def test_newer_per_user_v2_wins_over_old_program_files_v1(self):
        self.seed(self.root,'OLD V1',2)
        legacy=self.local/'Programs/Packing List Manager'
        self.seed(legacy,'EDITED V2',3)
        (legacy/'ProductNames.json.bak').write_text('USER BACKUP')
        result=initialize_user_data(self.root,self.local)
        self.assertEqual(json.loads((result/'ProductNames.json').read_text())['name'],'EDITED V2')
        self.assertEqual(json.loads((result/'Settings.json').read_text())['folder'],'EDITED V2')
        self.assertEqual((result/'ProductNames.json.bak').read_text(),'USER BACKUP')
        self.assertTrue((result/'SkuCatalog.json').is_file())
        self.assertEqual(json.loads((self.root/'ProductNames.json').read_text())['name'],'OLD V1')

    def test_reinstall_never_overwrites_edits_and_can_resume_missing_file(self):
        result=initialize_user_data(self.root,self.local)
        (result/'ProductNames.json').write_text('MY EDITED DATA')
        (result/'Settings.json').write_text('MY SETTINGS')
        (result/'SkuCatalog.json').unlink()
        initialize_user_data(self.root,self.local)
        self.assertEqual((result/'ProductNames.json').read_text(),'MY EDITED DATA')
        self.assertEqual((result/'Settings.json').read_text(),'MY SETTINGS')
        self.assertTrue((result/'SkuCatalog.json').is_file())

    def test_existing_program_files_v2_is_preferred(self):
        self.seed(self.root,'CURRENT V2',3)
        self.seed(self.local/'Programs/Packing List Manager','OLDER V2',3)
        result=initialize_user_data(self.root,self.local)
        self.assertEqual(json.loads((result/'ProductNames.json').read_text())['name'],'CURRENT V2')


if __name__=='__main__':
    unittest.main()
