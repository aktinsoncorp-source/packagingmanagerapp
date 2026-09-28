"""Hidden processing records, legacy migration, locking and crash recovery."""
import json,os,shutil,stat,unittest
from pathlib import Path
from unittest.mock import patch
import test_regeneration as fixtures
import engine
from extractors import PackingError,file_hash

class InternalStorageTests(unittest.TestCase):
    setUp=fixtures.RegenerationTests.setUp
    tearDown=fixtures.RegenerationTests.tearDown
    parse=fixtures.RegenerationTests.parse
    source=fixtures.RegenerationTests.source
    scan=fixtures.RegenerationTests.scan
    generate=fixtures.RegenerationTests.generate
    history=fixtures.RegenerationTests.history
    def prepared(self):
        self.source('one.pdf',['A','B']);output=self.generate()[0]
        return self.folder/'000 Output',output
    def test_new_outputs_keep_only_packing_documents_visible(self):
        out,output=self.prepared()
        self.assertEqual({p.name for p in out.iterdir()},{engine.INTERNAL,output['filename']})
        self.assertTrue((out/engine.HISTORY).is_file());self.assertTrue((out/engine.LOG).is_file());self.assertTrue((out/engine.LOCK).is_file())
        self.assertTrue((out/engine.INTERNAL).stat().st_file_attributes & stat.FILE_ATTRIBUTE_HIDDEN)
        self.assertFalse(list((out/engine.INTERNAL).glob('.pending-*')))
    def test_existing_records_migrate_byte_for_byte_and_keep_courier_choices(self):
        out,output=self.prepared();source=self.folder/'SPX/one.pdf'
        choices={'schema':1,'choices':{'SPX/one.pdf':{'sha256':file_hash(source),'carrier':'SPX'}}}
        engine.atomic_json(out/engine.COURIERS,choices)
        originals={r:(out/r).read_bytes() for r in (engine.HISTORY,engine.LOG,engine.COURIERS,engine.LOCK)}
        doc_hash=file_hash(output['path'])
        for relative in originals:os.rename(out/relative,out/Path(relative).name)
        scan=self.scan();self.assertEqual(scan['entries'][0]['status'],'DONE')
        self.assertEqual(engine.load_couriers(self.folder),choices['choices'])
        for relative,data in originals.items():
            self.assertEqual((out/relative).read_bytes(),data);self.assertFalse((out/Path(relative).name).exists())
        self.assertEqual(file_hash(output['path']),doc_hash)
    def test_partly_migrated_date_resumes(self):
        out,_=self.prepared();data=(out/engine.HISTORY).read_bytes()
        os.rename(out/engine.HISTORY,out/'Processing_History.json')
        self.scan();self.assertEqual((out/engine.HISTORY).read_bytes(),data)
        self.assertFalse((out/'Processing_History.json').exists());self.assertTrue((out/engine.LOG).exists())
    def test_identical_legacy_duplicate_is_removed(self):
        out,_=self.prepared();shutil.copyfile(out/engine.HISTORY,out/'Processing_History.json')
        self.scan();self.assertFalse((out/'Processing_History.json').exists())
    def test_conflicting_copies_stop_without_overwriting_or_moving_any_record(self):
        out,_=self.prepared();before=(out/engine.HISTORY).read_bytes()
        (out/'Processing_History.json').write_text('{"different":true}',encoding='utf-8')
        os.rename(out/engine.LOG,out/'Processed_Files.txt')
        with self.assertRaisesRegex(PackingError,'Two different copies'):self.scan()
        self.assertEqual((out/engine.HISTORY).read_bytes(),before)
        self.assertEqual((out/'Processing_History.json').read_text(),'{"different":true}')
        self.assertTrue((out/'Processed_Files.txt').exists());self.assertFalse((out/engine.LOG).exists())
    def test_active_legacy_lock_prevents_migration(self):
        out,_=self.prepared();os.rename(out/engine.HISTORY,out/'Processing_History.json')
        with engine.file_lock(out/'.packing.lock'):
            with self.assertRaisesRegex(PackingError,'another window'):self.scan()
            self.assertTrue((out/'Processing_History.json').exists())
        self.scan();self.assertFalse((out/'.packing.lock').exists());self.assertTrue((out/engine.HISTORY).exists())
    def test_hidden_lock_still_prevents_concurrent_processing(self):
        out,_=self.prepared()
        with engine.folder_lock(self.folder):
            with self.assertRaisesRegex(PackingError,'another window'):self.scan()
    def test_hidden_stage_and_journal_recover_once(self):
        out=self.folder/'000 Output'
        with engine.folder_lock(self.folder):
            stage=out/engine.INTERNAL/'.pending-test';stage.mkdir()
            source=stage/'SPX_Packing_Checklist_2026-09-26_1.pdf';source.write_bytes(b'recoverable verified fixture')
            files=[{'target':source.name,'stage':source.relative_to(out).as_posix(),'sha256':file_hash(source)}]
            history={'schema':1,'date':self.folder.name,'sources':{},'orders':{},'outputs':{}}
            with patch.object(engine,'recover',side_effect=RuntimeError('simulated interruption')):
                with self.assertRaisesRegex(RuntimeError,'interruption'):engine.commit(out,history,files,stage,'Recovery fixture')
            self.assertTrue((out/engine.JOURNAL).is_file())
            plan=json.loads((out/engine.JOURNAL).read_text());self.assertEqual(plan['stage_dir'],engine.INTERNAL+'/.pending-test')
        with engine.folder_lock(self.folder):engine.recover(out);engine.recover(out)
        self.assertEqual((out/source.name).read_bytes(),b'recoverable verified fixture')
        self.assertFalse(stage.exists());self.assertFalse((out/engine.JOURNAL).exists())
        self.assertEqual((out/engine.LOG).read_text().count('Recovery fixture'),1)
