"""Fault injection; no real trust store, registry or Office writes."""
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

SPEC=importlib.util.spec_from_file_location('trial',Path(__file__).parents[1]/'app/trial-launcher.py')
m=importlib.util.module_from_spec(SPEC);SPEC.loader.exec_module(m)

class Safety(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.root=Path(self.tmp.name)
        self.stack=[]
        for key,name in [('DATA','data'),('STATE','data/installation.json'),('CERT','data/localhost.crt'),('KEY','data/localhost.key'),('ROOT','app')]:
            p=patch.object(m,key,self.root/name);p.start();self.stack.append(p)
        m.DATA.mkdir();m.ROOT.mkdir();(m.ROOT/'manifests').mkdir()
    def tearDown(self):
        for p in reversed(self.stack):p.stop()
        self.tmp.cleanup()
    def state(self):
        m.STATE.write_text(json.dumps({'thumbprint':'ABCD'}));m.CERT.write_text('test-cert');m.KEY.write_text('test-key')
    def test_busy_unknown_process_prevents_uninstall(self):
        self.state()
        with patch.object(m,'WIN',True),patch.object(m,'validate_certificate'),patch.object(m,'notice',return_value=True),patch.object(m,'port_busy',return_value=True),patch.object(m,'health',return_value=False):
            with self.assertRaises(RuntimeError):m.uninstall()
        self.assertTrue(m.STATE.exists());self.assertTrue(m.KEY.exists())
    def test_missing_token_prevents_stop(self):
        with patch.object(m,'port_busy',return_value=True),patch.object(m,'health',return_value=True):
            with self.assertRaises(RuntimeError):m.stop()
    def test_stop_timeout_fails(self):
        (m.DATA/'control-token').write_text('test-token')
        from unittest.mock import MagicMock
        with patch.object(m,'port_busy',return_value=True),patch.object(m,'health',return_value=True),patch.object(m.ssl,'create_default_context'),patch.object(m.urllib.request,'urlopen',return_value=MagicMock()),patch.object(m.time,'sleep'):
            with self.assertRaises(RuntimeError):m.stop()
    def test_no_listener_needs_no_control_token(self):
        with patch.object(m,'port_busy',return_value=False):m.stop()
    def test_install_retries_missing_trust(self):
        self.state()
        with patch.object(m,'WIN',False),patch.object(m,'protect_data'),patch.object(m,'certificate_trusted',return_value=False),patch.object(m,'notice',return_value=True),patch.object(m,'create_certificate') as create,patch.object(m,'trust_certificate') as trust,patch.object(m,'register_manifests'),patch.object(m,'start'):
            m.install();create.assert_not_called();trust.assert_called_once()
    def test_failed_trust_never_registers_or_starts(self):
        self.state()
        with patch.object(m,'protect_data'),patch.object(m,'certificate_trusted',return_value=False),patch.object(m,'notice',return_value=True),patch.object(m,'trust_certificate',side_effect=RuntimeError('denied')),patch.object(m,'register_manifests') as register,patch.object(m,'start') as start:
            with self.assertRaises(RuntimeError):m.install()
            register.assert_not_called();start.assert_not_called()
    def test_incomplete_record_blocks_new_certificate(self):
        m.STATE.write_text('{}')
        with patch.object(m,'protect_data'),patch.object(m,'create_certificate') as create:
            with self.assertRaises(RuntimeError):m.install()
            create.assert_not_called()
    def test_windows_acl_command_protects_inheritance(self):
        with patch.object(m,'WIN',True),patch.object(m,'powershell') as ps:
            m.protect_data()
            code=ps.call_args.args[0]
            self.assertIn('SetAccessRuleProtection($true,$false)',code)
            self.assertIn('WindowsIdentity]::GetCurrent().User',code)
            self.assertIn('FileSecurity',code)
            self.assertIn('Set-Acl -LiteralPath $file.FullName',code)
            self.assertNotIn('Everyone',code)
    def test_windows_trust_not_assumed_after_import(self):
        with patch.object(m,'WIN',True),patch.object(m,'validate_certificate'),patch.object(m,'run'),patch.object(m,'certificate_trusted',return_value=False):
            with self.assertRaises(RuntimeError):m.trust_certificate()
    def test_mac_delete_failure_preserves_record(self):
        self.state();m.STATE.write_text(json.dumps({'sha1':'ABCD','trusted':False}))
        def run(args,**kw):
            if 'find-certificate' in args:return subprocess.CompletedProcess(args,0,stdout='SHA-1 hash: ABCD')
            if 'delete-certificate' in args:raise subprocess.CalledProcessError(1,args)
            return subprocess.CompletedProcess(args,0,stdout='')
        with patch.object(m,'WIN',False),patch.object(m,'validate_certificate'),patch.object(m,'notice',return_value=True),patch.object(m,'stop'),patch.object(m.subprocess,'run'),patch.object(m,'run',side_effect=run),patch.object(m.Path,'home',return_value=self.root):
            with self.assertRaises(subprocess.CalledProcessError):m.uninstall()
            self.assertTrue(m.STATE.exists());self.assertTrue(m.KEY.exists())
    def test_mismatched_certificate_blocks_all_cleanup(self):
        self.state()
        with patch.object(m,'notice',return_value=True),patch.object(m,'validate_certificate',side_effect=RuntimeError('mismatch')),patch.object(m,'run') as run,patch.object(m,'stop') as stop:
            with self.assertRaises(RuntimeError):m.uninstall()
            run.assert_not_called();stop.assert_not_called()
            self.assertTrue(m.STATE.exists())
    def test_windows_certificate_delete_error_is_not_success(self):
        from unittest.mock import MagicMock
        self.state()
        with patch.object(m,'WIN',True),patch.object(m,'notice',return_value=True),patch.object(m,'validate_certificate'),patch.object(m,'stop'),patch.dict('sys.modules',{'winreg':MagicMock()}),patch.object(m,'store_has_certificate',return_value=True),patch.object(m,'run',side_effect=subprocess.CalledProcessError(1,['certutil'])):
            with self.assertRaises(subprocess.CalledProcessError):m.uninstall()
            self.assertTrue(m.STATE.exists());self.assertTrue(m.KEY.exists())

    def test_mac_success_reads_completed_process_stdout(self):
        self.state();m.STATE.write_text(json.dumps({'sha1':'ABCD','trusted':False}))
        removed=False
        def run(args,**kw):
            nonlocal removed
            if 'delete-certificate' in args:removed=True
            output='SHA-1 hash: ABCD' if 'find-certificate' in args and not removed else ''
            return subprocess.CompletedProcess(args,0,stdout=output)
        with patch.object(m,'WIN',False),patch.object(m,'validate_certificate'),patch.object(m,'notice',return_value=True),patch.object(m,'stop'),patch.object(m.subprocess,'run'),patch.object(m,'run',side_effect=run),patch.object(m.Path,'home',return_value=self.root):
            m.uninstall()
        self.assertTrue(removed);self.assertFalse(m.STATE.exists())

if __name__=='__main__':unittest.main()
