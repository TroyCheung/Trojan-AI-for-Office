"""Per-user installer/launcher for the isolated Office AI trial edition."""
import base64
import ctypes
import hashlib
import json
import os
from pathlib import Path
import plistlib
import secrets
import socket
import ssl
import subprocess
import sys
import threading
import time
import urllib.request
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
WIN = sys.platform == 'win32'
DATA = (Path(os.environ['LOCALAPPDATA']) / 'OfficeAITrial' / 'user-data' if WIN
        else Path.home() / 'Library/Application Support/OfficeAITrial')
PORT = 18443
LABEL = 'local.office-ai-trial'
ORIGIN = f'https://localhost:{PORT}'
STATE = DATA / 'installation.json'
CERT = DATA / 'localhost.crt'
KEY = DATA / 'localhost.key'
os.environ['OFFICE_AI_TRIAL_DATA'] = str(DATA)
os.environ['SSL_CERT_FILE'] = str(ROOT / 'cacert.pem')
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
if not WIN:
    os.umask(0o077)


def run(args, **kwargs):
    return subprocess.run([str(a) for a in args], check=True, capture_output=True, text=True, **kwargs)


def notice(text, confirm=False):
    if WIN:
        return ctypes.windll.user32.MessageBoxW(None, text, 'Trojan AI for Office', 0x24 if confirm else 0x40) == 6
    script = 'display dialog ' + json.dumps(text, ensure_ascii=False) + ' with title "Trojan AI for Office" '
    script += 'buttons {"取消", "继续"} default button "继续"' if confirm else 'buttons {"好"} default button "好"'
    try:
        run(['/usr/bin/osascript', '-e', script])
        return True
    except subprocess.CalledProcessError:
        return False


def powershell(code):
    return run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', code]).stdout.strip()


def ps_quote(value):
    return "'" + str(value).replace("'", "''") + "'"


def protect_data():
    DATA.mkdir(parents=True, exist_ok=True)
    if WIN:
        powershell("$ErrorActionPreference='Stop'; $p="+ps_quote(DATA)+r""";
$acl=New-Object System.Security.AccessControl.DirectorySecurity
$acl.SetAccessRuleProtection($true,$false)
$sid=[System.Security.Principal.WindowsIdentity]::GetCurrent().User
$acl.SetOwner($sid)
foreach($id in @($sid,(New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')))) {
  $rule=New-Object System.Security.AccessControl.FileSystemAccessRule($id,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
  $acl.AddAccessRule($rule)
}
Set-Acl -LiteralPath $p -AclObject $acl
foreach($file in Get-ChildItem -LiteralPath $p -File -Force) {
  if($file.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Refusing linked data file'}
  $fa=New-Object System.Security.AccessControl.FileSecurity
  $fa.SetAccessRuleProtection($true,$false)
  $fa.SetOwner($sid)
  foreach($id in @($sid,(New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')))) {
    $fa.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($id,'FullControl','Allow')))
  }
  Set-Acl -LiteralPath $file.FullName -AclObject $fa
}
""")
    else:
        DATA.chmod(0o700)


def validate_certificate():
    state=json.loads(STATE.read_text())
    expected=state.get('thumbprint') if WIN else state.get('sha1')
    actual=hashlib.sha1(ssl.PEM_cert_to_DER_cert(CERT.read_text())).hexdigest()
    if not expected or actual.lower()!=expected.lower():
        raise RuntimeError('本机证书与安装记录不一致，已停止操作，请保留安装目录供排查。')
    ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER).load_cert_chain(CERT,KEY)
    return expected


def store_has_certificate(store, thumb):
    return powershell('Test-Path -LiteralPath '+ps_quote('Cert:\\CurrentUser\\'+store+'\\'+thumb)).lower()=='true'


def certificate_trusted():
    thumb=validate_certificate()
    if WIN:
        return store_has_certificate('Root',thumb)
    result=subprocess.run(['/usr/bin/security','verify-cert','-c',str(CERT),'-p','ssl','-s','localhost'],capture_output=True)
    return result.returncode==0


def trust_certificate():
    validate_certificate()
    if WIN:
        run(['certutil.exe','-user','-addstore','Root',CERT])
    else:
        run(['/usr/bin/security','add-trusted-cert','-r','trustRoot','-p','ssl','-k',Path.home()/'Library/Keychains/login.keychain-db',CERT])
    if not certificate_trusted():
        raise RuntimeError('系统尚未信任本机证书，初始化未完成。请保留安装记录后重试。')
    state=json.loads(STATE.read_text());state['trusted']=True
    STATE.write_text(json.dumps(state))


def der(tag, content):
    length = len(content)
    size = bytes([length]) if length < 128 else bytes([128 + (length.bit_length()+7)//8]) + length.to_bytes((length.bit_length()+7)//8, 'big')
    return bytes([tag]) + size + content


def rsa_pem(values):
    def integer(raw):
        raw = raw.lstrip(b'\0') or b'\0'
        return der(2, (b'\0' if raw[0] & 128 else b'') + raw)
    body = integer(b'\0') + b''.join(integer(base64.b64decode(values[k])) for k in ('Modulus','Exponent','D','P','Q','DP','DQ','InverseQ'))
    encoded = base64.b64encode(der(48, body)).decode()
    return '-----BEGIN RSA PRIVATE KEY-----\n' + '\n'.join(encoded[i:i+64] for i in range(0,len(encoded),64)) + '\n-----END RSA PRIVATE KEY-----\n'


def create_certificate():
    protect_data()
    if WIN:
        # Export only the new installation's key, never a user's existing certificate.
        code = r"""
$ErrorActionPreference='Stop'
$c=New-SelfSignedCertificate -DnsName 'localhost' -CertStoreLocation 'Cert:\CurrentUser\My' -FriendlyName 'Trojan AI for Office localhost' -KeyAlgorithm RSA -KeyLength 2048 -KeyExportPolicy Exportable -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider' -NotAfter (Get-Date).AddYears(1)
try {
$r=[System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($c)
$p=$r.ExportParameters($true)
$o=@{Certificate=[Convert]::ToBase64String($c.RawData);Thumbprint=$c.Thumbprint}
foreach($n in @('Modulus','Exponent','D','P','Q','DP','DQ','InverseQ')){$o[$n]=[Convert]::ToBase64String($p.$n)}
$o | ConvertTo-Json -Compress
} catch { Remove-Item -LiteralPath $c.PSPath -ErrorAction Stop; throw }
"""
        values=json.loads(powershell(code))
        STATE.write_text(json.dumps({'thumbprint':values['Thumbprint']}))
        CERT.write_text(ssl.DER_cert_to_PEM_cert(base64.b64decode(values['Certificate'])))
        KEY.write_text(rsa_pem(values))
    else:
        config=DATA/'certificate.cnf'
        config.write_text('[req]\nprompt=no\ndistinguished_name=dn\nx509_extensions=ext\n[dn]\nCN=Trojan AI for Office localhost\n[ext]\nsubjectAltName=DNS:localhost,IP:127.0.0.1\nbasicConstraints=critical,CA:FALSE\nkeyUsage=critical,digitalSignature,keyEncipherment\nextendedKeyUsage=serverAuth\n')
        run(['/usr/bin/openssl','req','-x509','-newkey','rsa:2048','-nodes','-days','365','-config',config,'-keyout',KEY,'-out',CERT])
        STATE.write_text(json.dumps({'sha1':hashlib.sha1(ssl.PEM_cert_to_DER_cert(CERT.read_text())).hexdigest()}))
        config.unlink()
    KEY.chmod(0o600)


def port_busy():
    try:
        with socket.create_connection(('127.0.0.1',PORT),timeout=.4):
            return True
    except ConnectionRefusedError:
        return False
    except OSError as error:
        raise RuntimeError('无法确认本机服务是否已经退出，已停止清理。') from error


def health():
    try:
        ctx=ssl.create_default_context(cafile=str(CERT))
        with urllib.request.urlopen(ORIGIN+'/trial-health',context=ctx,timeout=2) as response:
            return json.load(response).get('application') == LABEL
    except Exception:
        return False


def stop():
    tokenfile=DATA/'control-token'
    if not port_busy(): return
    if not health() or not tokenfile.exists():
        raise RuntimeError('端口仍在使用，但无法核对试用服务身份或关闭凭据。已停止卸载，保留文件供排查。')
    if health() and tokenfile.exists():
        req=urllib.request.Request(ORIGIN+'/trial-stop', data=b'', headers={'X-Trial-Control':tokenfile.read_text()})
        with urllib.request.urlopen(req,context=ssl.create_default_context(cafile=str(CERT)),timeout=3):
            pass
        for _ in range(50):
            if not port_busy(): return
            time.sleep(.1)
    raise RuntimeError('服务尚未退出，已停止卸载。请稍后重试。')


def serve():
    import server
    DATA.mkdir(parents=True,exist_ok=True)
    token=secrets.token_hex(32)
    class TrialHandler(server.Handler):
        def do_GET(self):
            if self.path == '/trial-health':
                self.send_json(200,{'application':LABEL,'version':'0.1.2'})
            else:
                super().do_GET()
        def do_POST(self):
            if self.path == '/trial-stop':
                if not secrets.compare_digest(self.headers.get('X-Trial-Control',''),token):
                    self.send_json(403,{'error':'forbidden'})
                    return
                self.send_json(200,{'stopping':True})
                threading.Thread(target=self.server.shutdown,daemon=True).start()
            else:
                super().do_POST()
    os.chdir(ROOT)
    httpd=server.ThreadingHTTPServer(('127.0.0.1',PORT),TrialHandler)
    context=ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version=ssl.TLSVersion.TLSv1_2
    context.load_cert_chain(CERT,KEY)
    httpd.socket=context.wrap_socket(httpd.socket,server_side=True)
    (DATA/'control-token').write_text(token)
    try: httpd.serve_forever()
    finally: httpd.server_close()


def register_manifests():
    for file in (ROOT/'manifests').glob('*.xml'):
        if WIN:
            import winreg
            ident=ET.parse(file).getroot().find('{http://schemas.microsoft.com/office/appforoffice/1.1}Id').text
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Office\16.0\WEF\Developer') as k:
                winreg.SetValueEx(k,ident,0,winreg.REG_SZ,str(file))
        else:
            app={'word':'Word','ppt':'Powerpoint','excel':'Excel'}[file.name.split('.')[1]]
            dest=Path.home()/f'Library/Containers/com.microsoft.{app}/Data/Documents/wef'
            dest.mkdir(parents=True,exist_ok=True)
            (dest/('office-ai-trial-'+file.name)).write_bytes(file.read_bytes())


def start():
    if health(): return
    if WIN:
        exe=ROOT.parent/'runtime/pythonw.exe'
        with (DATA/'service.log').open('a') as log:
            subprocess.Popen([str(exe),'-B',str(ROOT/'trial-launcher.py'),'--serve'],stdout=log,stderr=log,creationflags=0x08000000)
    else:
        exe=ROOT.parent/'runtime/bin/python3.13'
        plist=Path.home()/f'Library/LaunchAgents/{LABEL}.plist'
        plist.parent.mkdir(parents=True,exist_ok=True)
        content={'Label':LABEL,'ProgramArguments':[str(exe),'-B',str(ROOT/'trial-launcher.py'),'--serve'],'EnvironmentVariables':{'PYTHONHOME':str(ROOT.parent/'runtime'),'SSL_CERT_FILE':str(ROOT/'cacert.pem')},'RunAtLoad':True,'KeepAlive':False,'StandardOutPath':str(DATA/'service.log'),'StandardErrorPath':str(DATA/'service.log')}
        plist.write_bytes(plistlib.dumps(content))
        subprocess.run(['/bin/launchctl','bootout',f'gui/{os.getuid()}',str(plist)],capture_output=True)
        run(['/bin/launchctl','bootstrap',f'gui/{os.getuid()}',plist])
    for _ in range(30):
        if health(): return
        time.sleep(.3)
    raise RuntimeError('本地服务未成功启动。可能是端口被占用或安装不完整，请保留 service.log 供排查。')


def install():
    protect_data()
    complete=STATE.exists() and CERT.exists() and KEY.exists()
    if STATE.exists() and not complete:
        raise RuntimeError('上次安装留下了不完整的证书记录。请先卸载本试用版，再重新安装。')
    if not complete or not certificate_trusted():
        if not notice('将安装 Word、PowerPoint 和 Excel 试用插件，并为仅限本机的加密连接创建证书、加入当前用户的信任列表。系统可能再次询问确认。\n\n不包含模型额度，请填写自己的 API Key。是否继续？',True):
            raise RuntimeError('你取消了安装，未完成初始化。')
        if not complete: create_certificate()
        trust_certificate()
    register_manifests()
    start()
    if WIN:
        import winreg
        command='"'+str(ROOT.parent/'runtime/pythonw.exe')+'" -B "'+str(ROOT/'trial-launcher.py')+'" --serve'
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Run') as k:
            winreg.SetValueEx(k,'OfficeAITrial',0,winreg.REG_SZ,command)
    notice('已完成本机初始化。请保存文件并完全退出 Word、PowerPoint 和 Excel，再重新打开。\n\n在“开始 → 加载项”中找到 Trojan AI for Office，首次使用在设置里填写自己的 API Key。建议先用文件副本试用。')


def uninstall():
    if not notice('移除 Trojan AI for Office 的加载入口、自动启动和专用证书？\n\n保留你的设置与历史记录，不影响其他插件。',True):
        raise RuntimeError('已取消卸载。')
    if STATE.exists(): validate_certificate()
    if not WIN:
        plist=Path.home()/f'Library/LaunchAgents/{LABEL}.plist'
        subprocess.run(['/bin/launchctl','bootout',f'gui/{os.getuid()}',str(plist)],capture_output=True)
        plist.unlink(missing_ok=True)
    stop()
    for file in (ROOT/'manifests').glob('*.xml'):
        if WIN:
            import winreg
            ident=ET.parse(file).getroot().find('{http://schemas.microsoft.com/office/appforoffice/1.1}Id').text
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Office\16.0\WEF\Developer',0,winreg.KEY_SET_VALUE) as k: winreg.DeleteValue(k,ident)
            except FileNotFoundError: pass
        else:
            app={'word':'Word','ppt':'Powerpoint','excel':'Excel'}[file.name.split('.')[1]]
            (Path.home()/f'Library/Containers/com.microsoft.{app}/Data/Documents/wef'/('office-ai-trial-'+file.name)).unlink(missing_ok=True)
    if WIN:
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Run',0,winreg.KEY_SET_VALUE) as k: winreg.DeleteValue(k,'OfficeAITrial')
        except FileNotFoundError: pass
        if STATE.exists():
            thumb=json.loads(STATE.read_text())['thumbprint']
            for store in ('Root','My'):
                if store_has_certificate(store,thumb):
                    run(['certutil.exe','-user','-delstore',store,thumb])
                if store_has_certificate(store,thumb):
                    raise RuntimeError('证书未移除，已保留安装记录，请重试卸载。')
    elif STATE.exists():
        sha=json.loads(STATE.read_text())['sha1']
        keychain=Path.home()/'Library/Keychains/login.keychain-db'
        listed=run(['/usr/bin/security','find-certificate','-a','-Z',keychain]).stdout
        if sha.upper() in listed.upper():
            if CERT.exists() and json.loads(STATE.read_text()).get('trusted'):
                run(['/usr/bin/security','remove-trusted-cert',CERT])
                state=json.loads(STATE.read_text());state['trusted']=False;STATE.write_text(json.dumps(state))
            run(['/usr/bin/security','delete-certificate','-Z',sha,keychain])
        if sha.upper() in run(['/usr/bin/security','find-certificate','-a','-Z',keychain]).stdout.upper():
            raise RuntimeError('证书未移除，已保留安装记录。')
    STATE.unlink(missing_ok=True)
    CERT.unlink(missing_ok=True)
    KEY.unlink(missing_ok=True)
    notice('加载入口和自动启动已移除。请重启 Office。设置与历史仍保留在本机；Mac 用户可将 Trojan AI for Office 应用移到废纸篓。')


if __name__ == '__main__':
    try:
        action=sys.argv[1] if len(sys.argv)>1 else '--install'
        {'--serve':serve,'--install':install,'--uninstall':uninstall,'--stop':stop}[action]()
    except Exception as error:
        if '--serve' in sys.argv: raise
        notice('操作未完成：'+str(error))
        sys.exit(1)
