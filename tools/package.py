"""Build installers from the clean payload and verified official runtimes."""
from pathlib import Path
import hashlib
import os
import plistlib
import shutil
import subprocess

BASE=Path(__file__).resolve().parents[1]
STAGE=BASE/'staging'
PAYLOAD=BASE/'app'
CACHE=BASE/'build-cache'
OUT=BASE/'releases'
OUT.mkdir(exist_ok=True)
def run(args):
    return subprocess.run([str(a) for a in args],check=True,capture_output=True,text=True).stdout
def check(name,sha):
    assert hashlib.sha256((CACHE/name).read_bytes()).hexdigest()==sha, name+' checksum mismatch'
check('python-mac.pkg','3b7eaf7f29825f796e8267024435540ddf1f17fc9a97ad58095daa7a75bfdcd3')
# macOS: relocate the universal2 framework into the application bundle.
app=STAGE/'mac-root/Applications/Trojan AI for Office.app'
resources=app/'Contents/Resources'
shutil.copytree(PAYLOAD,resources/'payload',ignore=shutil.ignore_patterns('__pycache__','*.pyc','.DS_Store','.git','shared-settings*.json','installation.json','control-token'))
src=CACHE/'python-mac-expanded/Python_Framework.pkg/Payload/Versions/3.13'
runtime=resources/'runtime'
runtime.mkdir()
shutil.copy2(src/'Python',runtime/'Python')
shutil.copytree(src/'Resources',runtime/'Resources',symlinks=True)
shutil.copytree(src/'lib',runtime/'lib',symlinks=True,ignore=shutil.ignore_patterns('__pycache__','test','tests','pkgconfig'))
(runtime/'bin').mkdir()
shutil.copy2(src/'bin/python3.13',runtime/'bin/python3.13')
license_source=src/'lib/python3.13/LICENSE.txt'
if license_source.exists(): shutil.copy2(license_source,resources/'payload/licenses/Python-LICENSE.txt')
prefix='/Library/Frameworks/Python.framework/Versions/3.13/'
machs=[]
for p in runtime.rglob('*'):
    if p.is_symlink():
        target=os.readlink(p)
        if target.startswith(prefix):
            p.unlink();p.symlink_to(os.path.relpath(runtime/target[len(prefix):],p.parent))
        continue
    if not p.is_file(): continue
    with p.open('rb') as stream: magic=stream.read(4)
    if magic not in (b'\xca\xfe\xba\xbe',b'\xcf\xfa\xed\xfe',b'\xfe\xed\xfa\xcf'): continue
    machs.append(p)
    dependencies=run(['otool','-L',p])
    for line in set(dependencies.splitlines()):
        dep=line.strip().split(' (')[0]
        if dep.startswith(prefix):
            dest=runtime/dep[len(prefix):]
            new='@loader_path/'+os.path.relpath(dest,p.parent)
            run(['install_name_tool','-change',dep,new,p])
    if p.name=='Python': run(['install_name_tool','-id','@loader_path/Python',p])
    run(['codesign','--force','--sign','-',p])
(app/'Contents/MacOS').mkdir()
launcher=app/'Contents/MacOS/OfficeAITrial'
launcher.write_text('#!/bin/sh\nBASE_DIR="$(CDPATH= cd -- "$(dirname -- "$0")/../Resources" && pwd)"\nexport PYTHONHOME="$BASE_DIR/runtime"\nexport SSL_CERT_FILE="$BASE_DIR/payload/cacert.pem"\nexec "$BASE_DIR/runtime/bin/python3.13" -B "$BASE_DIR/payload/trial-launcher.py" --install\n')
launcher.chmod(0o755)
(app/'Contents/Info.plist').write_bytes(plistlib.dumps({'CFBundleExecutable':'OfficeAITrial','CFBundleIdentifier':'local.office-ai-trial.app','CFBundleName':'Trojan AI for Office','CFBundlePackageType':'APPL','CFBundleShortVersionString':'0.1.4','LSMinimumSystemVersion':'11.0','LSUIElement':True}))
unapp=app.parent/'卸载 Trojan AI for Office.app'
(unapp/'Contents/MacOS').mkdir(parents=True)
unlauncher=unapp/'Contents/MacOS/Uninstall'
unlauncher.write_text('#!/bin/sh\nBASE_DIR="$(CDPATH= cd -- "$(dirname -- "$0")/../../../Trojan AI for Office.app/Contents/Resources" && pwd)"\nexport PYTHONHOME="$BASE_DIR/runtime"\nexport SSL_CERT_FILE="$BASE_DIR/payload/cacert.pem"\nexec "$BASE_DIR/runtime/bin/python3.13" -B "$BASE_DIR/payload/trial-launcher.py" --uninstall\n')
unlauncher.chmod(0o755)
(unapp/'Contents/Info.plist').write_bytes(plistlib.dumps({'CFBundleExecutable':'Uninstall','CFBundleIdentifier':'local.office-ai-trial.uninstall','CFBundleName':'卸载 Trojan AI for Office','CFBundlePackageType':'APPL','LSUIElement':True}))
run(['codesign','--force','--deep','--sign','-',app])
run(['codesign','--force','--sign','-',unapp])
env=dict(os.environ,PYTHONHOME=str(runtime),SSL_CERT_FILE=str(resources/'payload/cacert.pem'))
subprocess.run([str(runtime/'bin/python3.13'),'-B','-c','import ssl, json, http.server, urllib.request; print("Bundled Python and TLS imports OK")'],env=env,check=True)
components=STAGE/'mac-components.plist'
run(['pkgbuild','--analyze','--root',STAGE/'mac-root',components])
component_rules=plistlib.loads(components.read_bytes())
for component in component_rules:
    component['BundleIsRelocatable']=False
components.write_bytes(plistlib.dumps(component_rules))
run(['pkgbuild','--root',STAGE/'mac-root','--component-plist',components,'--identifier','local.office-ai-trial.installer','--version','0.1.4','--install-location','/',OUT/'Trojan-AI-for-Office-0.1.4-mac.pkg'])
for p in OUT.iterdir():
    if p.suffix == '.pkg':
        print(p.name,p.stat().st_size,hashlib.sha256(p.read_bytes()).hexdigest())
