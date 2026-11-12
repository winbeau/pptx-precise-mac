#!/usr/bin/env python3
"""Build a self-contained app with the current Python architecture."""
import argparse
import importlib.metadata
from pathlib import Path
import plistlib
import platform
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parent

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dist', type=Path, required=True)
    parser.add_argument('--work', type=Path, required=True)
    args = parser.parse_args()
    dist, work = args.dist.resolve(), args.work.resolve()
    package = dist / 'PPTX-Precise-Mac'
    if package.exists():
        raise SystemExit('Choose an empty build output directory: ' + str(package))
    package.mkdir(parents=True)
    work.mkdir(parents=True, exist_ok=True)
    subprocess.run([sys.executable,'-m','PyInstaller','--noconfirm','--clean','--onedir',
        '--name','precise-pptx','--distpath',str(work/'binary'), '--workpath',str(work/'pyinstaller'),
        '--specpath',str(work),'--add-data',str(ROOT/'powerpoint_export.applescript')+':.',
        str(ROOT/'precise_pptx.py')], check=True)
    app = package/'PPTX Precise.app'
    (app/'Contents/MacOS').mkdir(parents=True)
    subprocess.run(['/usr/bin/swiftc','-swift-version','5','-O','-target',platform.machine()+'-apple-macos11.0',str(ROOT/'App.swift'),'-o',str(app/'Contents/MacOS/PPTXPrecise')],check=True)
    resources=app/'Contents/Resources'
    resources.mkdir()
    shutil.copy2(ROOT/'assets/AppIcon.icns',resources/'AppIcon.icns')
    shutil.copytree(work/'binary/precise-pptx',resources/'bin/precise-pptx',symlinks=True)
    tool=resources/'tool'; tool.mkdir()
    shutil.copy2(ROOT/'run.command',tool/'run.command')
    for name in ['precise_pptx.py','powerpoint_export.applescript','requirements.txt','run.command',
                 'App.swift','build_macos.py','README.md','LICENSE','CHANGELOG.md','安装依赖.command','test_precise.py','test_export.py']:
        shutil.copy2(ROOT/name,package/name)
    shutil.copytree(ROOT/'assets', package/'assets')
    notices=package/'licenses'; notices.mkdir()
    for name in ['pypdf','Pillow','pyinstaller']:
        d=importlib.metadata.distribution(name)
        for item in d.files or []:
            if 'license' in str(item).lower() or 'copying' in str(item).lower():
                path=Path(d.locate_file(item))
                if path.is_file(): shutil.copy2(path,notices/(name+'-'+path.name))
    for path in [Path(sys.base_prefix)/'LICENSE.txt',Path(sys.base_prefix)/('lib/python'+str(sys.version_info.major)+'.'+str(sys.version_info.minor)+'/LICENSE.txt')]:
        if path.is_file(): shutil.copy2(path,notices/'Python-LICENSE.txt');break
    info=app/'Contents/Info.plist'
    p={'CFBundleExecutable':'PPTXPrecise','CFBundleName':'PPTX Precise','CFBundlePackageType':'APPL','CFBundleIconFile':'AppIcon','LSMinimumSystemVersion':'11.0','NSHighResolutionCapable':True}
    for key in list(p):
        if key.startswith('NS') and key.endswith('UsageDescription'):
            del p[key]
    p.update(CFBundleIdentifier='local.pptxprecise.mac',CFBundleShortVersionString='1.1.0',CFBundleVersion='110',
        NSAppleEventsUsageDescription='本地调用 Microsoft PowerPoint 导出 PDF，不上传文件。')
    p['CFBundleDocumentTypes']=[{'CFBundleTypeName':'PowerPoint presentation or folder','CFBundleTypeRole':'Viewer',
        'LSItemContentTypes':['org.openxmlformats.presentationml.presentation','public.folder']}]
    info.write_bytes(plistlib.dumps(p))
    subprocess.run(['/usr/bin/codesign','--force','--deep','--sign','-',str(app)],check=True)
    subprocess.run(['/usr/bin/codesign','--verify','--deep','--strict',str(app)],check=True)
    print(package)

if __name__=='__main__':main()
