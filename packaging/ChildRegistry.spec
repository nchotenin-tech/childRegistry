# Run from the repository root: python -m PyInstaller packaging/ChildRegistry.spec
from pathlib import Path
root=Path(SPECPATH).parent
analysis=Analysis([str(root/'run.py')],pathex=[str(root)],binaries=[],datas=[(str(root/'templateRegistry.xlsx'),'.'),(str(root/'USER_MANUAL_TH.html'),'.'),(str(root/'DASHBOARD_MANUAL_TH.html'),'.')],hiddenimports=[],hookspath=[],hooksconfig={},runtime_hooks=[],excludes=['pytest'],noarchive=False)
# Qt uses Windows ICU; a Poppler ICU on PATH exports incompatible symbols.
analysis.binaries = [entry for entry in analysis.binaries if not Path(entry[0]).name.lower().startswith(('icuuc', 'icudt'))]
pyz=PYZ(analysis.pure)
exe=EXE(pyz,analysis.scripts,[],exclude_binaries=True,name='ChildRegistry',debug=False,bootloader_ignore_signals=False,strip=False,upx=False,console=False)
coll=COLLECT(exe,analysis.binaries,analysis.datas,strip=False,upx=False,name='ChildRegistry')
