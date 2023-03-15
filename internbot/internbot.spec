# -*- mode: python ; coding: utf-8 -*-


block_cipher = None

from PyInstaller.utils.hooks import collect_submodules

hidden_imports_html = collect_submodules('html')
hidden_imports_json = collect_submodules('json')
hidden_imports_openpyxl = collect_submodules('openpyxl')
hidden_imports_docx = collect_submodules('docx')
hidden_imports_pptx = collect_submodules('pptx')
hidden_imports_filedialog = collect_submodules('tkinter')
all_hidden_imports = hidden_imports_html + hidden_imports_json + hidden_imports_openpyxl + hidden_imports_docx + hidden_imports_pptx + hidden_imports_filedialog

a = Analysis(
    ['internbot.py'],
    pathex=[],
    binaries=[],
    datas=[('resources', 'resources'), ('model', 'model'), ('view', 'view')],
    hiddenimports=all_hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='internbot',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='internbot',
)
app = BUNDLE(
    coll,
    name='internbot.app',
    icon='/Users/kathryn/Documents/GitHub/internbot/internbot/resources/images/y2.icns',
    bundle_identifier=None,
)

