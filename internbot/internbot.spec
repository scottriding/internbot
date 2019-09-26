# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['/users/y2analytics/Documents/GitHub/internbot/internbot/internbot.py'],
             pathex=['/Users/y2analytics/Documents/GitHub/internbot/internbot/'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['_tkinter', 'Tkinter', 'enchant', 'twisted'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='internbot',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe, Tree('/users/y2analytics/Documents/GitHub/internbot/internbot/'),
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='internbot')
app = BUNDLE(coll,
             name='internbot.app',
             icon='/users/y2analytics/Documents/GitHub/internbot/internbot/internbot.ico',
             bundle_identifier=None)
