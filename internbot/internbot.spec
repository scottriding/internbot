# -*- mode: python -*-

block_cipher = None


a = Analysis(['internbot.py'],
             pathex=['/Users/tristanbowler/Documents/GitHub/internbot/internbot'],
             binaries=[('/System/Library/Frameworks/Tk.framework/Tk', 'tk'), ('/System/Library/Frameworks/Tcl.framework/Tcl', 'tcl')],
             datas=[('templates_images/', 'templates_images' )],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
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
          console=False,
          icon='templates_images/y2.icns')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='internbot')
app = BUNDLE(coll,
             name='internbot.app',
             icon='templates_images/y2.icns',
             bundle_identifier=None,
             version=1.0.0)
