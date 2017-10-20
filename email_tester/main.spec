# -*- mode: python -*-
import sys

block_cipher = None


a = Analysis(['main.py'],
             pathex=['C:\\Users\\ics-user\\PycharmProjects\\email_tester'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)



exe = EXE(pyz,
          a.scripts,
		a.binaries + [('msvcp100.dll', 'C:\\Windows\\System32\\msvcp100.dll', 'BINARY'),
					  ('msvcr100.dll', 'C:\\Windows\\System32\\msvcr100.dll', 'BINARY')]
		if sys.platform == 'win32' else a.binaries,
          a.zipfiles,
          a.datas,
          name='main',
          debug=False,
          strip=False,
          upx=True,
          console=True )
