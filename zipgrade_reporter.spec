# -*- mode: python -*-

import sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p)

block_cipher = None

a = Analysis(['zipgrade_reporter.py'],
             pathex=['C:\\Users\\jccooper\\Downloads\\zipgrade-reporter-master'],
             binaries=[],
             datas=[
               (path.join(site_packages,"docx","templates"), "docx/templates")
             ],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          name='ZipGrade Reporter',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
