# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['app_msdocx_templates.py'],
             pathex=['C:\\Users\\lulwh\\UG_work\\DocxTemplates'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

for d in a.datas:
    if 'pyconfig' in d[0]:
        a.datas.remove(d)
        break
a.datas += [('logo_ug.png','C:\\Users\\baidak\\acad_py_scripts\\app_msdocx_templates\\logo_ug.png', 'Data')]

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='app_msdocx_templates',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
