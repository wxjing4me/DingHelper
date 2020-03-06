# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['app.py', 'windows\\excelWin.py', 'windows\\MainWin.py', 'common\\analyse.py', 'common\\draw_map.py', 'common\\excel_action.py', 'common\\tencent_api.py', 'config\\logging_setting.py', 'logs\\error'],
             pathex=['D:\\Anaconda3\\envs\\py36\\Lib\\site-packages', 'I:\\Github\\DingHelper'],
             binaries=[],
             datas=[('logs/error', 'logs'), ('images/favicon.ico', 'images'), ('images/excel_tip.png', 'images'), ('include/pyecharts', 'pyecharts'), ('include/pyecharts-1.6.2.dist-info', 'pyecharts-1.6.2.dist-info')],
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
          name='app',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='images\\favicon.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='app')
