# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['app.py', 'windows\\excelWin.py', 'windows\\MainWin.py', 'common\\analyse.py', 'common\\draw_map.py', 'common\\excel_action.py', 'common\\tencent_api.py', 'configure\\logging_setting.py', 'logs\\error'],
             pathex=['D:\\Anaconda3\\envs\\py36x32\\Lib\\site-packages', 'I:\\Github\\DingHelper'],
             binaries=[],
             datas=[('logs/error', 'logs'), ('images/favicon.ico', 'images'), ('images/excel_tip.png', 'images'), ('D:/Anaconda3/envs/py36-x32/Lib/site-packages/pyecharts', 'pyecharts'), ('D:/Anaconda3/envs/py36-x32/Lib/site-packages/pyecharts-1.7.0.dist-info', 'pyecharts-1.7.0.dist-info')],
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
