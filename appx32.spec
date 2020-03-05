# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['app.py', 'windows\\excelWin.py', 'windows\\MainWin.py', 'functions\\analyse.py', 'functions\\draw_map.py', 'functions\\excel_action.py', 'functions\\tencent_api.py', 'logs\\error.log'],
             pathex=['D:\\Anaconda3\\envs\\py36x32\\Lib\\site-packages', 'I:\\Github\\DingHelper'],
             binaries=[],
             datas=[('logs/error.log', 'logs'), ('images/favicon.ico', 'images'), ('images/excel_tip.png', 'images'), ('includex32/pyecharts', 'pyecharts'), ('includex32/pyecharts-1.7.0.dist-info', 'pyecharts-1.7.0.dist-info')],
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
