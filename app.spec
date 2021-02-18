# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis([
            'app.py', 
            'windows\\MainWin.py', 
            'windows\\excelWin.py', 
            'windows\\SettingWin.py', 
            'common\\analyse.py', 
            'common\\draw_map.py', 
            'common\\excel_action.py', 
            'common\\test_api_key.py', 
            'configure\\config_action.py', 
            'configure\\config_values.py', 
            'configure\\logging_action.py'
            ],
            pathex=[
                'E:\\Github\\DingHelper\\.env\\Lib\\site-packages\\', 
                'E:\\Github\\DingHelper'
                ],
            binaries=[],
            datas=[
                ('images/favicon.ico', 'images'), 
                ('images/excel_tip.png', 'images'), 
                ('settings/settings_default.json', 'settings'), 
                ('logs/error.log', 'logs'), 
                ('E:\Github\DingHelper\.env\Lib\site-packages\pyecharts', 'pyecharts'), 
                ('E:\Github\DingHelper\.env\Lib\site-packages\pyecharts-1.9.0.dist-info', 'pyecharts-1.9.0.dist-info')
                ],
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
