# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['rail_car.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Yashn.jain\\AppData\\Roaming\\Python\\Python38\\site-packages\\customtkinter', 'customtkinter'), ('C:\\Users\\Yashn.jain\\AppData\\Roaming\\Python\\Python38\\site-packages\\bu_snowflake\\rsa_key.p8', './bu_snowflake'), ('C:\\rail car\\Car_type_Mapping', 'Car_type_Mapping'), ('C:\\rail car\\database_old', 'database_old'), ('C:\\rail car\\final_report', 'final_report'), ('C:\\rail car\\customProfile', 'customProfile'), ('biourjaLogo.png', '.')],
    hiddenimports=['snowflake', 'snowflake-connector-python', 'webdriver_manager.firefox', 'sharepy', 'tkcalendar', 'babel.numbers', 'xlwings', 'pandas', 'bu_config', 'bu_alerts'],
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
    name='rail_car',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['biourjaLogo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='rail_car',
)
