# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:/OneDrive - Simpson Strong-Tie (PROD)/All/_GitRepository/pa-form-2023/executive.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/vudinh/.conda/envs/SST_PA_form/Lib/site-packages/smartsheet', 'smartsheet/'), ('E:/OneDrive - Simpson Strong-Tie (PROD)/All/_GitRepository/pa-form-2023/script', 'script/')],
    hiddenimports=['smartsheet.models', 'smartsheet.sheets', 'smartsheet.search', 'smartsheet.attachments'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PA Form Automation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['E:\\OneDrive - Simpson Strong-Tie (PROD)\\All\\_GitRepository\\pa-form-2023\\SSTlogo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PA Form Automation',
)
