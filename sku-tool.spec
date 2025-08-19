# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['SkuTool Revamped Backup.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['flask', 'werkzeug', 'click', 'itsdangerous', 'blinker', 'jinja2', 'google', 'googleapiclient', 'googleapiclient.discovery', 'googleapiclient.http', 'googleapiclient.errors', 'google_auth_oauthlib', 'google_auth_oauthlib.flow', 'google.auth', 'google.auth.transport.requests', 'httplib2', 'uritemplate'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='sku-tool',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='sku-tool',
)
