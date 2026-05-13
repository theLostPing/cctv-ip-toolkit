# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cctv_toolkit.py'],
    pathex=[],
    binaries=[],
    # app.ico = runtime taskbar icon (loaded via iconbitmap from sys._MEIPASS)
    # logo.ico = file/EXE icon (set via icon=[...] below)
    # Both must be in datas so PyInstaller bundles them; runtime code at line ~5764
    # explicitly looks for app.ico — without it, Tk falls back to its feather icon.
    datas=[('app.ico', '.'), ('logo.ico', '.')],
    # v5.1.0b7 — scapy + zeroconf needed for Switch Loading L2 discovery.
    # PyInstaller can't auto-detect scapy submodules; explicit hidden imports.
    hiddenimports=[
        'scapy.all',
        'scapy.arch.windows',
        'scapy.layers.l2',
        'scapy.layers.inet',
        'scapy.layers.dhcp',
        'zeroconf',
    ],
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
    a.binaries,
    a.datas,
    [],
    name='CCTVIPToolkit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=True,
    icon=['logo.ico'],
)
