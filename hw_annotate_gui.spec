# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# pymupdf（排除 mupdf-devel 瘦身）
tmp = collect_all('pymupdf')
datas += [(s, d) for s, d in tmp[0] if 'mupdf-devel' not in s]
binaries += [(s, d) for s, d in tmp[1] if 'mupdf-devel' not in s]
hiddenimports += tmp[2]

# tkinterdnd2（tkdnd 拖拽库，仅保留 Windows x64，剔除 linux/osx/win-x86/arm64 约 100 个无用文件）
tmp2 = collect_all('tkinterdnd2')
datas += [(s, d) for s, d in tmp2[0] if 'win-x64' in s]
binaries += [(s, d) for s, d in tmp2[1] if 'win-x64' in s]
hiddenimports += tmp2[2]

a = Analysis(
    ['hw_annotate_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'PIL'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='hw_annotate_gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # GUI 不弹控制台
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
    name='hw_annotate_gui',
)
