# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

added_files = [
    ('*.py', '.'),
    ('./consts/*', 'consts'),
    ('./txt_processors/*.py', 'txt_processors'),
    ('./utils/*.py', 'utils'),
]
added_files += collect_data_files("tabula")

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

a.datas += [
    ('poc_tool8_template.xlsm', './templates/poc_tool8_v1.0.xlsm', 'DATA'),
    ('jre-8u211-windows-x64.tar.gz', './binaries/jre-8u211-windows-x64.tar.gz', 'DATA'),
]


pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='poc_tool8',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir='D:\\temp\\',
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
