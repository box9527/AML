# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
block_cipher = None

added_files = [
    ('*.py', '.'),
    ('./consts/*', 'consts'),
    ('./txt_processors/*.py', 'txt_processors'),
    ('./utils/*.py', 'utils'),
    ('./templates/poc_tool8*.xlsm', '.'),
    ('./extra_excels/*.xlsx', 'extra_excels'),
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
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

a.datas += [
    ('jre-8u211-windows-x64.tar.gz', './binaries/jre-8u211-windows-x64.tar.gz', 'DATA'),
]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='poc_tool8',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='poc_tool8'
)
