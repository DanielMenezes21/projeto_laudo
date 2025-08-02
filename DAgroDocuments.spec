# -*- mode: python ; coding: utf-8 -*-

import sys
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app', 'app'),
        ('assets', 'assets'),
        ('models', 'models'),
        ('modules', 'modules'),
        ('token.json', '.'),
        ('venv/Lib/site-packages/kivy/data', 'data'),
        ('venv/Lib/site-packages/kivymd/fonts', 'kivymd/fonts'),
    ],
    hiddenimports=['kivymd.icon_definitions'],
    hookspath=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DAgroDocuments_v2',
    icon='models/icon/DAgro_Logo_03.ico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # Altere para False se quiser --noconsole
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DAgroDocuments_v2',
    distpath='C:/Temp/output'  # Caminho de saída customizado
)
