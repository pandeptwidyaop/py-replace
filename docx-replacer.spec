# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file untuk DOCX Placeholder Replacer
Build single executable yang bisa didistribusikan tanpa Python
"""

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Get customtkinter path
import customtkinter
ctk_path = Path(customtkinter.__file__).parent

# Get python-docx paths
import docx
docx_pkg_path = Path(docx.__file__).parent
docx_templates_path = docx_pkg_path / 'templates'

# Manually add docx templates with correct structure
# This ensures templates are placed relative to docx/parts as expected
docx_datas = [
    (str(docx_templates_path), 'docx/templates'),  # Templates dir
]

a = Analysis(
    ['src/main.py'],
    pathex=['src'],  # Add src to path
    binaries=[],
    datas=[
        (str(ctk_path), 'customtkinter'),  # Include CustomTkinter files
        ('src/gui', 'gui'),  # Include gui module
        ('src/utils', 'utils'),  # Include utils module
    ] + docx_datas,  # Include python-docx template files
    hiddenimports=[
        'customtkinter',
        'PIL._tkinter_finder',
        'docx',
        'lxml',
        'pandas',
        'openpyxl',
        'darkdetect',
        'gui',
        'gui.app',
        'utils',
        'utils.docx_handler',
        'utils.placeholder',
        'utils.config_loader',
        'utils.image_handler',
        'urllib',
        'urllib.request',
    ],
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
    exclude_binaries=True,  # Changed to onedir mode
    name='DOCX-Replacer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window (GUI only)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon file here if you have one
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DOCX-Replacer',
)

# For macOS, create .app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='DOCX-Replacer.app',
        icon=None,
        bundle_identifier='com.docxreplacer.app',
        info_plist={
            'NSHighResolutionCapable': 'True',
            'LSBackgroundOnly': 'False',
        },
    )
