# -*- mode: python ; coding: utf-8 -*-
# filepath: excel_diff.spec
"""
PyInstaller spec file for Excel Diff Application
Build: pyinstaller excel_diff.spec
"""

block_cipher = None

a = Analysis(
    ['excel_diff_app/__main__.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.comments',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelDiff',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 모드
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 아이콘 파일 경로로 변경 가능
)