# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置文件
打包命令: pyinstaller grading_system.spec
"""

import os
import sys

block_cipher = None

# 项目根目录
project_dir = os.path.dirname(os.path.abspath(__file__))

a = Analysis(
    ['app.py'],
    pathex=[project_dir],
    binaries=[],
    datas=[
        (os.path.join(project_dir, 'templates'), 'templates'),
        (os.path.join(project_dir, 'static'), 'static'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell._writer',
        'openpyxl.styles',
        'docx',
        'docx.opc',
        'requests',
        'flask',
        'werkzeug',
        'jinja2',
        'markupsafe',
        'itsdangerous',
        'click',
        'blinker',
        'urllib3',
        'chardet',
        'idna',
        'certifi',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'test',
        'unittest',
        'PIL',
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
    ],
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
    name='作业批改系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 保留控制台窗口，显示服务器日志
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
