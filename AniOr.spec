# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 配置文件
用于将 AniOr 打包成 Windows 可执行文件
"""

block_cipher = None

a = Analysis(
    ['anior.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # PyQt5 模块
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        # 第三方库
        'requests',
        # 标准库（PyInstaller 可能遗漏）
        'json',
        're',
        'typing',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 排除不需要的重型模块，减小体积
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'PIL',
        'tkinter',
        'unittest',
        'doctest',
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
    name='AniOr',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 启用 UPX 压缩，减小 EXE 体积
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 程序，不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='docs/icon.ico',  # 应用程序图标
)
