# -*- mode: python ; coding: utf-8 -*-
# macOS 专用打包配置：生成原生 填表工具.app（双击即用）
# 用法（必须在 macOS 上执行）：  pyinstaller 填表工具_mac.spec
import os
from PyInstaller.utils.hooks import collect_data_files

datas = []
datas += collect_data_files('docx')
datas += collect_data_files('pptx')

# 把 ffmpeg 随包（供 ⬓ 视频合成 / ⬔ 去水印免安装使用）
binaries = []
try:
    import imageio_ffmpeg
    _ff = imageio_ffmpeg.get_ffmpeg_exe()
    if _ff and os.path.exists(_ff):
        binaries.append((_ff, 'tools'))
except Exception:
    pass

a = Analysis(
    ['exe_main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[],
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
    name='填表工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='填表工具',
)

app = BUNDLE(
    coll,
    name='填表工具.app',
    icon=None,
    bundle_identifier='com.petshop.xlsxfilltool',
    info_plist={
        'NSHighResolutionCapable': 'True',
        'CFBundleDisplayName': '填表工具',
        'CFBundleShortVersionString': '1.0.0',
    },
)
