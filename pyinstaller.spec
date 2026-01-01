# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

# 把下面的入口改成你的实际脚本路径：
# 如果你保留空格路径：
entry_script = r'src/for_lab.py'
# 更推荐的改名后写法（无空格）：
# entry_script = 'src/for_lab.py'

datas = [
    # 如需打包资源文件在此添加：('assets/icon.ico', '.'),
]
hiddenimports = [
    # 某些库需要隐藏导入，可在此补充
]

a = Analysis(
    [entry_script],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='app.exe',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # 若是 GUI 程序改为 False
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,     # 可设置 .ico 图标
)
``
