# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import (
    collect_submodules,
    collect_data_files,
    collect_dynamic_libs,
)

block_cipher = None

# -----------------------------
# 入口脚本：src/for_lab.py
# 使用绝对路径，避免 CI/本地工作目录不一致
# -----------------------------
ROOT = Path(__file__).resolve().parent
entry_script = str(ROOT / "src" / "for_lab.py")

# -----------------------------
# pywin32 / win32com 常见需要项
# -----------------------------
# 1) 收集 win32com 的所有子模块（避免运行时 ImportError）
hiddenimports = []
hiddenimports += collect_submodules("win32com")
hiddenimports += collect_submodules("win32com.client")

# 2) pythoncom / pywintypes 通常是二进制扩展模块，直接列入隐藏导入
hiddenimports += ["pythoncom", "pywintypes"]

# 3) 收集 pywin32_system32 下的 DLL（非常关键，否则 Word COM 可能异常）
binaries = []
binaries += collect_dynamic_libs("pywin32_system32")

# 4) 有时需要带上 pywin32_system32 的数据文件（一般少量）
datas = []
datas += collect_data_files("pywin32_system32", include_py_files=False)

# 如需额外资源文件可在此添加，例如：
# datas += [(str(ROOT / "assets" / "icon.ico"), "assets")]

a = Analysis(
    [entry_script],
    pathex=[str(ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# -----------------------------
# 生成单文件 EXE（onefile）
# 若你想要文件夹形式（onedir），我也可以给你对应 spec
# -----------------------------
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="app",            # ✅ 不要写 app.exe
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,             # ✅ pywin32 + upx 有时会引发问题，建议先关掉
    console=True,          # ✅ 先开控制台方便看报错；稳定后可改 False
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
