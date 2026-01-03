
# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import (
    collect_submodules,
    collect_data_files,
    collect_dynamic_libs,
)

block_cipher = None

# -----------------------------
# 关键修复：不要用 __file__
# PyInstaller 执行 spec 时不保证 __file__ 存在
# 优先用 SPECPATH（PyInstaller 常用的 spec 路径变量），否则用当前工作目录
# -----------------------------
ROOT = Path(globals().get("SPECPATH", Path.cwd())).resolve()

# -----------------------------
# 入口脚本：src/for_lab.py
# 使用绝对路径，避免 CI/本地工作目录不一致
# -----------------------------
entry_script = str(ROOT / "src" / "for_lab.py")

# -----------------------------
# pywin32 / win32com 常见需要项
# -----------------------------
hiddenimports = []
hiddenimports += collect_submodules("win32com")
hiddenimports += collect_submodules("win32com.client")
hiddenimports += ["pythoncom", "pywintypes"]

binaries = []
binaries += collect_dynamic_libs("pywin32_system32")

datas = []
datas += collect_data_files("pywin32_system32", include_py_files=False)

# -----------------------------
# （可选但强烈建议）如果你要用 PDF 签名：补齐 PyMuPDF
# 你之前遇到 fitz not installed，就是没把 pymupdf/fitz 收进去
# -----------------------------
try:
    hiddenimports += collect_submodules("pymupdf")
    hiddenimports += collect_submodules("fitz")
    binaries += collect_dynamic_libs("pymupdf")
    binaries += collect_dynamic_libs("fitz")
    datas += collect_data_files("pymupdf", include_py_files=False)
    datas += collect_data_files("fitz", include_py_files=False)
except Exception as e:
    print(f"[WARN] PyMuPDF collection skipped: {e}")

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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="app",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
