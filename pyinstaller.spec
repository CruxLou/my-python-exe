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
# -----------------------------
ROOT = Path(__file__).resolve().parent
entry_script = str(ROOT / "src" / "for_lab.py")

# -----------------------------
# hiddenimports / binaries / datas
# -----------------------------
hiddenimports = []
binaries = []
datas = []

# ===== pywin32 / win32com =====
hiddenimports += collect_submodules("win32com")
hiddenimports += collect_submodules("win32com.client")
hiddenimports += ["pythoncom", "pywintypes"]

binaries += collect_dynamic_libs("pywin32_system32")
datas += collect_data_files("pywin32_system32", include_py_files=False)

# ===== PyMuPDF (pymupdf / fitz) =====
# PyMuPDF 文档建议通过 pip 安装 pymupdf，并注意避免无关 fitz 包干扰。[3](https://github.com/pymupdf/PyMuPDF/issues/1976)[4](https://readthedocs.org/projects/pymupdf/downloads/pdf/latest/)
# PyInstaller 对这类带二进制依赖的库通常需要显式收集 hiddenimports/binaries/datas。[1](https://pymupdf.cn/)[2](https://blog.csdn.net/weixin_54537901/article/details/128210875)
try:
    # 新版推荐导入名：pymupdf，但很多代码仍用 fitz（别名/兼容）。
    hiddenimports += collect_submodules("pymupdf")
    hiddenimports += collect_submodules("fitz")

    # 收集 PyMuPDF 的动态库（Windows 下是 .pyd/.dll）
    binaries += collect_dynamic_libs("pymupdf")
    binaries += collect_dynamic_libs("fitz")

    # 收集必要的数据文件（若没有也无妨）
    datas += collect_data_files("pymupdf", include_py_files=False)
    datas += collect_data_files("fitz", include_py_files=False)

except Exception as e:
    # 如果 CI 没装 pymupdf，这里会触发异常——但你应该在 workflow 里先安装（见下文）
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
