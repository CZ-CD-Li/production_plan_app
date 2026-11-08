# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：目录模式，启动快，依赖齐全。"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files

# Windows 上 PyInstaller 的隔离分析子进程会先导入 PyQt5，再导入 onnxruntime，
# 这个顺序必定 0xC0000005 崩溃（程序里也因此在入口处先导入 onnxruntime 规避）。
# 这里让隔离进程跳过 onnxruntime / rapidocr 的导入：它们的 DLL 由 hook-onnxruntime
# 收集、模型由 collect_data_files 收集，不依赖这一步。
import PyInstaller.building.build_main as _build_main

_ORIGINAL_FIND_BINARY_DEPENDENCIES = _build_main.find_binary_dependencies
_SKIP_IMPORT = ("onnxruntime", "rapidocr_onnxruntime")


def _find_binary_dependencies(binaries, import_packages, symlink_suppression_patterns):
    packages = [
        name for name in import_packages if name.split(".")[0] not in _SKIP_IMPORT
    ]
    return _ORIGINAL_FIND_BINARY_DEPENDENCIES(
        binaries, packages, symlink_suppression_patterns
    )


_build_main.find_binary_dependencies = _find_binary_dependencies

SPEC_DIR = Path(SPECPATH).resolve()
ROOT = SPEC_DIR.parent
APP_NAME = "生产计划协同管理平台"
TEMPLATE = ROOT / "data" / "templates" / "控制计划格式样板-3.xlsx"
WI_TEMPLATE = ROOT / "data" / "templates" / "C9标准作业指导书模板.xlsx"
GREASE_WI_TEMPLATE = ROOT / "data" / "templates" / "涂润滑脂标准作业指导书.xlsx"
SPC_TEMPLATE = ROOT / "SPCsheet示例.xlsx"
SPC_XMR_TEMPLATE = ROOT / "SPC单值极差图模板.xlsx"
SPC_WULING_TEMPLATE = ROOT / "五菱SPC模板.xlsx"

# 表格识图：RapidOCR 的模型/配置需随包携带。
# 不用 collect_submodules：PyInstaller 会在隔离进程里逐个导入子模块扫描依赖，
# 触发 onnxruntime/cv2 的 DLL 冲突直接崩溃；rapidocr 的子模块由静态分析带入。
OCR_DATAS = collect_data_files(
    "rapidocr_onnxruntime", includes=["**/*.yaml", "**/*.onnx", "**/*.txt", "**/*.json"]
)
OCR_HIDDEN = ["rapidocr_onnxruntime"]

EXCLUDES = [
    "tkinter",
    "unittest",
    "pydoc",
    "doctest",
    "pandas",
    "matplotlib",
    "scipy",
    "setuptools",
    "pip",
    "pytest",
    "PyQt5.QtWebEngine",
    "PyQt5.QtWebEngineCore",
    "PyQt5.QtWebEngineWidgets",
    "PyQt5.QtWebSockets",
    "PyQt5.QtWebChannel",
    "PyQt5.QtQuick",
    "PyQt5.QtQuickWidgets",
    "PyQt5.QtQml",
    "PyQt5.QtMultimedia",
    "PyQt5.QtMultimediaWidgets",
    "PyQt5.QtBluetooth",
    "PyQt5.QtNfc",
    "PyQt5.QtPositioning",
    "PyQt5.QtLocation",
    "PyQt5.QtSerialPort",
    "PyQt5.QtSql",
    "PyQt5.QtTest",
    "PyQt5.QtXml",
    "PyQt5.QtXmlPatterns",
    "PyQt5.QtDesigner",
    "PyQt5.QtHelp",
    "PyQt5.QtOpenGL",
    "PyQt5.QtPrintSupport",
]

a = Analysis(
    [str(ROOT / "main.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=OCR_DATAS
    + ([(str(TEMPLATE), "data/templates")] if TEMPLATE.is_file() else [])
    + ([(str(WI_TEMPLATE), "data/templates")] if WI_TEMPLATE.is_file() else [])
    + (
        [(str(GREASE_WI_TEMPLATE), "data/templates")]
        if GREASE_WI_TEMPLATE.is_file()
        else []
    )
    + ([(str(SPC_TEMPLATE), ".")] if SPC_TEMPLATE.is_file() else [])
    + ([(str(SPC_XMR_TEMPLATE), ".")] if SPC_XMR_TEMPLATE.is_file() else [])
    + (
        [(str(SPC_WULING_TEMPLATE), ".")]
        if SPC_WULING_TEMPLATE.is_file()
        else []
    ),
    hiddenimports=OCR_HIDDEN + ["onnxruntime", "cv2", "pymupdf"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

# PyQt5 的 Qt5/bin 目录自带一套 VS2015 运行时（14.26），Windows 会优先加载它，
# 与 onnxruntime（需要 14.4x）冲突并在导入时 0xC0000005 崩溃。
# 去掉这几份旧副本，统一使用 _internal 根目录下随包收集的系统运行时。
_STALE_MSVC = {
    "msvcp140.dll",
    "msvcp140_1.dll",
    "vcruntime140.dll",
    "vcruntime140_1.dll",
}
a.binaries = [
    entry
    for entry in a.binaries
    if not (
        entry[0].replace("\\", "/").startswith("PyQt5/Qt5/bin/")
        and entry[0].replace("\\", "/").rsplit("/", 1)[-1].lower() in _STALE_MSVC
    )
]

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
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
    icon=str(SPEC_DIR / "app.ico"),
    version=str(SPEC_DIR / "version_info.txt"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name=APP_NAME,
)
