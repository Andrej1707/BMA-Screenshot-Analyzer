# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all, collect_submodules, copy_metadata


def keep_module(name: str) -> bool:
    return not (
        name.startswith("paddle.tensorrt")
        or name.startswith("paddlex.inference.serving")
        or ".tests" in name
    )


datas = []
binaries = []
hiddenimports = []


def add_pkg(pkg: str):
    global datas, binaries, hiddenimports
    try:
        package_datas, package_binaries, package_hiddenimports = collect_all(pkg)
        datas += package_datas
        binaries += package_binaries
        hiddenimports += [name for name in package_hiddenimports if keep_module(name)]
        hiddenimports += [name for name in collect_submodules(pkg) if keep_module(name)]
    except Exception:
        pass


for package in ["paddle", "paddleocr", "paddlex", "pypdfium2", "bidi"]:
    add_pkg(package)

hiddenimports += [
    "cv2",
    "paddlex.inference.pipelines.ocr",
    "paddlex.inference.pipelines.ocr.pipeline",
    "paddlex.inference.pipelines._parallel",
    "paddlex.inference.pipelines.doc_preprocessor",
    "paddlex.inference.pipelines.doc_preprocessor.pipeline",
]
hiddenimports = sorted(set(hiddenimports))

for distribution in [
    "paddlex",
    "paddleocr",
    "paddlepaddle",
    "numpy",
    "Pillow",
    "opencv-python",
    "opencv-contrib-python",
    "opencv_python",
    "pyclipper",
    "Shapely",
    "pypdfium2",
    "python-bidi",
]:
    try:
        datas += copy_metadata(distribution)
    except Exception:
        pass

a = Analysis(
    ["puddle.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "paddle.tensorrt",
        "paddle.tensorrt.*",
        "paddlex.inference.serving",
        "paddlex.inference.serving.*",
        "numpy.f2py.tests",
        "numpy.f2py.tests.*",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="BMA Screenshot Analyzer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon="assets/app.ico",
    version="version_info.txt",
    contents_directory="Dependencies",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="BMA Screenshot Analyzer",
)
