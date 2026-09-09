# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller specification for PDF2text Desktop Application.
Bundles app.pyw into a standalone Windows executable without a console window.
Zero-bloat build utilizing RapidOCR (ONNX Runtime) and PyMuPDF.
"""

import warnings
warnings.filterwarnings("ignore")

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

# ---------------------------------------------------------------------------
# Data Assets & Models
# ---------------------------------------------------------------------------
datas = [
    ('dictionary.json', '.'),
    ('models', 'models'),
]
binaries = []
hiddenimports = [
    'rapidocr_onnxruntime',
    'onnxruntime',
    'pymupdf',
    'fitz',
    'openpyxl',
    'docx',
    'cv2',
    'numpy',
    'PIL',
    'PIL.Image',
    'rapidfuzz',
    'customtkinter',
    'multiprocessing',
    'concurrent.futures',
]

# Collect essential packages
for pkg in ('rapidocr_onnxruntime', 'onnxruntime', 'customtkinter', 'pymupdf'):
    p_datas, p_binaries, p_hidden = collect_all(pkg, on_error='ignore')
    datas += p_datas
    binaries += p_binaries
    hiddenimports += p_hidden

hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('cv2')
hiddenimports += collect_submodules('rapidfuzz')

a = Analysis(
    ['app.pyw'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'torch',
        'torchvision',
        'easyocr',
        'surya',
        'tensorboard',
        'torch.utils.tensorboard',
        'torch.distributed',
        'torch.testing',
        'caffe2',
        'scipy',
        'matplotlib',
        'pytest',
        'IPython',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PDF2text',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Suppresses command prompt window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PDF2text',
)
