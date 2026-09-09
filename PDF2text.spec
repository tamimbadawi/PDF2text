# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller specification for PDF2text Desktop Application.
Bundles app.pyw into a standalone Windows executable without a console window.
"""

import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

# ---------------------------------------------------------------------------
# Bootstrap: collect_all ensures every DLL, data file, and sub-module for the
# heaviest dependencies is captured — prevents "missing library" crashes in the
# frozen EXE.
# ---------------------------------------------------------------------------
datas = [('dictionary.json', '.')]
binaries = []
hiddenimports = ['torch', 'torchvision', 'torchvision.ops', 'torchvision.models', 'easyocr', 'rapidfuzz']

def torch_submodule_filter(name):
    # Exclude unused modules that trigger missing dependency warnings or deprecations
    for excluded in ('tensorboard', 'distributed', 'testing', 'caffe2'):
        if excluded in name:
            return False
    return True

p_datas, p_binaries, p_hidden = collect_all('torch', filter_submodules=torch_submodule_filter, on_error='ignore')
datas += p_datas
binaries += p_binaries
hiddenimports += p_hidden

for pkg in ('torchvision', 'easyocr'):
    p_datas, p_binaries, p_hidden = collect_all(pkg, on_error='ignore')
    datas     += p_datas
    binaries  += p_binaries
    hiddenimports += p_hidden

# ---------------------------------------------------------------------------
# Additional data assets
# ---------------------------------------------------------------------------
datas += collect_data_files('customtkinter')
datas += collect_data_files('pymupdf')
datas += collect_data_files('surya')
datas += [('models', 'models')]

# ---------------------------------------------------------------------------
# Additional hidden imports
# ---------------------------------------------------------------------------
hiddenimports += [
    'customtkinter',
    'pymupdf',
    'fitz',
    'openpyxl',
    'docx',
    'cv2',
    'numpy',
    'PIL',
    'PIL.Image',
    'surya',
    'surya.layout',
    'surya.table_rec',
    'transformers',
    'safetensors',
    'pypdfium2',
    'dotenv',
    'platformdirs',
    'multiprocessing',
    'concurrent.futures',
]
hiddenimports += collect_submodules('pymupdf')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('cv2')
hiddenimports += collect_submodules('surya')

a = Analysis(
    ['app.pyw'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=['hooks'],
    hooksconfig={},
    excludes=[
        'tensorboard',
        'torch.utils.tensorboard',
        'torch.distributed',
        'torch.testing',
        'caffe2',
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
    console=False,  # Suppresses black console command prompt
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
