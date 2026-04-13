# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PowerUp Portal desktop.

Cross-platform: runs on both Windows and macOS. Produces a single-file
bundle on Windows (.exe) and a .app bundle on macOS.

Usage:
    cd desktop
    pyinstaller PowerUp-Portal.spec --clean --noconfirm
"""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Resolve paths relative to this spec file, not wherever PyInstaller is invoked.
SPEC_DIR = os.path.abspath(os.path.dirname(SPEC) if hasattr(sys.modules[__name__], 'SPEC') else '.')
DESKTOP_DIR = SPEC_DIR
REPO_ROOT = os.path.dirname(DESKTOP_DIR)

block_cipher = None

# ── Files to bundle alongside the Python code ────────────────
datas = [
    # Credentials + icons + env → resources/ inside the bundle
    (os.path.join(DESKTOP_DIR, 'resources'),     'resources'),
    # Whole portal/ folder → portal/ inside the bundle. app_config.py adds
    # this to sys.path at runtime so `from m2_engine import ...` works.
    (os.path.join(REPO_ROOT,   'portal'),        'portal'),
    # Shim folder that shadows portal/google_auth.py with a streamlit-free
    # version. app_config.py inserts this BEFORE portal/ on sys.path.
    (os.path.join(DESKTOP_DIR, 'portal_shims'),  'portal_shims'),
]

# matplotlib ships data files PyInstaller won't auto-detect
datas += collect_data_files('matplotlib')

# ── Hidden imports (modules PyInstaller's static analyser misses) ─
# - portal/* are imported dynamically via sys.path, not via `import portal.x`
# - googleapiclient.discovery_cache.* are loaded by string name at build()
# - customtkinter ships a big tree of widget subclasses
hiddenimports = []
hiddenimports += collect_submodules('customtkinter')
hiddenimports += collect_submodules('googleapiclient')
hiddenimports += collect_submodules('google')
hiddenimports += [
    # portal/ flat modules (imported by bare name once portal/ is on sys.path)
    'm2_engine', 'm3_engine', 'agreement_engine', 'sheets', 'google_auth',
    'config',         # portal/config.py — flat module, not our app_config
    'app_config',     # desktop/app_config.py — bootstraps env + sys.path
    # pandas / numpy optional engines
    'openpyxl', 'openpyxl.cell._writer', 'openpyxl.styles',
    # matplotlib backends
    'matplotlib.backends.backend_agg',
    # python-pptx internals
    'pptx.oxml.shapes.graphfrm',
]

# ── Analysis ──────────────────────────────────────────────────
a = Analysis(
    [os.path.join(DESKTOP_DIR, 'main.py')],
    pathex=[DESKTOP_DIR, os.path.join(REPO_ROOT, 'portal')],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # streamlit is only used by portal/google_auth.py (which we shadow via
    # portal_shims/google_auth.py). Excluding it shaves ~50 MB off the bundle
    # AND avoids the "No package metadata was found for streamlit" crash
    # that happens when PyInstaller includes streamlit but drops its
    # metadata. Related deps (altair, pyarrow, pydeck, tornado, watchdog)
    # come in transitively via streamlit — drop them too.
    excludes=[
        'tests', 'pytest',
        'streamlit', 'altair', 'pyarrow', 'pydeck',
        'tornado', 'watchdog', 'blinker', 'click',
        'gitpython', 'rich',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Single-file executable ────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PowerUp-Portal',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,             # windowed app (no terminal popup)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                 # swap to resources/icon.ico / icon.icns when you have one
)

# ── macOS app bundle ──────────────────────────────────────────
# On macOS, a .app bundle is conventional; on Windows this block is ignored.
app = BUNDLE(
    exe,
    name='PowerUp-Portal.app',
    icon=None,
    bundle_identifier='money.powerup.portal',
    info_plist={
        'CFBundleName': 'PowerUp Portal',
        'CFBundleDisplayName': 'PowerUp Portal',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0',
        'NSHighResolutionCapable': True,
        # Tells macOS this is a regular app (not a background agent),
        # so it shows up in the Dock and the App Switcher.
        'LSApplicationCategoryType': 'public.app-category.productivity',
    },
)
