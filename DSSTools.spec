# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

_repo = Path(__file__).resolve().parent
# Run: python tools/ensure_dss_tools_ico.py [--force]  (DSS-Tools Icon.png, lone .ico, or tools/default_dss_tools.ico)
_icon = _repo / "dss_tools.ico"
datas = [('dss_app_version.txt', '.')]
if _icon.is_file():
    datas.append((str(_icon), '.'))
binaries = []
hiddenimports = ["dss_tools_updater"]
tmp_ret = collect_all('pywin32')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PySide6')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

a = Analysis(
    ['dss_hours_tracker.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

_exe_kw = dict(
    name='DSSTools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
if _icon.is_file():
    _exe_kw['icon'] = str(_icon)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, **_exe_kw)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DSSTools',
)

_updater_datas: list = []
if _icon.is_file():
    _updater_datas.append((str(_icon), '.'))

updater_a = Analysis(
    ['dss_tools_updater.py'],
    pathex=[],
    binaries=[],
    datas=_updater_datas,
    hiddenimports=['tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.scrolledtext'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
updater_pyz = PYZ(updater_a.pure)
_updater_exe_kw = dict(
    name='DSSToolsUpdater',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    uac_admin=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
if _icon.is_file():
    _updater_exe_kw['icon'] = str(_icon)
updater_exe = EXE(updater_pyz, updater_a.scripts, updater_a.binaries, updater_a.datas, [], **_updater_exe_kw)
