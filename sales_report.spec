# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = []
datas += collect_data_files('customtkinter')

# AIインサイト機能（任意）: anthropic "ライブラリ"（APIを呼ぶコード）のみ同梱する。
# APIキーは同梱しない。キーは各ユーザーが実行時にGUIで入力し、各自のPCの
# config.json に保存される（config.json はこのビルドには一切含まれない）。
# anthropic 未導入でもビルドは通る（その場合 exe ではAI機能が無効になるだけ）。
hiddenimports = []
try:
    import anthropic  # noqa: F401
    hiddenimports += collect_submodules('anthropic')
except ImportError:
    pass


a = Analysis(
    ['app\\gui_launcher.pyw'],
    pathex=[],
    binaries=[],
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='sales_report',
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
