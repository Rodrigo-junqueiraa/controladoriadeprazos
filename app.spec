# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('Planilha de prazos - atualizada.xlsx', '.'), ('justica.png', '.'), ('icon_inicio.png', '.'), ('icon_calc.png', '.'), ('icon_notification.png', '.'), ('icon_listagem.png', '.'), ('icon_config.png', '.'), ('DejaVuSans.ttf', '.'), ('DejaVuSans-Bold.ttf', '.'), ('DejaVuSans-Oblique.ttf', '.')],
    hiddenimports=[],
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
    name='app',
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
    icon=['icone_juridico.ico'],
)
