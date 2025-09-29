# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = [
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'selenium',
    'selenium.webdriver',
    'selenium.webdriver.common.by',
    'selenium.webdriver.support.ui',
    'selenium.webdriver.support.expected_conditions',
    'selenium.webdriver.chrome.options',
    'selenium.common.exceptions',
    'requests',
]
tmp_ret = collect_all('docx')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Bundle project assets (templates/images) so the exe works offline
datas += [
    ('/home/perucci/Projetos/relatorio/preencher_relatorio_gui.py/template_relatorio.docx', 'assets'),
    ('/home/perucci/Projetos/relatorio/preencher_relatorio_gui.py/RELATORIO_DEFAULT_LV.docx', 'assets'),
    ('/home/perucci/Projetos/relatorio/preencher_relatorio_gui.py/print.png', 'assets'),
]


a = Analysis(
    ['/home/perucci/Projetos/relatorio/preencher_relatorio_gui.py/preencher_relatorio_gui.py'],
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='GeradorRelatorio',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
