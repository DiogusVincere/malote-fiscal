# -*- mode: python ; coding: utf-8 -*-
r"""Receita do PyInstaller para o pacote portátil do Malote Fiscal.

Rode por build\empacotar.py, que chama
    python -m PyInstaller --noconfirm --clean --distpath <dist> --workpath <trabalho> build\malote.spec

Sai "Malote Fiscal.exe" (sem console) + _internal\ com o Python, o tkinter, o pacote malote e as bibliotecas
(pywin32 para o Outlook, rapidfuzz, openpyxl, pdfminer.six).
"""
import os
from pathlib import Path

RAIZ = Path(SPECPATH).parent  # SPECPATH é a pasta deste arquivo (build\)
NOME = "Malote Fiscal"
VERSAO = os.environ.get("MALOTE_VERSAO", "1.0.0")


def gerar_icone(destino: Path):
    """Ícone do executável: um envelope com a aba em azul-ardósia, nas cores do tema."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None
    tam = 256
    im = Image.new("RGBA", (tam, tam), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((8, 8, tam - 8, tam - 8), radius=52, fill="#1F2A37")
    d.rounded_rectangle((40, 76, 216, 192), radius=10, fill="#F4F5F7")
    d.polygon([(40, 86), (128, 150), (216, 86), (216, 100), (128, 164), (40, 100)], fill="#2F5D8A")
    d.rounded_rectangle((150, 60, 232, 112), radius=12, fill="#9DB8D4")
    d.rounded_rectangle((164, 80, 218, 92), radius=5, fill="#1F2A37")
    destino.parent.mkdir(parents=True, exist_ok=True)
    im.save(destino, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return str(destino)


def info_de_versao():
    from PyInstaller.utils.win32.versioninfo import (FixedFileInfo, StringFileInfo, StringStruct, StringTable,
                                                     VarFileInfo, VarStruct, VSVersionInfo)
    partes = tuple(int(p) for p in (VERSAO.split(".") + ["0", "0", "0"])[:3]) + (0,)
    return VSVersionInfo(
        ffi=FixedFileInfo(filevers=partes, prodvers=partes, mask=0x3F, flags=0x0, OS=0x40004,
                          fileType=0x1, subtype=0x0, date=(0, 0)),
        kids=[StringFileInfo([StringTable("041604B0", [
            StringStruct("FileDescription", "Malote Fiscal — serviços tomados por e-mail"),
            StringStruct("ProductName", NOME),
            StringStruct("FileVersion", VERSAO),
            StringStruct("ProductVersion", VERSAO),
            StringStruct("OriginalFilename", NOME + ".exe"),
            StringStruct("LegalCopyright", "Licença MIT"),
        ])]), VarFileInfo([VarStruct("Translation", [1046, 1200])])])


a = Analysis(
    [str(RAIZ / "build" / "lancador.py")],
    pathex=[str(RAIZ)],
    binaries=[],
    datas=[],
    hiddenimports=["malote", "malote.app", "malote.anexos", "malote.cadastro", "malote.clientes", "malote.competencia",
                   "malote.controle", "malote.demo", "malote.devolucao", "malote.falso", "malote.layouts",
                   "malote.outlook", "malote.planilha", "malote.registro", "malote.sugestao", "malote.tema",
                   "tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox", "tkinter.simpledialog",
                   "win32com", "win32com.client", "pythoncom", "pywintypes", "win32timezone",
                   "rapidfuzz", "rapidfuzz.fuzz", "rapidfuzz.process", "rapidfuzz.distance",
                   "openpyxl", "openpyxl.styles", "pdfminer.high_level"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["numpy", "pandas", "matplotlib", "scipy", "IPython", "jupyter", "streamlit", "pdfplumber", "xlrd",
              "lxml", "pytest", "test", "tkinter.test", "idlelib", "turtle", "turtledemo", "lib2to3", "pydoc_data",
              "PIL", "pywebview", "webview"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=NOME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon=gerar_icone(Path(workpath) / "icone.ico"),
    version=info_de_versao(),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name=NOME,
)
