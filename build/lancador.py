r"""Entrada do executável "Malote Fiscal.exe": abre a tela (malote.app) com os argumentos recebidos.

    Malote Fiscal.exe            com o Outlook desta máquina
    Malote Fiscal.exe --demo     Outlook simulado e clientes fictícios

Sem console: se a tela não abrir, o erro vai para %LOCALAPPDATA%\Malote Fiscal\abertura.log e numa caixa de mensagem.
"""
import logging
import os
import sys
from pathlib import Path

TITULO = "Malote Fiscal"


def avisar(texto: str) -> None:
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, texto, TITULO, 0x10)
    except Exception:  # noqa: BLE001
        print(texto, file=sys.stderr)


def main() -> int:
    sys.dont_write_bytecode = True
    pasta = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / TITULO
    pasta.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=pasta / "abertura.log", level=logging.INFO, encoding="utf-8",
                        format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    try:
        from malote.app import main as abrir
        return abrir(sys.argv[1:])
    except SystemExit as saida:  # --help e afins
        return int(saida.code or 0)
    except Exception as erro:  # noqa: BLE001
        logging.exception("falha ao abrir a tela")
        avisar(f"Não foi possível abrir o {TITULO}:\n\n{erro}\n\nDetalhes em {pasta / 'abertura.log'}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
