"""Malote Fiscal — abre a tela.

    python malote.pyw            com o Outlook desta máquina
    python malote.pyw --demo     com o Outlook simulado e clientes fictícios (não precisa de Outlook)

Dois cliques no Windows (pythonw) também funcionam; os erros de abertura vão para a pasta de dados.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from malote.app import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
