"""Competência (mês de referência) do e-mail de serviços tomados.

O e-mail sai no começo do mês perguntando do mês anterior, com o assunto
"Serviços tomados - Agosto de 2026" e o texto "... no mês de 08/2026?".
"""
from __future__ import annotations

import re
import unicodedata
from datetime import date

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
         "Setembro", "Outubro", "Novembro", "Dezembro"]

Competencia = tuple[int, int]  # (ano, mês)


def sem_acento(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


_MESES_SEM_ACENTO = [sem_acento(m).lower() for m in MESES]


def anterior(hoje: date | None = None) -> Competencia:
    hoje = hoje or date.today()
    return (hoje.year - 1, 12) if hoje.month == 1 else (hoje.year, hoje.month - 1)


def proxima(comp: Competencia) -> Competencia:
    ano, mes = comp
    return (ano + 1, 1) if mes == 12 else (ano, mes + 1)


def ultimas(quantas: int = 12, hoje: date | None = None) -> list[Competencia]:
    """As últimas competências, da mais nova para a mais velha, começando no mês anterior."""
    lista, comp = [], anterior(hoje)
    for _ in range(quantas):
        lista.append(comp)
        ano, mes = comp
        comp = (ano - 1, 12) if mes == 1 else (ano, mes - 1)
    return lista


def texto(comp: Competencia) -> str:
    """08/2026"""
    return f"{comp[1]:02d}/{comp[0]}"


def do_texto(valor: str) -> Competencia:
    mes, ano = valor.strip().split("/")
    return int(ano), int(mes)


def assunto(comp: Competencia) -> str:
    return f"Serviços tomados - {MESES[comp[1] - 1]} de {comp[0]}"


_ASSUNTO = re.compile(r"^servicos tomados - (\w+) de (\d{4})$", re.I)


def do_assunto(valor: str) -> Competencia | None:
    """Competência de um assunto exatamente no padrão do envio (não de uma resposta)."""
    m = _ASSUNTO.match(sem_acento(valor or "").strip())
    if not m or m.group(1).lower() not in _MESES_SEM_ACENTO:
        return None
    return int(m.group(2)), _MESES_SEM_ACENTO.index(m.group(1).lower()) + 1


_NOME_MES = re.compile(r"(?<![a-z])(" + "|".join(_MESES_SEM_ACENTO) + r")(?![a-z])"
                       r"(?:\s*(?:de|/|-|\.)?\s*(\d{4}|\d{2})(?!\d))?", re.I)
_MES_ANO = re.compile(r"(?<!\d)(0[1-9]|1[0-2])\s*[/.-]\s*(20\d{2})(?!\d)")


def mencionada(trecho: str, referencia: Competencia) -> Competencia | None:
    """Mês citado num assunto ou nome de arquivo ("Agosto de 2026", "08/2026", ".Setembro - X.zip").

    Mês sem ano fica no ano que deixa a data mais perto da `referencia`.
    Mais de um mês diferente no mesmo texto não decide nada.
    """
    achados: set[Competencia] = set()
    limpo = sem_acento(trecho or "").lower()
    for m in _NOME_MES.finditer(limpo):
        if m.group(1) == "marco" and not m.group(2):
            continue  # "NF 265 MARCO TULIO.pdf" é gente, não mês
        mes = _MESES_SEM_ACENTO.index(m.group(1)) + 1
        if m.group(2):
            ano = int(m.group(2))
            ano = ano + 2000 if ano < 100 else ano
        else:
            ano = min((referencia[0] - 1, referencia[0], referencia[0] + 1),
                      key=lambda a: abs((a * 12 + mes) - (referencia[0] * 12 + referencia[1])))
        achados.add((ano, mes))
    for m in _MES_ANO.finditer(limpo):
        achados.add((int(m.group(2)), int(m.group(1))))
    return achados.pop() if len(achados) == 1 else None
