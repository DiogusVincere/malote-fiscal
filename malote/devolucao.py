"""Aviso de "não entregue": quais endereços falharam, por quê e se vale mandar de novo.

Formatos que os servidores de e-mail costumam mandar:
- hospedagem compartilhada, "Delivery report":  Final-Recipient: rfc822;contato@empresa.example
                               Status: 4.4.4 (unable to route: dns lookup failure)
- Postfix, "Undelivered Mail Returned to Sender":  <fiscal@empresa.example>: host ... said: 454 4.7.1
  ... Relay access denied
- Outlook.com, "Não é possível entregar: <assunto>":  endereço numa linha e o motivo em português,
  depois "Cabeçalhos de mensagem originais". Ex.: caixa cheia (554 5.2.2) e Gmail recusando por
  falta de SPF/DKIM (550 5.7.26).
Nenhum deles traz "tomados" no assunto, por isso a devolução é procurada pelo remetente/assunto e
ligada ao envio pelos endereços que foram perguntados.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_ASSUNTO = re.compile(r"(não é possível entregar|nao e possivel entregar|undeliverable|undelivered|"
                      r"delivery (status|report|failure|has failed)|returned mail|falha na entrega|"
                      r"mail delivery|devolvid|n[aã]o entregue)", re.I)
_REMETENTE = re.compile(r"postmaster|mailer-daemon|mail delivery subsystem", re.I)
_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
_FIM_DO_AVISO = re.compile(r"(Cabeçalhos de mensagem originais|Original message headers|"
                           r"^-+\s*Original message|^Received:|^Return-Path:)", re.I | re.M)

# (padrão, motivo, vale reenviar) — na ordem: o primeiro que bater decide
MOTIVOS = [
    (r"5\.1\.(1|2|3|10)\b|user unknown|does not exist|n[aã]o existe|no such (user|domain|recipient)|"
     r"recipient (address )?rejected|address rejected|unable to route|dns lookup failure|recipient not found|"
     r"bad destination|mailbox unavailable|mailbox not found|bad-mailbox|bad-domain|invalid recipient",
     "endereço ou domínio não existe", False),
    (r"5\.2\.2\b|mailbox full|quota|caixa de correio .{0,40}cheia|over quota",
     "caixa do cliente cheia", True),
    (r"5\.7\.\d+|spam|blocked|bloquead|unauthenticated|\bSPF\b|DKIM|policy|rejeitou",
     "recusado pelo servidor do cliente (spam/autenticação)", True),
    (r"\b4\.\d\.\d+\b|relay access denied|try again|tente reenviar|temporar",
     "falha temporária", True),
    # aviso resumido de hospedagem compartilhada: o mesmo endereço volta assim todo mês
    (r"will not continue trying|permanent(ly)? fail|falha permanente",
     "falha definitiva (o servidor desistiu)", False),
]


@dataclass
class Falha:
    email: str
    motivo: str
    reenviar: bool


def e_devolucao(remetente: str, assunto: str, classe: str) -> bool:
    return (classe.upper().startswith("REPORT.") or bool(_REMETENTE.search(remetente or ""))
            or bool(_ASSUNTO.search(assunto or "")))


def motivo(texto: str) -> tuple[str, bool]:
    for padrao, descricao, reenviar in MOTIVOS:
        if re.search(padrao, texto, re.I):
            return descricao, reenviar
    return "não entregue (motivo não reconhecido)", True


def ler(corpo: str, ignorar: set[str]) -> list[Falha]:
    """Endereços que falharam. `ignorar`: o próprio remetente e afins."""
    corpo = corpo or ""
    aviso = _FIM_DO_AVISO.split(corpo, maxsplit=1)[0] if _FIM_DO_AVISO.search(corpo) else corpo
    enderecos: list[str] = []
    for m in re.finditer(r"Final-Recipient:\s*rfc822;\s*(\S+)", corpo, re.I):
        enderecos.append(m.group(1))
    for m in re.finditer(r"^\s*<([^>\s]+@[^>\s]+)>:", corpo, re.M):
        enderecos.append(m.group(1))
    if not enderecos:
        enderecos = _EMAIL.findall(aviso)
    vistos, falhas = set(), []
    razao, reenviar = motivo(corpo)
    for e in enderecos:
        e = e.strip("<>;,. ").lower()
        if e in vistos or e in ignorar or _REMETENTE.search(e) or not _EMAIL.fullmatch(e):
            continue
        vistos.add(e)
        falhas.append(Falha(e, razao, reenviar))
    return falhas
