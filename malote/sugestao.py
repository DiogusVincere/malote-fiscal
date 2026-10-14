"""Sugestão de resposta ao cliente, com aprendizado supervisionado e nada saindo da máquina.

Sem IA de fora: toda resposta passa pelo olho de quem opera; há sugestão para todas as respostas (inclusive "segue anexo" e "não houve").

Como decide, em ordem:
1. exemplo ensinado parecido (o usuário aprovou/ensinou uma resposta para mensagem semelhante);
2. regras para o que é comum (pedido para sair da lista, troca de e-mail, pergunta, anexo,
   link, "não houve");
3. resposta curta de agradecimento.
Toda resposta aceita vira exemplo novo (a menos que o usuário desmarque "aprender").

A parecença é TF-IDF por radical de palavra (5 primeiras letras, sem acento e sem palavras
vazias): "Despesas com pedágio e gasolina, entra?" encontra "pedágio entra como serviço?".

O formato da resposta copia o que a equipe já escreve à mão (aprendido de Itens Enviados):
    Bom dia, Maria.
    Espero que você esteja bem.
    <miolo>
    Atenciosamente,
"""
from __future__ import annotations

import json
import math
import os
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .competencia import sem_acento

ABERTURA = "Espero que você esteja bem."
FECHO = "Atenciosamente,"

_VAZIAS = set("""a o as os um uma uns umas de da do das dos e em no na nos nas por para pra com sem
que se ao aos à às ou mas mais muito muita já ja nao não sim boa bom dia tarde noite tudo bem obrigado
obrigada att atenciosamente prezados prezado prezada olá ola oi grato grata favor gentileza abraço abs
este esta isso esse essa aqui ai vc você voce me meu minha nosso nossa eu nós nos ele ela foi ser
está esta estão sao são tem ter teve vou vai the of mes mês""".split())


def _radicais(texto: str) -> list[str]:
    palavras = re.findall(r"[a-z0-9]+", sem_acento(texto or "").lower())
    return [p[:5] for p in palavras if len(p) >= 3 and p not in _VAZIAS and not p.isdigit()]


# ------------------------------------------------------------------ regras
_SAIR = re.compile(r"encerr\w* (o )?contrato|deixamos de ter|outro contador|outra contabilidade|"
                   r"retir\w* (o |meu |nosso )?e-?mail|remov\w* (o |meu |nosso )?e-?mail|descadastr|"
                   r"n[aã]o (somos|sou) mais cliente|n[aã]o (quero|desejo) (mais )?receber", re.I)
_TROCA = re.compile(r"(mudan[cç]a|troca|altera[cç][aã]o|alterar|trocar|mudar|atualizar) (do |o |de )?e-?mail|"
                    r"e-?mail (ser[aá]|vai ser) (descontinuado|desativado)|novo e-?mail|e-?mail novo", re.I)
_PERGUNTA = re.compile(r"\?|\b(posso|pode|podem|seria|vale|entra|entram|devo|preciso|precisa|"
                       r"como fa[cç]o|qual|quais)\b", re.I)

RESPOSTAS_PADRAO = {
    "Sair da lista": "Irei retirar o seu e-mail da lista, obrigado.",
    "Trocar e-mail": "Ok, recebido. Vamos atualizar o e-mail no nosso cadastro.",
    "Mandou anexo": "Ok, recebido, agradeço o retorno.",
    "Mandou link": "Ok, recebido, vou baixar as notas pelo link. Agradeço o retorno.",
    "Sem serviços": "Ok, agradeço o retorno.",
    "Pergunta": "",            # não há o que sugerir sem exemplo: o usuário escreve e ensina
    "Outro": "Ok, agradeço o retorno.",
}


# cortesia com ponto de interrogação não é pergunta ("Tudo bem? Não contratamos.")
_CORTESIA = re.compile(r"(tudo (bem|bom|certo|joia|ok)|e (com )?voc[eê]|como vai|como est[aá]s?|td bem|tudo em ordem)"
                       r"\s*[?!.]*", re.I)
_LIXO_ASSINATURA = re.compile(r"(Yahoo Mail|Obter o Outlook|Get Outlook|Enviado do meu|Sent from my|<https?://)\S*.*$",
                              re.I | re.S)


def categoria(texto: str, tipo: str) -> str:
    if _SAIR.search(texto):
        return "Sair da lista"
    if _TROCA.search(texto):
        return "Trocar e-mail"
    if tipo in ("Mandou anexo", "Mandou link"):
        return tipo
    if _PERGUNTA.search(_LIXO_ASSINATURA.sub("", _CORTESIA.sub(" ", texto))):
        return "Pergunta"
    if tipo == "Sem serviços":
        return tipo
    return "Outro"


# ------------------------------------------------------------------ saudação
_NAO_E_NOME = set("""financeiro fiscal contato atendimento administrativo administracao comercial compras
contabil contabilidade adm sac rh dp suporte diretoria gerencia recepcao clinica emporio loja escritorio
empresa grupo equipe pet shop studio estudio academia oficina padaria farmacia escola restaurante
transportadora construtora distribuidora consultoria informatica secretaria secretaria agencia auto
casa centro ponto espaco instituto laboratorio hotel pousada bar cafe mercado supermercado posto""".split())


def saudacao(nome_exibicao: str, quando: datetime | None = None) -> str:
    hora = (quando or datetime.now()).hour
    periodo = "Bom dia" if hora < 12 else ("Boa tarde" if hora < 18 else "Boa noite")
    nome = primeiro_nome(nome_exibicao)
    return f"{periodo}, {nome}." if nome else f"{periodo},"


def primeiro_nome(nome_exibicao: str) -> str:
    """'Joaquina Ferreira' -> 'Joaquina'; 'Financeiro | Escritório Modelo', 'ACME EMBALAGENS' -> ''."""
    nome = (nome_exibicao or "").strip().strip("'\"")
    if "@" in nome or not re.fullmatch(r"[A-ZÀ-Ý][a-zà-ÿ]+(?:\s+(?:[A-ZÀ-Ý][a-zà-ÿ.]*|d[aeo]s?|e))*", nome):
        return ""
    primeiro = nome.split()[0]
    return "" if sem_acento(primeiro).lower() in _NAO_E_NOME or len(primeiro) < 3 else primeiro


# ------------------------------------------------------------------ base de exemplos
@dataclass
class Exemplo:
    mensagem: str
    resposta: str
    categoria: str = ""
    origem: str = "ensinado"    # "ensinado" (tela), "aceito" (sugestão aprovada), "histórico" (Itens Enviados)
    criado: str = ""
    usos: int = 0


@dataclass
class Sugestao:
    categoria: str
    resposta: str
    porque: str
    exemplo: Exemplo | None = None
    parecenca: float = 0.0
    acoes: list[str] = field(default_factory=list)


class Base:
    LIMIAR = 0.34            # abaixo disso a mensagem não é "parecida" com o exemplo
    LIMIAR_PERGUNTA = 0.12   # pergunta sem exemplo parecido: ainda mostra o mais próximo, avisando

    def __init__(self, pasta: Path):
        self.caminho = Path(pasta) / "aprendizado.json"
        try:
            dados = json.loads(self.caminho.read_text(encoding="utf-8"))
        except FileNotFoundError:
            dados = {}
        self.exemplos = [Exemplo(**e) for e in dados.get("exemplos", [])]

    def gravar(self) -> None:
        tmp = self.caminho.with_suffix(".tmp")
        tmp.write_text(json.dumps({"exemplos": [e.__dict__ for e in self.exemplos]}, ensure_ascii=False, indent=1),
                       encoding="utf-8")
        os.replace(tmp, self.caminho)

    def ensinar(self, mensagem: str, resposta: str, categoria_: str, origem: str = "ensinado") -> None:
        mensagem, resposta = mensagem.strip(), resposta.strip()
        if not mensagem or not resposta:
            return
        for e in self.exemplos:  # a mesma mensagem ensinada de novo troca a resposta antiga
            if e.mensagem == mensagem:
                e.resposta, e.categoria, e.origem = resposta, categoria_, origem
                self.gravar()
                return
        self.exemplos.append(Exemplo(mensagem, resposta, categoria_, origem,
                                     datetime.now().isoformat(timespec="seconds")))
        self.gravar()

    def esquecer(self, exemplo: Exemplo) -> None:
        self.exemplos = [e for e in self.exemplos if e is not exemplo]
        self.gravar()

    def parecidos(self, texto: str, limite: int = 3) -> list[tuple[float, Exemplo]]:
        consulta = Counter(_radicais(texto))
        if not consulta or not self.exemplos:
            return []
        docs = [Counter(_radicais(e.mensagem)) for e in self.exemplos]
        n = len(docs) + 1
        df = Counter(r for d in docs for r in d)

        def vetor(c: Counter) -> dict[str, float]:
            return {r: (1 + math.log(q)) * math.log((n + 1) / (df.get(r, 0) + 1)) + 0.01 for r, q in c.items()}

        vq = vetor(consulta)
        nq = math.sqrt(sum(v * v for v in vq.values()))
        notas = []
        for e, d in zip(self.exemplos, docs):
            vd = vetor(d)
            nd = math.sqrt(sum(v * v for v in vd.values()))
            if nd and nq:
                notas.append((sum(vq[r] * vd.get(r, 0) for r in vq) / (nq * nd), e))
        notas.sort(key=lambda x: -x[0])
        return notas[:limite]

    def sugerir(self, texto: str, tipo: str) -> Sugestao:
        cat = categoria(texto, tipo)
        acoes = {"Sair da lista": ["Tirar o e-mail da lista de envio"],
                 "Trocar e-mail": ["Trocar o e-mail no contato/grupo do Outlook"]}.get(cat, [])
        achados = self.parecidos(texto)
        if achados and achados[0][0] >= self.LIMIAR:
            nota, ex = achados[0]
            return Sugestao(cat, ex.resposta,
                            f"Parecida ({nota:.0%}) com uma mensagem que você já respondeu: “{_curto(ex.mensagem)}”",
                            ex, nota, acoes)
        padrao = RESPOSTAS_PADRAO[cat]
        if cat == "Pergunta" and achados and achados[0][0] >= self.LIMIAR_PERGUNTA:
            # pergunta nova: melhor um ponto de partida (para corrigir) do que a caixa vazia
            nota, ex = achados[0]
            return Sugestao(cat, ex.resposta, f"Pouco parecida ({nota:.0%}) com: “{_curto(ex.mensagem)}”. "
                                              "Confira e ajuste antes de aceitar.", ex, nota, acoes)
        if cat == "Pergunta":
            porque = "Parece uma pergunta e ainda não há exemplo parecido: escreva a resposta e ela fica aprendida."
        else:
            porque = f"Regra: {cat.lower()}."
        return Sugestao(cat, padrao, porque, None, 0.0, acoes)


def _curto(texto: str, n: int = 90) -> str:
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto if len(texto) <= n else texto[: n - 1] + "…"


# ------------------------------------------------------------------ montar e desmontar
def montar_texto(saudacao_: str, miolo: str) -> str:
    return "\n".join(p for p in (saudacao_.strip(), ABERTURA, miolo.strip(), FECHO) if p)


def miolo_de_resposta_antiga(corpo: str) -> str:
    """Tira saudação, 'Espero que…' e 'Atenciosamente' do que a equipe escreveu à mão."""
    linhas = [l.strip() for l in corpo.splitlines() if l.strip()]
    if linhas and re.match(r"^(bom dia|boa tarde|boa noite|ol[aá]|prezad)", sem_acento(linhas[0]).lower()):
        linhas = linhas[1:]
    linhas = [l for l in linhas if not sem_acento(l).lower().startswith("espero que voce esteja bem")]
    for i, l in enumerate(linhas):
        if sem_acento(l).lower().startswith("atenciosamente"):
            linhas = linhas[:i]
            break
    return "\n".join(linhas)
