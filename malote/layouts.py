"""Layouts de mensagem: o e-mail de pergunta e as respostas prontas, editáveis e salvos na tela.

A pessoa mexe no texto ali mesmo na interface, guarda layouts e não manda "Bom dia" à tarde. Campos
que a ferramenta preenche na hora do envio:
    {saudacao}  Bom dia / Boa tarde / Boa noite, pela hora
    {mes}       09/2026          {mes_nome}  Setembro          {ano}  2026
    {nome}      primeiro nome do cliente (só nas respostas)
Formato do texto: linha em branco separa parágrafos; linha começando com "- " vira item de lista;
**assim** sai em negrito. A assinatura (imagem do último envio, se houver) entra sozinha no fim.

layouts.json (na pasta de dados):
    {"pergunta": {nome: {"assunto": ..., "texto": ...}}, "resposta": {nome: texto}, "pergunta_atual": nome}
"""
from __future__ import annotations

import html as html_lib
import json
import os
import re
from datetime import datetime
from pathlib import Path

from . import competencia as C

CAMPOS = [("{saudacao}", "Bom dia / Boa tarde / Boa noite (pela hora do envio)"),
          ("{mes}", "mês da pergunta, ex.: 09/2026"), ("{mes_nome}", "ex.: Setembro"), ("{ano}", "ex.: 2026"),
          ("{nome}", "primeiro nome do cliente (respostas)")]
ASSUNTO_PADRAO = "Serviços tomados - {mes_nome} de {ano}"
PADRAO = "Padrão"

# O texto da pergunta quando a conta nunca a mandou. Contém a frase que a leitura usa como marca
# ("tomou (contratou)"), por isso mude-a com cuidado — ver outlook.FRASE_DA_PERGUNTA.
TEXTO_PADRAO = """{saudacao}!

Sua empresa **tomou (contratou)** algum serviço de outra empresa ou de profissional autônomo no mês de {mes}?

- Se sim, por favor responda este e-mail com as notas fiscais de serviço (PDF ou XML), inclusive as de prestadores de outros municípios.
- Se não houve serviço tomado, basta responder "não houve".

Precisamos dessas notas para a apuração do ISS e para a escrituração do mês.

Atenciosamente,"""


def periodo(quando: datetime | None = None) -> str:
    hora = (quando or datetime.now()).hour
    return "Bom dia" if hora < 12 else ("Boa tarde" if hora < 18 else "Boa noite")


def preencher(texto: str, comp: C.Competencia, quando: datetime | None = None, nome: str = "") -> str:
    trocas = {"{saudacao}": periodo(quando), "{mes}": C.texto(comp), "{mes_nome}": C.MESES[comp[1] - 1],
              "{ano}": str(comp[0]), "{nome}": nome}
    for campo, valor in trocas.items():
        texto = texto.replace(campo, valor)
    return texto


def texto_do_html(html: str) -> str:
    """HTML do Outlook -> texto editável (negrito como **x**, lista como "- ")."""
    corpo = re.split(r"<body[^>]*>", html, maxsplit=1, flags=re.I)[-1]
    corpo = re.sub(r"(?is)<(style|head|script)\b.*?</\1>", "", corpo)
    corpo = re.sub(r"(?is)<img[^>]*>", "", corpo)
    corpo = re.sub(r"(?is)<(b|strong)\b[^>]*>(.*?)</\1>", r"**\2**", corpo)
    corpo = re.sub(r"(?is)<li\b[^>]*>", "\n- ", corpo)
    corpo = re.sub(r"(?is)<br\s*/?>", "\n", corpo)
    corpo = re.sub(r"(?is)</(p|div|ul|ol|li|h\d)>", "\n\n", corpo)
    corpo = html_lib.unescape(re.sub(r"<[^>]+>", "", corpo)).replace("\xa0", " ")
    linhas = [l.rstrip() for l in corpo.splitlines()]
    texto = re.sub(r"\n{3,}", "\n\n", "\n".join(linhas)).strip()
    texto = re.sub(r"\n\n(?=- )", "\n", texto)          # itens de lista um embaixo do outro
    texto = re.sub(r"\*\*\s*\*\*", "", texto)
    return texto


def para_layout(texto: str, comp: C.Competencia) -> str:
    """Texto de um envio já feito -> layout com campos ({saudacao}, {mes}...)."""
    texto = re.sub(r"^(Bom dia|Boa tarde|Boa noite)", "{saudacao}", texto, flags=re.I)
    texto = texto.replace(f"{C.MESES[comp[1] - 1]} de {comp[0]}", "{mes_nome} de {ano}")
    return texto.replace(C.texto(comp), "{mes}")


def _inline(linha: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", html_lib.escape(linha))


ESTILO = "margin:0 0 12pt 0;font-family:'Times New Roman',serif;font-size:12pt"


def html(texto: str, assinatura: str = "") -> str:
    """Texto preenchido -> HTML do e-mail, num visual sóbrio (Times New Roman 12), como os escritórios usam."""
    partes, paragrafo, lista = [], [], []

    def fecha():
        if paragrafo:
            partes.append(f'<p class=MsoNormal style="{ESTILO}">' + "<br>".join(paragrafo) + "</p>")
            paragrafo.clear()
        if lista:
            itens = "".join(f'<li style="font-family:\'Times New Roman\',serif;font-size:12pt">{i}</li>' for i in lista)
            partes.append(f'<ul style="margin-top:0;margin-bottom:12pt">{itens}</ul>')
            lista.clear()

    for linha in texto.splitlines():
        if not linha.strip():
            fecha()
        elif re.match(r"^\s*[-•*]\s+", linha):
            if paragrafo:
                fecha()
            lista.append(_inline(re.sub(r"^\s*[-•*]\s+", "", linha)))
        else:
            if lista:
                fecha()
            paragrafo.append(_inline(linha))
    fecha()
    if assinatura:
        partes.append(f"<p class=MsoNormal>{assinatura}</p>")
    return ('<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>'
            '<body lang=PT-BR>' + "".join(partes) + "</body></html>")


class Layouts:
    def __init__(self, pasta: Path):
        self.caminho = Path(pasta) / "layouts.json"
        try:
            dados = json.loads(self.caminho.read_text(encoding="utf-8"))
        except FileNotFoundError:
            dados = {}
        self.pergunta: dict[str, dict] = dados.get("pergunta", {})
        self.resposta: dict[str, str] = dados.get("resposta", {})
        self.pergunta_atual: str = dados.get("pergunta_atual", PADRAO)

    def gravar(self) -> None:
        self.caminho.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.caminho.with_suffix(".tmp")
        tmp.write_text(json.dumps({"pergunta": self.pergunta, "resposta": self.resposta,
                                   "pergunta_atual": self.pergunta_atual}, ensure_ascii=False, indent=1),
                       encoding="utf-8")
        os.replace(tmp, self.caminho)

    def garantir_padrao(self, html_modelo: str, comp_modelo: C.Competencia, texto: str | None = None) -> None:
        """Primeira vez: o layout "Padrão" sai do último e-mail enviado (ou de `texto`, quando não houve envio)."""
        if not self.pergunta:
            self.pergunta[PADRAO] = {"assunto": ASSUNTO_PADRAO,
                                     "texto": texto if texto is not None
                                     else para_layout(texto_do_html(html_modelo), comp_modelo)}
            self.pergunta_atual = PADRAO
            self.gravar()
        if self.pergunta_atual not in self.pergunta:
            self.pergunta_atual = next(iter(self.pergunta))

    def salvar_pergunta(self, nome: str, assunto: str, texto: str) -> None:
        self.pergunta[nome] = {"assunto": assunto, "texto": texto}
        self.pergunta_atual = nome
        self.gravar()

    def excluir_pergunta(self, nome: str) -> None:
        self.pergunta.pop(nome, None)
        self.pergunta_atual = next(iter(self.pergunta), PADRAO)
        self.gravar()

    def salvar_resposta(self, nome: str, texto: str) -> None:
        self.resposta[nome] = texto
        self.gravar()

    def excluir_resposta(self, nome: str) -> None:
        self.resposta.pop(nome, None)
        self.gravar()
