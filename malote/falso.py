"""O Outlook simulado e o escritório fictício: é o que a demonstração e os testes usam.

Nada aqui é real. O escritório "Contábil Exemplo" tem a conta fiscal@contabil.example, catorze
clientes com nomes inventados, e-mails em domínios `.example` (reservados, nunca entregues) e CNPJs
gerados por semente. As mensagens cobrem os casos que a leitura precisa tratar:

- resposta com nota em PDF, resposta "não houve", resposta com link para baixar;
- o mesmo e-mail servindo a duas empresas, com uma nota (XML) de cada — o CNPJ dentro do anexo decide;
- e-mail de cliente que NÃO responde à pergunta (assunto próprio, ZIP com as notas): antes era invisível;
- cliente respondendo antes de ser perguntado; resposta de outra competência; dúvida em vez de nota;
- remetente desconhecido do mesmo domínio do cliente; remetente só reconhecível pelo nome de exibição;
- resposta na subpasta "Clientes" e resposta no lixo eletrônico; spam no lixo (fica de fora);
- pedido para trocar o e-mail; pedido para sair da lista; e-mail interno da equipe;
- dois avisos de "não entregue" (endereço inexistente × caixa cheia).

`OutlookFalso` tem a mesma interface de `outlook.Outlook` e reaproveita a lógica pura do módulo
(ler_item, selecionar_respostas, gravar_na_pasta), que é a parte que importa testar.
"""
from __future__ import annotations

import io
import json
import re
import shutil
import struct
import tempfile
import zipfile
import zlib
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from . import competencia as C
from . import devolucao
from . import layouts as L
from .outlook import (AnexoLido, Conta, Envio, Grupo, ItemLido, Modelo, Resposta, contexto_de,
                      e_sobre_a_pergunta, gravar_na_pasta, janela_de_leitura, marcas_da_pergunta,
                      par_de_resposta_antiga, selecionar_respostas)

CONTA = "fiscal@contabil.example"
ROTULO = f"{CONTA}  ·  demonstração"      # como a conta simulada aparece na lista da tela
COLEGA = "ana@contabil.example"
NOME_CONTA = "Fiscal · Contábil Exemplo"

# (código, razão social, e-mails, regime municipal, responsável fiscal, ex-cliente)
EMPRESAS = [
    ("101", "PADARIA AURORA LTDA", ["financeiro@padariaaurora.example", "contabil@grupoaurora.example"],
     "Simples Nacional", "Ana", False),
    ("102", "CLINICA VITALIS SERVICOS MEDICOS LTDA",
     ["financeiro@clinicavitalis.example", "dra.helena@clinicavitalis.example"], "Lucro Presumido", "Bruno", False),
    ("103", "TRANSPORTADORA ROTA NORTE LTDA", ["antigo@rotanorte.example", "fiscal@rotanorte.example"], "", "Ana", False),
    ("104", "OFICINA MECANICA DOIS IRMAOS ME", ["oficina@dois-irmaos.example"], "Simples Nacional", "Carla", False),
    ("105", "ESTUDIO PIXEL DESIGN LTDA", ["oi@pixeldesign.example"], "Simples Nacional", "Bruno", False),
    ("106", "CONSTRUTORA HORIZONTE AZUL LTDA", [], "Lucro Presumido", "Carla", False),
    ("107", "ACADEMIA CORPO EM MOVIMENTO LTDA", ["academia@corpoemmovimento.example"], "Simples Nacional", "Ana", False),
    ("108", "RESTAURANTE SABOR DA SERRA LTDA", ["contabil@grupoaurora.example"], "Simples Nacional", "Ana", False),
    ("109", "DISTRIBUIDORA BOA SAFRA LTDA", ["fiscal@boasafra.example", COLEGA], "Lucro Real", "Bruno", False),
    ("110", "ESCOLA DE IDIOMAS FALA MUNDO LTDA", ["secretaria@falamundo.example"], "Simples Nacional", "Carla", False),
    ("111", "PET SHOP AMIGO FIEL ME", ["amigofiel@petshop.example"], "Simples Nacional", "Ana", False),
    ("112", "CONSULTORIA ATLAS GESTAO LTDA", ["adm@atlasgestao.example"], "Lucro Presumido", "Bruno", False),
    ("113", "INFORMATICA BYTE CERTO LTDA", ["ti@bytecerto.example"], "ISENTA MUNICIPAL", "Carla", False),
    ("114", "FARMACIA BEM ESTAR LTDA", ["farmacia@bemestar.example"], "Simples Nacional", "Ana", False),
    ("115", "GRAFICA IMPRESSAO RAPIDA LTDA", ["grafica@impressaorapida.example"], "Simples Nacional", "Bruno", False),
    ("000 EX-CLIENTE", "MOVEIS ANTIGA LTDA", ["moveis@antiga.example"], "Simples Nacional", "Carla", True),
]
SEM_PASTA = {"115", "000 EX-CLIENTE"}           # estão na planilha, mas não têm pasta na rede
PASTAS_EXTRAS = ["116 - LAVANDERIA BRILHO LTDA"]  # tem pasta, mas não está na planilha

GRUPOS = {
    "Clientes Fiscal 1": ["financeiro@padariaaurora.example", "contabil@grupoaurora.example",
                          "financeiro@clinicavitalis.example", "dra.helena@clinicavitalis.example",
                          "antigo@rotanorte.example", "fiscal@rotanorte.example", "oficina@dois-irmaos.example",
                          "oi@pixeldesign.example", "socio@horizonteazul.example", "academia@corpoemmovimento.example"],
    "Clientes Fiscal 2": ["fiscal@boasafra.example", "secretaria@falamundo.example", "amigofiel@petshop.example",
                          "adm@atlasgestao.example", "farmacia@bemestar.example", "grafica@impressaorapida.example",
                          "moveis@antiga.example"],
    "Fornecedores": ["suporte@sistemas.example", "comercial@papelaria.example"],
}
CONTATOS = {
    "financeiro@padariaaurora.example": "Padaria Aurora", "contabil@grupoaurora.example": "Grupo Aurora Contábil",
    "financeiro@clinicavitalis.example": "Clínica Vitalis", "dra.helena@clinicavitalis.example": "Helena Prado",
    "fiscal@rotanorte.example": "Transportadora Rota Norte", "oficina@dois-irmaos.example": "Oficina Dois Irmãos",
    "oi@pixeldesign.example": "Pixel Design", "socio@horizonteazul.example": "Construtora Horizonte Azul",
    "academia@corpoemmovimento.example": "Academia Corpo em Movimento", "fiscal@boasafra.example": "Boa Safra",
    "secretaria@falamundo.example": "Escola Fala Mundo", "amigofiel@petshop.example": "Pet Shop Amigo Fiel",
    "adm@atlasgestao.example": "Atlas Gestão", "farmacia@bemestar.example": "Farmácia Bem Estar",
    "grafica@impressaorapida.example": "Gráfica Impressão Rápida", "moveis@antiga.example": "Móveis Antiga",
}
PRESTADORES = {
    "limpeza": ("LIMPEZA TOTAL SERVICOS LTDA", 901), "seguranca": ("SEGURANCA ALFA VIGILANCIA LTDA", 902),
    "ti": ("TI SOLUCOES EM NUVEM LTDA", 903), "guincho": ("GUINCHO RAPIDO LTDA", 904),
    "personal": ("PERSONAL FIT TREINAMENTOS LTDA", 905), "marketing": ("MARKETING DIGITAL ONDA LTDA", 906),
}


# ------------------------------------------------------------------ dados fictícios
def cnpj_ficticio(semente: int) -> str:
    """CNPJ com dígitos verificadores válidos, derivado de uma semente (só dígitos)."""
    base = [int(d) for d in f"{(semente * 7919 + 1234567) % 100_000_000:08d}0001"]

    def dv(nums):
        pesos = list(range(len(nums) - 7, 1, -1)) + list(range(9, 1, -1))
        resto = sum(n * p for n, p in zip(nums, pesos)) % 11
        return 0 if resto < 2 else 11 - resto

    d1 = dv(base)
    return "".join(map(str, base + [d1, dv(base + [d1])]))


def formatar_cnpj(c: str) -> str:
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"


def cnpj_da_empresa(codigo: str) -> str:
    return cnpj_ficticio(int(re.sub(r"\D", "", codigo) or "0") + 1)


def pdf_simples(linhas: list[str]) -> bytes:
    """Um PDF de uma página com texto em Helvetica, sem biblioteca nenhuma (o pdfminer lê o texto dele)."""
    def esc(t: str) -> str:
        return t.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    fluxo = ("BT /F1 11 Tf 50 790 Td 15 TL " + " ".join(f"({esc(l)}) Tj T*" for l in linhas) + " ET").encode(
        "cp1252", "replace")
    objetos = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> "
        b"/Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        b"<< /Length %d >>\nstream\n" % len(fluxo) + fluxo + b"\nendstream",
    ]
    saida = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    posicoes = []
    for i, obj in enumerate(objetos, start=1):
        posicoes.append(len(saida))
        saida += b"%d 0 obj\n" % i + obj + b"\nendobj\n"
    xref = len(saida)
    saida += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objetos) + 1)
    for p in posicoes:
        saida += b"%010d 00000 n \n" % p
    saida += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objetos) + 1, xref)
    return bytes(saida)


def png_minusculo(largura: int = 160, altura: int = 48, cor=(47, 93, 138)) -> bytes:
    """PNG de uma cor só, para fazer de imagem de assinatura."""
    linha = b"\x00" + bytes(cor) * largura
    dados = zlib.compress(linha * altura)

    def bloco(tipo: bytes, corpo: bytes) -> bytes:
        return struct.pack(">I", len(corpo)) + tipo + corpo + struct.pack(">I", zlib.crc32(tipo + corpo) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + bloco(b"IHDR", struct.pack(">IIBBBBB", largura, altura, 8, 2, 0, 0, 0))
            + bloco(b"IDAT", dados) + bloco(b"IEND", b""))


def nota_pdf(numero: int, prestador: str, tomador_codigo: str, valor: float, descricao: str, comp: C.Competencia) -> bytes:
    nome_p, semente = PRESTADORES[prestador]
    razao = next(e[1] for e in EMPRESAS if e[0] == tomador_codigo)
    return pdf_simples([
        f"NOTA FISCAL DE SERVICOS ELETRONICA - NFS-e N. {numero}",
        f"Data de emissao: 15/{comp[1]:02d}/{comp[0]}",
        "",
        f"PRESTADOR: {nome_p}",
        f"CNPJ: {formatar_cnpj(cnpj_ficticio(semente))}",
        "",
        f"TOMADOR: {razao}",
        f"CNPJ: {formatar_cnpj(cnpj_da_empresa(tomador_codigo))}",
        "",
        f"Discriminacao: {descricao}",
        f"Valor dos servicos: R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        "ISS retido: nao",
        "",
        "Documento ficticio gerado pela demonstracao do Malote Fiscal.",
    ])


def nota_xml(numero: int, prestador: str, tomador_codigo: str, valor: float, descricao: str, comp: C.Competencia) -> bytes:
    nome_p, semente = PRESTADORES[prestador]
    razao = next(e[1] for e in EMPRESAS if e[0] == tomador_codigo)
    return (f'<?xml version="1.0" encoding="UTF-8"?>\n'
            f'<CompNfse xmlns="http://www.abrasf.org.br/nfse.xsd"><Nfse versao="2.04"><InfNfse Id="nfse{numero}">'
            f'<Numero>{numero}</Numero><DataEmissao>{comp[0]}-{comp[1]:02d}-15T10:00:00</DataEmissao>'
            f'<ValoresNfse><ValorLiquidoNfse>{valor:.2f}</ValorLiquidoNfse></ValoresNfse>'
            f'<PrestadorServico><IdentificacaoPrestador><CpfCnpj><Cnpj>{cnpj_ficticio(semente)}</Cnpj></CpfCnpj>'
            f'</IdentificacaoPrestador><RazaoSocial>{nome_p}</RazaoSocial></PrestadorServico>'
            f'<DeclaracaoPrestacaoServico><InfDeclaracaoPrestacaoServico><Servico><Valores>'
            f'<ValorServicos>{valor:.2f}</ValorServicos></Valores><Discriminacao>{descricao}</Discriminacao></Servico>'
            f'<Tomador><IdentificacaoTomador><CpfCnpj><Cnpj>{cnpj_da_empresa(tomador_codigo)}</Cnpj></CpfCnpj>'
            f'</IdentificacaoTomador><RazaoSocial>{razao}</RazaoSocial></Tomador>'
            f'</InfDeclaracaoPrestacaoServico></DeclaracaoPrestacaoServico></InfNfse></Nfse></CompNfse>\n').encode("utf-8")


def zip_de(arquivos: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nome, dados in arquivos.items():
            z.writestr(nome, dados)
    return buf.getvalue()


# ------------------------------------------------------------------ mensagens
@dataclass
class AnexoFalso:
    nome: str
    dados: bytes
    cid: str = ""
    oculto: bool = False


@dataclass
class MensagemFalsa:
    entry_id: str
    recebido: datetime
    email: str
    nome: str
    assunto: str
    corpo: str
    anexos: list[AnexoFalso] = field(default_factory=list)
    classe: str = "IPM.Note"
    pasta: str = "Caixa de entrada"
    respondido: bool = False

    @property
    def e_lixo(self) -> bool:
        return self.pasta.lower().startswith("lixo")

    @property
    def html(self) -> str:
        paragrafos = "".join(f"<p>{l}</p>" for l in self.corpo.splitlines() if l.strip())
        imagens = "".join(f'<img src="cid:{a.cid}">' for a in self.anexos if a.cid)
        return f"<html><body>{paragrafos}{imagens}</body></html>"

    def item(self) -> ItemLido:
        anexos = [AnexoLido(i, a.nome, len(a.dados), a.cid, a.oculto) for i, a in enumerate(self.anexos, start=1)]
        return ItemLido(self.entry_id, self.recebido, self.assunto, self.classe, self.corpo, self.html,
                        self.email, self.nome, anexos, self.respondido, self.pasta)


@dataclass
class EnviadoFalso:
    quando: datetime
    assunto: str
    corpo: str                               # texto
    html: str
    destinatarios: list[tuple[str, str]]     # (nome, e-mail) em Cco
    tipo: str = "pergunta"                   # "pergunta" | "resposta"


def _citacao(nome: str, email: str, quando: datetime, assunto: str, texto_cliente: str) -> str:
    return (f"\n\nDe: {nome} <{email}>\nEnviado: {quando:%A, %d de %B de %Y %H:%M}\nPara: {CONTA}\n"
            f"Assunto: {assunto}\n\n{texto_cliente}\n")


class Demonstracao:
    """O escritório fictício: pastas de clientes, planilha, caixa de entrada e itens enviados."""

    def __init__(self, pasta: str | Path | None = None, hoje: date | None = None):
        self.hoje = hoje or date.today()
        self.comp = C.anterior(self.hoje)                 # a competência que a tela abre
        self.pasta = Path(pasta) if pasta else Path(tempfile.gettempdir()) / "malote-fiscal-demo"
        self.raiz = self.pasta / "Clientes"
        self.dados = self.pasta / "dados"
        self.planilha = self.pasta / f"Clientes {self.hoje:%Y-%m-%d}.xlsx"
        self.caixa: list[MensagemFalsa] = []
        self.enviados: list[EnviadoFalso] = []
        self._n = 0
        self._montar()

    # ---------------------------------------------------------------- montagem
    def _data(self, comp: C.Competencia, dia: int, hora: int, minuto: int = 0, mes_seguinte: bool = True) -> datetime:
        ano, mes = C.proxima(comp) if mes_seguinte else comp
        return datetime(ano, mes, dia, hora, minuto)

    def _msg(self, comp_recebida, dia, hora, minuto, email, nome, assunto, corpo, anexos=(), quando=None,
             **extra) -> MensagemFalsa:
        """Uma mensagem na caixa. Sem `quando`, a data é `dia` do mês SEGUINTE à competência `comp_recebida`."""
        self._n += 1
        m = MensagemFalsa(f"demo-{self._n:03d}", quando or self._data(comp_recebida, dia, hora, minuto), email, nome,
                          assunto, corpo, list(anexos), **extra)
        self.caixa.append(m)
        return m

    def destinatarios_padrao(self) -> list[tuple[str, str]]:
        """Quem recebeu as perguntas anteriores: planilha (menos ex-cliente e interno) + grupos de clientes."""
        emails: list[str] = []
        for _cod, _razao, lista, regime, _resp, ex in EMPRESAS:
            if ex or regime.upper() in ("ISENTA MUNICIPAL", "INEXISTENTE MUNICIPAL"):
                continue
            emails += [m for m in lista if m != COLEGA]
        for nome_grupo, lista in GRUPOS.items():
            if nome_grupo.lower().startswith("clientes"):
                emails += [m for m in lista if m != "moveis@antiga.example"]
        vistos: list[str] = []
        for m in emails:
            if m not in vistos:
                vistos.append(m)
        return [(CONTATOS.get(m, m), m) for m in sorted(vistos)]

    def _pergunta(self, comp: C.Competencia) -> EnviadoFalso:
        quando = self._data(comp, 2, 14, 5)
        texto = L.preencher(L.TEXTO_PADRAO, comp, quando)
        return EnviadoFalso(quando, C.assunto(comp), texto, L.html(texto), self.destinatarios_padrao(), "pergunta")

    def _resposta_equipe(self, m: MensagemFalsa, quando: datetime, saudacao: str, miolo: str) -> None:
        corpo = (f"{saudacao}\nEspero que você esteja bem.\n{miolo}\nAtenciosamente,\n"
                 + _citacao(m.nome, m.email, m.recebido, m.assunto, m.corpo))
        m.respondido = True
        self.enviados.append(EnviadoFalso(quando, "RES: " + m.assunto, corpo, L.html(corpo), [(m.nome, m.email)],
                                          "resposta"))

    def _montar(self) -> None:
        comp = self.comp
        ant = C.anterior(date(comp[0], comp[1], 1))
        ant2 = C.anterior(date(ant[0], ant[1], 1))
        res = lambda c: "RES: " + C.assunto(c)  # noqa: E731
        mes_nome = C.MESES[comp[1] - 1].lower()

        # perguntas já mandadas (a deste mês inclusive: a tela abre com as respostas chegando)
        self.enviados += [self._pergunta(ant2), self._pergunta(ant), self._pergunta(comp)]

        # ---- dois meses atrás (respostas à pergunta de ant2), já respondidas pela equipe
        m = self._msg(ant2, 3, 9, 30, "financeiro@clinicavitalis.example", "Clínica Vitalis", res(ant2),
                      "Bom dia. Não houve serviço tomado no mês.")
        self._resposta_equipe(m, self._data(ant2, 3, 11, 0), "Bom dia,", "Ok, agradeço o retorno.")
        m = self._msg(ant2, 3, 10, 10, "fiscal@boasafra.example", "Boa Safra", res(ant2), "Segue a nota.",
                      [AnexoFalso("NFSe 410 - seguranca.pdf", nota_pdf(410, "seguranca", "109", 2300.0,
                                                                        "Vigilância patrimonial", ant2))])
        self._resposta_equipe(m, self._data(ant2, 3, 11, 5), "Bom dia,", "Ok, recebido, agradeço o retorno.")

        # ---- mês passado (respostas à pergunta de ant)
        m = self._msg(ant, 3, 9, 12, "financeiro@clinicavitalis.example", "Clínica Vitalis", res(ant),
                      "Bom dia. Não tivemos serviços tomados.")
        self._resposta_equipe(m, self._data(ant, 3, 15, 0), "Bom dia,", "Ok, agradeço o retorno.")
        m = self._msg(ant, 3, 9, 40, "financeiro@padariaaurora.example", "Maria Souza", res(ant),
                      "Bom dia! Segue a nota do mês.\nAtt,\nMaria",
                      [AnexoFalso("NFSe 1198 - Limpeza Total.pdf", nota_pdf(1198, "limpeza", "101", 850.0,
                                                                            "Limpeza e conservação", ant))])
        self._resposta_equipe(m, self._data(ant, 3, 15, 5), "Bom dia, Maria.", "Ok, recebido, agradeço o retorno.")
        m = self._msg(ant, 4, 11, 0, "oi@pixeldesign.example", "Pixel Design", res(ant),
                      "Oi! As notas estão neste link: https://we.tl/t-exemplo1\nAbraço")
        self._resposta_equipe(m, self._data(ant, 4, 14, 0), "Boa tarde,",
                              "Ok, recebido, vou baixar as notas pelo link. Agradeço o retorno.")
        m = self._msg(ant, 5, 10, 20, "amigofiel@petshop.example", "Pet Shop Amigo Fiel", res(ant),
                      "Bom dia. Pedágio e estacionamento entram como serviço tomado? Não tivemos outras notas.")
        self._resposta_equipe(m, self._data(ant, 5, 16, 0), "Bom dia,",
                              "Pedágio e estacionamento não entram: são despesas sem nota fiscal de serviço tomado. "
                              "Só precisamos das notas de serviço (NFS-e) emitidas contra a empresa.")
        # resposta de ant que chegou já dentro do mês de comp: fica de fora da leitura de comp
        self._msg(comp, 15, 9, 0, "fiscal@boasafra.example", "Boa Safra", res(ant), "Não houve.",
                  quando=datetime(comp[0], comp[1], 15, 9, 0))

        # ---- este mês: as respostas à pergunta de comp (mandada no dia 2 do mês seguinte)
        self._msg(comp, 2, 14, 6, "", "Mail Delivery System", "Undelivered Mail Returned to Sender",
                  "This is the mail system at host mx.contabil.example.\n\nI'm sorry to have to inform you that your "
                  "message could not be delivered to one or more recipients.\n\n<antigo@rotanorte.example>: host "
                  "mx.rotanorte.example said: 550 5.1.1 <antigo@rotanorte.example>: Recipient address rejected: "
                  "User unknown in virtual mailbox table (in reply to RCPT TO command)\n", classe="REPORT.IPM.Note.NDR")
        self._msg(comp, 2, 14, 7, "postmaster@mail.example", "postmaster@mail.example", "Delivery report",
                  "Delivery report\n\nFinal-Recipient: rfc822;academia@corpoemmovimento.example\nAction: failed\n"
                  "Status: 5.2.2 (mailbox full)\nDiagnostic-Code: smtp; 552 5.2.2 Over quota\n")
        self._msg(comp, 2, 16, 30, "fiscal@boasafra.example", "Boa Safra", res(comp), "Não houve.")
        self._msg(comp, 3, 9, 12, "financeiro@padariaaurora.example", "Maria Souza", res(comp),
                  "Bom dia! Segue a nota do mês.\nAtt,\nMaria",
                  [AnexoFalso("NFSe 1234 - Limpeza Total.pdf", nota_pdf(1234, "limpeza", "101", 850.0,
                                                                        "Limpeza e conservação", comp)),
                   AnexoFalso("image001.png", png_minusculo(), cid="image001.png@01DC0001")])
        self._msg(comp, 3, 10, 40, "financeiro@clinicavitalis.example", "Clínica Vitalis", res(comp),
                  f"Bom dia. Não tivemos serviços tomados em {mes_nome}.")
        self._msg(comp, 3, 12, 0, "corpoemmovimento.academia@provedor.example", "Academia Corpo em Movimento", res(comp),
                  "Segue a nota do personal trainer que presta serviço para nós.",
                  [AnexoFalso("NF 77 personal.pdf", nota_pdf(77, "personal", "107", 1200.0,
                                                             "Treinamento funcional - 20 horas", comp))])
        self._msg(comp, 3, 15, 22, "oi@pixeldesign.example", "Pixel Design", res(comp),
                  "Oi! As notas estão neste link: https://we.tl/t-exemplo2\nAbraço")
        self._msg(comp, 4, 8, 5, "contabil@grupoaurora.example", "Grupo Aurora Contábil", res(comp),
                  "Bom dia. Seguem as notas das duas empresas do grupo.",
                  [AnexoFalso("NFSe 88 - marketing.xml", nota_xml(88, "marketing", "101", 1500.0,
                                                                  "Gestão de redes sociais", comp)),
                   AnexoFalso("NFSe 301 - ti.xml", nota_xml(301, "ti", "108", 640.0,
                                                            "Hospedagem do sistema de pedidos", comp))])
        self._msg(comp, 4, 3, 0, "premio@loteria.example", "Loteria Premiada", "Você ganhou!",
                  "Clique aqui para receber o seu prêmio.", pasta="Lixo eletrônico")
        self._msg(comp, 5, 10, 0, "farmacia@bemestar.example", "Farmácia Bem Estar", res(comp),
                  "Bom dia. Solicito a troca do e-mail farmacia@bemestar.example para financeiro@bemestar.example. "
                  f"Não tivemos notas em {mes_nome}.")
        self._msg(comp, 5, 10, 30, "moveis@antiga.example", "Móveis Antiga", res(comp),
                  "Não somos mais clientes do escritório. Favor retirar nosso e-mail da lista.")
        self._msg(comp, 6, 11, 30, "secretaria@falamundo.example", "Secretaria Fala Mundo", f"Notas de {mes_nome}",
                  f"Boa tarde! Encaminho em anexo as notas dos serviços que recebemos em {mes_nome}.\nObrigada!",
                  [AnexoFalso(f"notas {mes_nome}.zip", zip_de({
                      "NFSe 512 - limpeza.pdf": nota_pdf(512, "limpeza", "110", 980.0, "Limpeza das salas", comp),
                      "NFSe 2044 - ti.pdf": nota_pdf(2044, "ti", "110", 350.0, "Suporte de informática", comp)}))])
        self._msg(comp, 6, 16, 45, "amigofiel@petshop.example", "Pet Shop Amigo Fiel", res(comp),
                  "Boa tarde. Despesas com pedágio e estacionamento entram? Não tivemos outras notas.",
                  pasta="Lixo eletrônico")
        self._msg(comp, 7, 9, 0, "joao@dois-irmaos.example", "João Pereira", "Nota fiscal",
                  "Segue a nota do guincho.",
                  [AnexoFalso("guincho.pdf", nota_pdf(5061, "guincho", "104", 420.0, "Remoção de veículo", comp))])
        self._msg(comp, 7, 14, 10, "adm@atlasgestao.example", "Atlas Gestão", res(comp),
                  "Boa tarde, segue a nota da consultoria de TI.",
                  [AnexoFalso("NFSe 919 - ti.pdf", nota_pdf(919, "ti", "112", 3200.0,
                                                            "Consultoria em infraestrutura", comp))],
                  pasta="Caixa de entrada/Clientes")
        self._msg(comp, 8, 9, 15, COLEGA, "Ana Lima", "ENC: " + C.assunto(comp),
                  "Encaminho a nota que o cliente mandou para o meu e-mail.",
                  [AnexoFalso("NFSe 640 - seguranca.pdf", nota_pdf(640, "seguranca", "109", 2300.0,
                                                                   "Vigilância patrimonial", comp))])
        self._msg(comp, 9, 17, 20, "novidades@ofertas.example", "Ofertas Imperdíveis", "Promoção do mês",
                  "Aproveite os descontos de até 50%.")
        # chegou no dia 1º do mês seguinte, citando a competência anterior: mostrada, em cinza
        self._msg(comp, 1, 8, 30, "fiscal@boasafra.example", "Boa Safra", res(ant),
                  f"Segue a nota de {C.MESES[ant[1] - 1].lower()} que havia ficado faltando.",
                  [AnexoFalso("NFSe 398 - seguranca.pdf", nota_pdf(398, "seguranca", "109", 2300.0,
                                                                   "Vigilância patrimonial", ant))])
        # chegou ANTES de a pergunta ser mandada, dentro do próprio mês da competência
        self._msg(comp, 28, 10, 0, "fiscal@rotanorte.example", "Transportadora Rota Norte", res(comp),
                  "Já adiantando: segue a única nota do mês.",
                  [AnexoFalso("NFSe 2210 - ti.pdf", nota_pdf(2210, "ti", "103", 700.0, "Rastreamento da frota", comp))],
                  quando=datetime(comp[0], comp[1], 28, 10, 0))

    # ---------------------------------------------------------------- disco
    def preparar(self, limpar: bool = True) -> None:
        """Cria (ou recria) a pasta da demonstração: árvore de clientes, planilha e config da ferramenta."""
        if limpar and self.pasta.exists():
            shutil.rmtree(self.pasta, ignore_errors=True)
        self.raiz.mkdir(parents=True, exist_ok=True)
        self.dados.mkdir(parents=True, exist_ok=True)
        ano = str(self.comp[0])
        mm = f"{self.comp[1]:02d}"
        for codigo, razao, *_ in EMPRESAS:
            if codigo in SEM_PASTA:
                continue
            pasta = self.raiz / f"{codigo} - {razao}"
            pasta.mkdir(exist_ok=True)
            if codigo == "101":
                (pasta / "NFE" / ano).mkdir(parents=True, exist_ok=True)
            elif codigo == "102":
                (pasta / "NFe" / ano / f"{mm}.{ano}" / "Serviços tomados").mkdir(parents=True, exist_ok=True)
            elif codigo == "112":  # a nota deste mês já está lá: salvar de novo tem de dizer "já estava"
                destino = pasta / "NF-E" / ano / f"{mm}{ano}" / "Tomados" / "Tomados e-mail"
                destino.mkdir(parents=True, exist_ok=True)
                m = next(x for x in self.caixa if x.email == "adm@atlasgestao.example" and x.anexos)
                (destino / m.anexos[0].nome).write_bytes(m.anexos[0].dados)
        for nome in PASTAS_EXTRAS:
            (self.raiz / nome).mkdir(exist_ok=True)
        self._gravar_planilha()
        config = self.dados / "config.json"
        if not config.exists():
            config.write_text(json.dumps({"conta": CONTA, "raiz_clientes": str(self.raiz), "planilha": str(self.planilha),
                                          "grupos": ["Clientes Fiscal 1", "Clientes Fiscal 2"]},
                                         ensure_ascii=False, indent=1), encoding="utf-8")

    def _gravar_planilha(self) -> None:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Clientes"
        ws.append(["Código", "Razão Social", "CNPJ/CPF", "E-mails", "Regime municipal", "Resp. fiscal", "Ex-cliente"])
        for codigo, razao, emails, regime, resp, ex in EMPRESAS:
            ws.append([codigo, razao, formatar_cnpj(cnpj_da_empresa(codigo)), " | ".join(emails), regime, resp,
                       "Sim" if ex else ""])
        wb.save(self.planilha)

    def mensagem(self, entry_id: str) -> MensagemFalsa:
        return next(m for m in self.caixa if m.entry_id == entry_id)

    def abrir(self) -> "OutlookFalso":
        return OutlookFalso(self)


def contas() -> list[Conta]:
    return [Conta(CONTA, NOME_CONTA, "Demonstração")]


class OutlookFalso:
    """Mesma interface de `outlook.Outlook`, sobre a `Demonstracao`."""

    def __init__(self, demo: Demonstracao):
        self.demo = demo
        self.endereco = CONTA
        self.dominio = CONTA.rsplit("@", 1)[-1]
        self.assinaturas: set[int] = set()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return None

    # ---- o que a tela lê ao abrir
    def modelo(self) -> Modelo:
        ultimo = max((e for e in self.demo.enviados if e.tipo == "pergunta"), key=lambda e: e.quando)
        comp = C.do_assunto(ultimo.assunto) or self.demo.comp
        return Modelo(ultimo.assunto, comp, ultimo.quando, ultimo.html, [], "")

    def envios(self) -> dict[C.Competencia, list[Envio]]:
        achados: dict[C.Competencia, list[Envio]] = {}
        for e in sorted(self.demo.enviados, key=lambda e: e.quando, reverse=True):
            comp = C.do_assunto(e.assunto)
            if comp is not None and e.tipo == "pergunta":
                achados.setdefault(comp, []).append(Envio(comp, e.quando, list(e.destinatarios)))
        return achados

    def grupos(self) -> list[Grupo]:
        return [Grupo(nome, [(CONTATOS.get(m, m), m) for m in lista]) for nome, lista in GRUPOS.items()]

    def nomes_dos_contatos(self) -> dict[str, str]:
        return dict(CONTATOS)

    def respostas_antigas(self) -> list[tuple[str, str]]:
        pares = []
        for e in self.demo.enviados:
            if e.tipo == "resposta":
                par = par_de_resposta_antiga(e.assunto, e.corpo)
                if par:
                    pares.append(par)
        return pares

    # ---- leitura
    def _itens(self, limite: datetime, com_lixo: bool = True):
        for m in sorted(self.demo.caixa, key=lambda m: m.recebido, reverse=True):
            if m.recebido >= limite and (com_lixo or not m.e_lixo):
                yield m.item(), m.e_lixo

    def respostas(self, comp: C.Competencia, envios: dict, perguntados: set[str],
                  conhecidos: set[str] | None = None) -> list[Resposta]:
        ctx = contexto_de(self.endereco, envios, self.assinaturas, perguntados, conhecidos or ())
        inicio_do_mes, _desde = janela_de_leitura(comp, envios)
        return selecionar_respostas(self._itens(inicio_do_mes), comp, envios, ctx)

    def quem_respondeu(self, desde: datetime, envios: dict | None = None) -> dict[str, datetime]:
        marcas = marcas_da_pergunta(self.endereco, envios, C.ultimas(24, self.demo.hoje))
        achados: dict[str, datetime] = {}
        for d, _lixo in self._itens(desde, com_lixo=False):
            if (not d.classe.startswith("IPM.Note") or not d.email or d.email == self.endereco
                    or d.email.endswith("@" + self.dominio) or devolucao.e_devolucao(d.email, d.assunto, d.classe)):
                continue
            if e_sobre_a_pergunta(d.assunto, d.corpo, marcas) and (d.email not in achados or d.recebido > achados[d.email]):
                achados[d.email] = d.recebido
        return achados

    # ---- enviar
    def montar(self, assunto: str, html: str, imagens: list, cco: list[str], enviar: bool, so_testar: bool = False):
        if so_testar:
            return {"conta": self.endereco, "assunto": assunto, "para": self.endereco, "cco": list(cco),
                    "anexos": [(Path(a).name, cid) for a, cid in imagens], "html": html, "salvar_em": "Itens Enviados"}
        if not enviar:
            return "Rascunhos (simulado)"
        self.demo.enviados.append(EnviadoFalso(datetime.now(), assunto, L.texto_do_html(html), html,
                                               [(CONTATOS.get(m, m), m) for m in cco], "pergunta"))
        return "enviado"

    def ja_respondido(self, entry_id: str) -> bool:
        return self.demo.mensagem(entry_id).respondido

    def responder(self, entry_id: str, texto: str, modelo: Modelo, so_testar: bool = False):
        m = self.demo.mensagem(entry_id)
        if so_testar:
            return {"conta": self.endereco, "assunto": "RES: " + m.assunto, "para": m.email, "anexos": [],
                    "texto": texto, "html": L.html(texto)}
        m.respondido = True
        corpo = texto + _citacao(m.nome, m.email, m.recebido, m.assunto, m.corpo)
        self.demo.enviados.append(EnviadoFalso(datetime.now(), "RES: " + m.assunto, corpo, L.html(corpo),
                                               [(m.nome, m.email)], "resposta"))
        return None

    # ---- anexos
    def _anexos(self, entry_id: str, indices: list[int]) -> list[tuple[int, AnexoFalso]]:
        m = self.demo.mensagem(entry_id)
        return [(i, a) for i, a in enumerate(m.anexos, start=1) if i in indices]

    def salvar_anexos(self, entry_id: str, indices: list[int], pasta: Path) -> list[tuple[str, str]]:
        feitos = []
        tmp = Path(tempfile.mkdtemp(prefix="malote_demo_"))
        try:
            for i, a in self._anexos(entry_id, indices):
                arquivo = tmp / f"{i}.bin"
                arquivo.write_bytes(a.dados)
                feitos.append(gravar_na_pasta(arquivo, a.nome, Path(pasta)))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        return feitos

    def cnpjs_dos_anexos(self, entry_id: str, indices: list[int]) -> dict[int, set[str]]:
        from . import anexos
        achados: dict[int, set[str]] = {}
        tmp = Path(tempfile.mkdtemp(prefix="malote_demo_"))
        try:
            for i, a in self._anexos(entry_id, indices):
                arquivo = tmp / f"{i}{Path(a.nome).suffix}"
                arquivo.write_bytes(a.dados)
                achados[i] = anexos.cnpjs(arquivo, a.nome)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        return achados
