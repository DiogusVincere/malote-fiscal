"""Conversa com o Outlook da máquina (COM, via pywin32). Nada de senha nem de servidor.

Tudo o que é objeto do Outlook nasce e morre dentro de `Outlook(conta)`, na mesma thread: objeto
COM não atravessa thread. Para a tela só voltam dados simples (texto, datas, EntryID).

A conta é escolhida na tela (`contas()` lista as que o Outlook tem configuradas). Tudo é relativo a
ela: a Caixa de entrada que se lê, os Itens Enviados onde fica o modelo e de onde as respostas saem,
e o domínio que define o que é "e-mail interno".

Como o fluxo costuma ser num escritório contábil:
- o envio é um e-mail por lote de contatos (~40 cada), com Para = a própria conta e os clientes
  em Cco;
- o texto é sempre o mesmo, só muda o mês. Por isso o modelo é **o último envio** da pasta Itens
  Enviados: se alguém mudar o texto no Outlook, o mês seguinte já sai com o texto novo. Sem envio
  anterior, vale o texto padrão da ferramenta (`layouts.TEXTO_PADRAO`);
- os grupos de contatos podem estar em outra conta (uma conta IMAP não tem contatos): por isso os
  contatos são lidos em todas as lojas, e os e-mails só na conta escolhida;
- a resposta chega na Caixa de entrada, quase sempre como "RE:/Res:/Re: <assunto>", mas há quem
  mande e-mail novo ("NFS-e de serviços tomados de agosto"), quem mande antes de ser perguntado e
  quem caia em subpasta ou no lixo eletrônico.

A leitura VÊ TUDO que chegou de fora — Caixa de entrada, as subpastas dela e o lixo eletrônico,
desde o começo do mês da competência — e diz POR QUE cada e-mail entrou (`Resposta.relacao`):
"resposta" (cita a pergunta), "contato" (quem manda está no cadastro/planilha) ou "outro". Nada é
escondido: e-mail de outro assunto ou de outra competência aparece marcado, e a pessoa decide.

A parte que decide (ler_item, atribuir_competencia, janela_de_leitura, selecionar_respostas,
gravar_na_pasta) não toca no COM: recebe `ItemLido` e é a mesma para o Outlook de verdade e para o
Outlook simulado da demonstração e dos testes (falso.py).
"""
from __future__ import annotations

import filecmp
import glob
import html as html_lib
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

from . import competencia as C
from . import devolucao

# A conta é a escolhida na tela e guardada no config.json; não há conta padrão.
CONTA_PADRAO = ""
# grupos de contatos do Outlook marcados por padrão na 1ª abertura: os que começam com "Clientes"
GRUPO_PADRAO = re.compile(r"^clientes", re.I)
# quantos e-mails em Cco por mensagem (servidores costumam limitar perto de 50)
TAMANHO_LOTE = 40

PR_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
PR_LAST_VERB = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
DISPID_SEND_USING_ACCOUNT = 64209
OL_MAIL, OL_TO, OL_BCC, OL_BY_VALUE, OL_EMBEDDED_ITEM = 0, 1, 3, 1, 5

PASTAS_ENTRADA = ("caixa de entrada", "inbox")
PASTAS_ENVIADOS = ("itens enviados", "sent items", "enviados")
PASTAS_LIXO = ("lixo eletrônico", "lixo eletronico", "lixo", "lixeira", "junk e-mail", "junk email", "junk", "spam")

# onde termina o que a pessoa escreveu e começa a mensagem citada
_CITACAO = re.compile(
    r"^\s*(?:Em .{0,160}escreveu:|On .{0,160}wrote:|De:\s|From:\s|Enviado:\s|-{3,}\s*Original|"
    r"-{3,}\s*Mensagem|_{8,}|Obter o Outlook|Enviado do meu|Sent from my|Assunto:\s)",
    re.I | re.M)
_LINK = re.compile(r"https?://\S*(?:1drv\.ms|onedrive|drive\.google|we\.tl|wetransfer|dropbox|sharepoint|mega\.nz)",
                   re.I)
_NEGA = re.compile(r"\b(n[aã]o|nenhum|nenhuma|nada|sem servi)", re.I)
_SEGUE = re.compile(r"\b(segue|seguem|anexo|anexos|anexa|anexado|envio|enviando|encaminho)\b", re.I)
_SIM = re.compile(r"\b(sim|tivemos|tomamos|contratamos)\b", re.I)
# a imagem da assinatura do envio volta citada em muita resposta; e a assinatura pequena do próprio
# cliente (image001.jpg, Outlook-abc.png) também não é nota
_IMAGEM_ASSINATURA = re.compile(r"^(image\d{3}|outlook-\w+)\.(png|jpe?g|gif|bmp)$", re.I)
_IMG_CID = re.compile(r"<img[^>]*cid:[^>]*>", re.I)
# a frase que o texto original da pergunta sempre teve; continua valendo como marca, mas não é mais a única
FRASE_DA_PERGUNTA = "tomou (contratou)"

TIPOS = ("Mandou anexo", "Mandou link", "Sem serviços", "Ler", "Não entregue", "Interno")
RELACOES = {"resposta": "cita a pergunta", "contato": "está no cadastro", "outro": "outro assunto"}


def _data(valor) -> datetime:
    # pywin32 devolve hora local marcada como UTC; tirar o fuso e ficar com a hora como está
    return datetime(valor.year, valor.month, valor.day, valor.hour, valor.minute, valor.second)


def _prop(objeto, tag, padrao=None):
    try:
        return objeto.PropertyAccessor.GetProperty(tag)
    except Exception:
        return padrao


def _norm(texto: str) -> str:
    return re.sub(r"\s+", " ", C.sem_acento(texto or "").lower())


@dataclass
class Conta:
    endereco: str
    nome: str            # nome que aparece no Outlook
    loja: str            # nome da loja (arquivo de dados) onde a conta entrega

    @property
    def dominio(self) -> str:
        return self.endereco.rsplit("@", 1)[-1].lower()


@dataclass
class Grupo:
    nome: str
    membros: list[tuple[str, str]]  # (nome, e-mail)


@dataclass
class Modelo:
    assunto: str
    competencia: C.Competencia
    enviado_em: datetime
    html: str
    imagens: list[tuple[str, str]]  # (arquivo temporário, content-id)
    assinatura: str = ""            # a tag <img src="cid:..."> da assinatura, para o texto editado
    de_envio: bool = True           # False: não havia envio anterior; este é o texto padrão da ferramenta


@dataclass
class Envio:
    competencia: C.Competencia
    enviado_em: datetime
    destinatarios: list[tuple[str, str]]


@dataclass
class Anexo:
    indice: int
    nome: str
    tamanho: int


@dataclass
class AnexoLido:
    """Um anexo como veio do Outlook, com o que a leitura precisa para separar nota de assinatura."""
    indice: int
    nome: str
    tamanho: int
    cid: str = ""
    oculto: bool = False
    tipo: int = OL_BY_VALUE


@dataclass
class ItemLido:
    """Um e-mail já tirado do COM: só dados. É o que `ler_item` recebe (e o que os testes fabricam)."""
    entry_id: str
    recebido: datetime
    assunto: str
    classe: str
    corpo: str
    html: str
    email: str                       # remetente, minúsculo
    nome: str                        # nome de exibição do remetente
    anexos: list[AnexoLido] = field(default_factory=list)
    respondido: bool = False
    pasta: str = "Caixa de entrada"  # onde estava no Outlook (subpasta, lixo eletrônico...)


@dataclass
class Resposta:
    entry_id: str
    recebido: datetime
    nome: str
    email: str
    assunto: str
    texto: str                       # só o que a pessoa escreveu, sem a citação
    anexos: list[Anexo] = field(default_factory=list)
    embutidas: int = 0               # imagens de assinatura que não serão salvas
    tipo: str = "Ler"
    competencia: C.Competencia | None = None
    como_competencia: str = ""       # "assunto", "anexo" ou "data"
    nao_entregues: list[str] = field(default_factory=list)
    motivos: dict[str, tuple[str, bool]] = field(default_factory=dict)  # e-mail -> (motivo, vale reenviar)
    respondido: bool = False         # alguém já clicou "Responder" no Outlook
    relacao: str = "outro"           # "resposta" | "contato" | "outro" — por que este e-mail entrou na lista
    pasta_outlook: str = "Caixa de entrada"
    da_competencia: bool = True      # False: atribuído a outra competência (mostrado, não escondido)

    @property
    def relevante(self) -> bool:
        """O que vale tratar como resposta de cliente: cita a pergunta ou veio de quem está no cadastro."""
        return self.relacao in ("resposta", "contato")


@dataclass
class Contexto:
    """O que a leitura precisa saber além do próprio e-mail."""
    endereco: str                            # a conta escolhida, minúsculo
    dominio: str                             # domínio dela: e-mail desse domínio é "Interno"
    marcas: set[str]                         # trechos que denunciam a pergunta citada (assunto, nosso endereço...)
    assinaturas: set[int] = field(default_factory=set)   # tamanhos das imagens de assinatura do modelo
    perguntados: set[str] = field(default_factory=set)   # quem recebeu a pergunta desta competência
    conhecidos: set[str] = field(default_factory=set)    # e-mails do cadastro/planilha/confirmados


def contexto_de(endereco: str, envios: dict, assinaturas: set[int] | None = None,
                perguntados: Iterable[str] = (), conhecidos: Iterable[str] = ()) -> Contexto:
    endereco = endereco.lower()
    return Contexto(endereco, endereco.rsplit("@", 1)[-1], marcas_da_pergunta(endereco, envios, C.ultimas(13)),
                    set(assinaturas or ()), {p.lower() for p in perguntados}, {c.lower() for c in conhecidos})


def modelo_padrao(comp: C.Competencia | None = None) -> Modelo:
    """O modelo quando a conta nunca mandou a pergunta: assunto e texto padrão, sem assinatura."""
    from . import layouts as L
    comp = comp or C.anterior()
    agora = datetime.now()
    html = L.html(L.preencher(L.TEXTO_PADRAO, comp, agora))
    return Modelo(C.assunto(comp), comp, agora, html, [], "", de_envio=False)


def marcas_da_pergunta(endereco: str, envios: dict | None = None, competencias: list | None = None) -> set[str]:
    """Trechos que, no corpo de um e-mail, mostram que a pergunta está citada ali: o nosso endereço, a
    frase do texto original e os assuntos dos envios (de qualquer competência)."""
    marcas = {endereco.lower(), FRASE_DA_PERGUNTA}
    comps = set(envios or {}) | set(competencias or [])
    for comp in comps:
        marcas.add(_norm(C.assunto(comp)))
    return marcas


def e_sobre_a_pergunta(assunto: str, corpo: str, marcas: set[str]) -> bool:
    if "tomad" in _norm(assunto):
        return True
    corpo_n = _norm(corpo)
    return any(m and m in corpo_n for m in marcas)


def e_assinatura(a: AnexoLido, html: str, assinaturas: set[int]) -> bool:
    if a.oculto or (a.cid and f"cid:{a.cid}" in html):
        return True
    if not _IMAGEM_ASSINATURA.match(a.nome or ""):
        return False
    return a.tamanho < 60_000 or any(abs(a.tamanho - s) <= s * 0.02 for s in assinaturas)


def ler_item(d: ItemLido, ctx: Contexto) -> Resposta | None:
    """Classifica um e-mail. Devolve None só para o que não é e-mail de ninguém de fora (o que a própria
    conta mandou, item que não é mensagem, devolução de endereço que não foi perguntado)."""
    assunto, classe, corpo = d.assunto or "", d.classe or "", d.corpo or ""
    email = d.email if classe.startswith("IPM.Note") else ""
    sobre = e_sobre_a_pergunta(assunto, corpo, ctx.marcas)
    if devolucao.e_devolucao(email or d.nome, assunto, classe):
        # o aviso de "não entregue" raramente fala em "tomados": liga pelo endereço que foi perguntado
        falhas = devolucao.ler(corpo, {ctx.endereco})
        if ctx.perguntados:
            falhas = [f for f in falhas if f.email in ctx.perguntados]
        elif not sobre:
            falhas = []
        if not falhas:
            return None
        resposta = Resposta(d.entry_id, d.recebido, d.nome or email, email, assunto,
                            re.sub(r"\s+", " ", corpo)[:600], tipo="Não entregue", relacao="resposta",
                            pasta_outlook=d.pasta)
        resposta.nao_entregues = [f.email for f in falhas]
        resposta.motivos = {f.email: (f.motivo, f.reenviar) for f in falhas}
        return resposta
    if not classe.startswith("IPM.Note") or not email or email == ctx.endereco:
        return None

    texto = _CITACAO.split(corpo, maxsplit=1)[0].strip()
    resposta = Resposta(d.entry_id, d.recebido, d.nome or email, email, assunto,
                        re.sub(r"[ \t​]+", " ", re.sub(r"\n\s*\n+", "\n", texto)).strip()[:1500],
                        respondido=d.respondido, pasta_outlook=d.pasta)
    resposta.relacao = "resposta" if sobre else ("contato" if email in ctx.conhecidos else "outro")

    for a in d.anexos:
        if e_assinatura(a, d.html or "", ctx.assinaturas):
            resposta.embutidas += 1
        else:
            resposta.anexos.append(Anexo(a.indice, a.nome, a.tamanho))

    if ctx.dominio and email.endswith("@" + ctx.dominio):
        resposta.tipo = "Interno"
    elif resposta.anexos:
        resposta.tipo = "Mandou anexo"
    elif _LINK.search(corpo.split(FRASE_DA_PERGUNTA)[0]):
        resposta.tipo = "Mandou link"
    elif sobre and _NEGA.search(texto) and not _SEGUE.search(texto) and not _SIM.search(_sem_negacao(texto)):
        resposta.tipo = "Sem serviços"
    return resposta


def _sem_negacao(texto: str) -> str:
    """'Não tivemos serviços' não é 'tivemos': tira a palavra que vem logo depois de um 'não'."""
    return re.sub(r"\bn[aã]o\s+\w+", " ", texto, flags=re.I)


def atribuir_competencia(resposta: Resposta, comp: C.Competencia, inicios: list) -> tuple[C.Competencia, str]:
    """A competência de um e-mail sai do assunto ("Re: Serviços tomados - Agosto de 2026"), senão dos
    nomes dos anexos (".Setembro - Empresa.zip"), senão da data: vale o último envio feito antes de o
    e-mail chegar (`inicios` = [(data do 1º envio, competência)], em ordem)."""
    citada, como = C.mencionada(resposta.assunto, comp), "assunto"
    if citada is None and resposta.anexos:
        citada, como = C.mencionada(" ".join(a.nome for a in resposta.anexos), comp), "anexo"
    if citada is None:
        antes = [c for inicio, c in inicios if inicio <= resposta.recebido]
        citada, como = (antes[-1] if antes else comp), "data"
    return citada, como


def janela_de_leitura(comp: C.Competencia, envios: dict) -> tuple[datetime, datetime]:
    """(inicio_do_mes, desde): lê desde o 1º dia do mês da competência, mas o que chegou ANTES de `desde`
    (= o 1º dia do mês seguinte, ou o envio da pergunta, o que vier primeiro) só entra se citar a
    competência no assunto ou no nome do anexo — senão o mês inteiro de respostas da competência
    anterior viria junto."""
    inicio_do_mes = datetime(comp[0], comp[1], 1)
    seguinte = datetime(*C.proxima(comp), 1)
    desde = min([seguinte] + [e.enviado_em for e in envios.get(comp, [])])
    return inicio_do_mes, desde


def selecionar_respostas(itens: Iterable[tuple[ItemLido, bool]], comp: C.Competencia, envios: dict,
                         ctx: Contexto) -> list[Resposta]:
    """O miolo da busca, sem COM. `itens`: (e-mail já lido, veio do lixo eletrônico?), restritos ao que
    chegou desde o 1º dia do mês da competência. Devolve do mais novo para o mais velho."""
    inicios = sorted((min(e.enviado_em for e in lista), c) for c, lista in envios.items())
    _inicio_do_mes, desde = janela_de_leitura(comp, envios)
    achadas, vistos = [], set()
    for d, e_lixo in itens:
        if d.entry_id in vistos:
            continue
        r = ler_item(d, ctx)
        if r is None:
            continue
        citada, como = atribuir_competencia(r, comp, inicios)
        r.competencia, r.como_competencia, r.da_competencia = citada, como, citada == comp
        if d.recebido < desde and not (r.da_competencia and como in ("assunto", "anexo")):
            continue  # dentro do próprio mês, só o que já cita a competência
        if e_lixo and not r.relevante:
            continue  # lixo eletrônico sem relação com a pergunta nem com o cadastro: é lixo mesmo
        vistos.add(d.entry_id)
        achadas.append(r)
    achadas.sort(key=lambda r: r.recebido, reverse=True)
    return achadas


def gravar_na_pasta(tmp: Path, nome_original: str, pasta: Path) -> tuple[str, str]:
    """Leva o arquivo temporário `tmp` para `pasta` com o nome de origem. Nunca sobrescreve: arquivo igual
    (byte a byte) já na pasta vira "já estava"; nome repetido com conteúdo diferente ganha " (2)"."""
    nome = _nome_seguro(nome_original, pasta)
    candidatos = [pasta / nome] + sorted(pasta.glob(f"{glob.escape(Path(nome).stem)} (*){Path(nome).suffix}"))
    if pasta.is_dir():  # o mesmo arquivo pode estar lá com outro nome (nome encurtado pelo limite de caminho)
        tamanho = tmp.stat().st_size
        candidatos += [p for p in sorted(pasta.iterdir()) if p.is_file() and p.stat().st_size == tamanho]
    for existente in candidatos:
        if existente.is_file() and filecmp.cmp(existente, tmp, shallow=False):
            return existente.name, "já estava"
    alvo, n = pasta / nome, 2
    while alvo.exists():
        alvo = pasta / f"{Path(nome).stem} ({n}){Path(nome).suffix}"
        n += 1
    shutil.copyfile(tmp, alvo)
    return alvo.name, "salvo"


def contas() -> list[Conta]:
    """As contas de e-mail configuradas neste Outlook (sem abrir nenhuma)."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        achadas = []
        for a in ns.Accounts:
            endereco = (a.SmtpAddress or "").strip().lower()
            if not endereco:
                continue
            try:
                loja = a.DeliveryStore.DisplayName
            except Exception:
                loja = ""
            achadas.append(Conta(endereco, a.DisplayName or endereco, loja or ""))
        return achadas
    finally:
        pythoncom.CoUninitialize()


class Outlook:
    def __init__(self, conta: str):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self.assinaturas: set[int] = set()  # tamanhos das imagens do modelo; modelo() preenche
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        if not conta:
            raise RuntimeError("Escolha a conta do Outlook na parte de cima da tela.")
        self.conta = next((a for a in self.ns.Accounts if (a.SmtpAddress or "").lower() == conta.lower()), None)
        if self.conta is None:
            raise RuntimeError(f"A conta {conta} não está configurada neste Outlook.")
        self.endereco = conta.lower()
        self.dominio = self.endereco.rsplit("@", 1)[-1]
        self.loja = self.conta.DeliveryStore

    def __enter__(self):
        return self

    def __exit__(self, *_):
        import pythoncom
        self.app = self.ns = self.conta = self.loja = None
        pythoncom.CoUninitialize()

    # -------------------------------------------------------------- pastas
    def _pasta(self, nomes: tuple[str, ...], obrigatoria: bool = True):
        filhas = {f.Name.lower(): f for f in self.loja.GetRootFolder().Folders}
        for nome in nomes:
            if nome in filhas:
                return filhas[nome]
        if obrigatoria:
            raise RuntimeError(f"Não achei a pasta {nomes[0]!r} na conta {self.endereco}.")
        return None

    def enviados(self):
        return self._pasta(PASTAS_ENVIADOS)

    def entrada(self):
        return self._pasta(PASTAS_ENTRADA)

    def lixo(self):
        return self._pasta(PASTAS_LIXO, obrigatoria=False)

    def pastas_de_leitura(self, com_lixo: bool = True) -> list[tuple[object, str, bool]]:
        """[(pasta, rótulo, é lixo eletrônico)]: Caixa de entrada, todas as subpastas dela e o lixo."""
        achadas = []

        def descer(pasta, rotulo):
            achadas.append((pasta, rotulo, False))
            try:
                filhas = list(pasta.Folders)
            except Exception:
                filhas = []
            for f in sorted(filhas, key=lambda x: x.Name.lower()):
                descer(f, f"{rotulo}/{f.Name}")

        descer(self.entrada(), "Caixa de entrada")
        if com_lixo:
            lixo = self.lixo()
            if lixo is not None:
                achadas.append((lixo, "Lixo eletrônico", True))
        return achadas

    # -------------------------------------------------------------- contatos
    def _pastas_de_contatos(self):
        for loja in self.ns.Stores:
            try:
                raiz = loja.GetRootFolder()
            except Exception:
                continue
            for pasta in raiz.Folders:
                if pasta.DefaultItemType == 2:  # olContactItem
                    yield pasta

    def grupos(self) -> list[Grupo]:
        grupos = []
        for pasta in self._pastas_de_contatos():
            for item in pasta.Items.Restrict("[MessageClass] = 'IPM.DistList'"):
                membros = []
                for i in range(1, item.MemberCount + 1):
                    r = item.GetMember(i)
                    endereco = (r.Address or "").strip()
                    if "@" in endereco:
                        membros.append((r.Name or endereco, endereco))
                grupos.append(Grupo(item.DLName, membros))
        return grupos

    def nomes_dos_contatos(self) -> dict[str, str]:
        """{e-mail: nome} de todos os contatos, para reconhecer o cliente pelo nome."""
        nomes: dict[str, str] = {}
        for pasta in self._pastas_de_contatos():
            tabela = pasta.GetTable("[MessageClass] = 'IPM.Contact'")
            tabela.Columns.RemoveAll()
            for coluna in ("FullName", "CompanyName", "Email1Address", "Email2Address", "Email3Address"):
                tabela.Columns.Add(coluna)
            while not tabela.EndOfTable:
                linha = tabela.GetNextRow()
                nome = linha.Item("CompanyName") or linha.Item("FullName") or ""
                for coluna in ("Email1Address", "Email2Address", "Email3Address"):
                    email = (linha.Item(coluna) or "").strip().lower()
                    if "@" in email and nome:
                        nomes.setdefault(email, nome)
        return nomes

    # -------------------------------------------------------------- envios anteriores
    def envios(self) -> dict[C.Competencia, list[Envio]]:
        """Todos os envios no padrão "Serviços tomados - <Mês> de <ano>", por competência."""
        itens = self.enviados().Items
        itens.Sort("[SentOn]", True)
        achados: dict[C.Competencia, list[Envio]] = {}
        for m in itens:
            comp = C.do_assunto(getattr(m, "Subject", "") or "")
            if comp is None or m.Class != 43:
                continue
            destinatarios = []
            for r in m.Recipients:
                endereco = (r.Address or "").strip()
                if "@" in endereco and endereco.lower() != self.endereco:
                    destinatarios.append((r.Name or endereco, endereco))
            achados.setdefault(comp, []).append(Envio(comp, _data(m.SentOn), destinatarios))
        return achados

    def modelo(self) -> Modelo:
        itens = self.enviados().Items
        itens.Sort("[SentOn]", True)
        for m in itens:
            comp = C.do_assunto(getattr(m, "Subject", "") or "")
            if comp is None or m.Class != 43:
                continue
            html = m.HTMLBody
            pasta_tmp = Path(tempfile.mkdtemp(prefix="malote_modelo_"))
            imagens = []
            for a in m.Attachments:
                cid = _prop(a, PR_CONTENT_ID, "")
                arquivo = pasta_tmp / a.FileName
                a.SaveAsFile(str(arquivo))
                imagens.append((str(arquivo), cid or ""))
                self.assinaturas.add(a.Size)
            assinatura = _IMG_CID.search(html)
            return Modelo(m.Subject, comp, _data(m.SentOn), html, imagens,
                          assinatura.group(0) if assinatura else "")
        return modelo_padrao()  # primeira vez: esta conta nunca mandou a pergunta

    # -------------------------------------------------------------- enviar
    def _definir_conta(self, mail) -> None:
        mail.SendUsingAccount = self.conta
        atual = mail.SendUsingAccount
        if atual is None or (atual.SmtpAddress or "").lower() != self.endereco:
            # contorno conhecido do pywin32: atribuir pela DISPID quando a propriedade não "pega"
            mail._oleobj_.Invoke(DISPID_SEND_USING_ACCOUNT, 0, 8, 0, self.conta)
        atual = mail.SendUsingAccount
        if atual is None or (atual.SmtpAddress or "").lower() != self.endereco:
            raise RuntimeError(f"O Outlook não aceitou enviar pela conta {self.endereco}.")

    def montar(self, assunto: str, html: str, imagens: list[tuple[str, str]], cco: list[str], enviar: bool,
               so_testar: bool = False):
        """Um e-mail com os endereços em Cco. Devolve onde ficou ("enviado" ou a pasta do rascunho).

        html vem de layouts.html() (texto editado na tela + assinatura); imagens = as do modelo (cid).
        so_testar: monta tudo, devolve um resumo do que sairia e descarta (não envia nem salva).
        """
        mail = self.app.CreateItem(OL_MAIL)
        self._definir_conta(mail)
        mail.Subject = assunto
        mail.To = self.endereco
        for endereco in cco:
            r = mail.Recipients.Add(endereco)
            r.Type = OL_BCC
        if not mail.Recipients.ResolveAll():
            ruins = [r.Name for r in mail.Recipients if not r.Resolved]
            mail.Close(1)  # olDiscard
            raise RuntimeError("Endereço que o Outlook não reconheceu: " + ", ".join(ruins))
        for arquivo, cid in imagens:
            anexo = mail.Attachments.Add(arquivo, OL_BY_VALUE, 0)
            if cid:
                anexo.PropertyAccessor.SetProperty(PR_CONTENT_ID, cid)
        mail.HTMLBody = html
        mail.SaveSentMessageFolder = self.enviados()  # o mês seguinte procura o modelo aqui
        if so_testar:
            resumo = {"conta": mail.SendUsingAccount.SmtpAddress, "assunto": mail.Subject, "para": mail.To,
                      "cco": [r.Address for r in mail.Recipients if r.Type == OL_BCC],
                      "anexos": [(a.FileName, _prop(a, PR_CONTENT_ID, "")) for a in mail.Attachments],
                      "html": mail.HTMLBody, "salvar_em": mail.SaveSentMessageFolder.FolderPath}
            mail.Close(1)  # olDiscard
            return resumo
        if enviar:
            mail.Send()
            return "enviado"
        mail.Save()
        return mail.Parent.FolderPath

    # -------------------------------------------------------------- respostas
    def _remetente(self, item) -> str:
        try:
            if item.SenderEmailType == "EX":
                return item.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
        except Exception:
            pass
        return (getattr(item, "SenderEmailAddress", "") or "").lower()

    def _extrair(self, item, recebido: datetime, pasta: str) -> ItemLido:
        """Tira do COM tudo o que a leitura usa. Só aqui se toca no item."""
        classe = getattr(item, "MessageClass", "") or ""
        anexos = []
        try:
            for a in item.Attachments:
                anexos.append(AnexoLido(a.Index, a.FileName or "", a.Size or 0, _prop(a, PR_CONTENT_ID, "") or "",
                                        bool(_prop(a, PR_HIDDEN, False)), a.Type))
        except Exception:
            pass
        return ItemLido(
            entry_id=item.EntryID, recebido=recebido, assunto=getattr(item, "Subject", "") or "", classe=classe,
            corpo=getattr(item, "Body", "") or "", html=getattr(item, "HTMLBody", "") or "",
            email=self._remetente(item) if classe.startswith("IPM.Note") else "",
            nome=getattr(item, "SenderName", "") or "", anexos=anexos,
            respondido=_prop(item, PR_LAST_VERB, 0) in (102, 103), pasta=pasta)

    def _itens_desde(self, pasta, limite: datetime):
        """Itens da pasta do mais novo para o mais velho, até `limite` (exclusivo)."""
        itens = pasta.Items
        itens.Sort("[ReceivedTime]", True)
        for item in itens:
            try:
                recebido = _data(item.ReceivedTime)
            except Exception:
                continue
            if recebido < limite:
                break
            yield item, recebido

    def _itens_lidos(self, limite: datetime, com_lixo: bool = True):
        """(ItemLido, é lixo) de todas as pastas de leitura, do que chegou desde `limite`."""
        for pasta, rotulo, e_lixo in self.pastas_de_leitura(com_lixo=com_lixo):
            for item, recebido in self._itens_desde(pasta, limite):
                yield self._extrair(item, recebido, rotulo), e_lixo

    def respostas(self, comp: C.Competencia, envios: dict[C.Competencia, list[Envio]], perguntados: set[str],
                  conhecidos: set[str] | None = None) -> list[Resposta]:
        """Tudo o que chegou de fora para a competência `comp`: respostas à pergunta, e-mails de quem está
        no cadastro e até o que chegou sem relação aparente (relacao "outro"), para a tela mostrar e a
        pessoa decidir. Olha a Caixa de entrada, as subpastas dela e o lixo eletrônico (deste só entra o
        que é resposta ou de contato conhecido). E-mail atribuído a outra competência vem com
        `da_competencia=False` em vez de sumir."""
        if not self.assinaturas:
            try:
                self.modelo()  # aprende o tamanho da imagem da assinatura, para não salvá-la como nota
            except RuntimeError:
                pass
        ctx = contexto_de(self.endereco, envios, self.assinaturas, perguntados, conhecidos or ())
        inicio_do_mes, _desde = janela_de_leitura(comp, envios)
        return selecionar_respostas(self._itens_lidos(inicio_do_mes), comp, envios, ctx)

    def quem_respondeu(self, desde: datetime, envios: dict | None = None) -> dict[str, datetime]:
        """{e-mail: última vez que respondeu à pergunta} na Caixa de entrada (e subpastas), desde `desde`.

        Conta qualquer resposta ("não houve" também): o que interessa é quem lê e responde.
        """
        marcas = marcas_da_pergunta(self.endereco, envios, C.ultimas(24))
        achados: dict[str, datetime] = {}
        for pasta, _rotulo, _e_lixo in self.pastas_de_leitura(com_lixo=False):
            for item, recebido in self._itens_desde(pasta, desde):
                classe = getattr(item, "MessageClass", "") or ""
                if not classe.startswith("IPM.Note"):
                    continue
                email = self._remetente(item)
                assunto = getattr(item, "Subject", "") or ""
                if (not email or email == self.endereco or email.endswith("@" + self.dominio)
                        or devolucao.e_devolucao(email, assunto, classe)):
                    continue
                if e_sobre_a_pergunta(assunto, getattr(item, "Body", "") or "", marcas):
                    if email not in achados or recebido > achados[email]:
                        achados[email] = recebido
        return achados

    # -------------------------------------------------------------- salvar
    def salvar_anexos(self, entry_id: str, indices: list[int], pasta: Path) -> list[tuple[str, str]]:
        """Grava os anexos pedidos em `pasta`. Nunca sobrescreve: nome repetido ganha " (2)".

        Devolve [(nome gravado, "salvo" | "já estava")].
        """
        item = self.ns.GetItemFromID(entry_id, self.loja.StoreID)
        feitos: list[tuple[str, str]] = []
        for a in item.Attachments:
            if a.Index not in indices:
                continue
            if a.Type == OL_EMBEDDED_ITEM:  # e-mail anexado (.msg): o que interessa são as notas dentro dele
                feitos += self._salvar_de_dentro(a, pasta)
            else:
                feitos.append(self._salvar_um(a, pasta))
        return feitos

    def _salvar_um(self, a, pasta: Path) -> tuple[str, str]:
        """Grava primeiro no %TEMP% e deixa `gravar_na_pasta` decidir (igual, repetido, novo).

        O a.Size do Outlook inclui sobra do MAPI e não serve para comparar com o arquivo.
        """
        tmp = Path(tempfile.mkdtemp(prefix="malote_anexo_")) / "anexo"
        a.SaveAsFile(str(tmp))
        try:
            return gravar_na_pasta(tmp, a.FileName, pasta)
        finally:
            shutil.rmtree(tmp.parent, ignore_errors=True)

    def _salvar_de_dentro(self, anexo, pasta: Path) -> list[tuple[str, str]]:
        tmp = Path(tempfile.mkdtemp(prefix="malote_msg_")) / "anexado.msg"
        anexo.SaveAsFile(str(tmp))
        try:
            interno = self.ns.OpenSharedItem(str(tmp))
            html = getattr(interno, "HTMLBody", "") or ""
            de_dentro = [a for a in interno.Attachments
                         if not e_assinatura(AnexoLido(a.Index, a.FileName or "", a.Size or 0,
                                                       _prop(a, PR_CONTENT_ID, "") or "",
                                                       bool(_prop(a, PR_HIDDEN, False)), a.Type),
                                             html, self.assinaturas)]
        except Exception:
            de_dentro = []
        if not de_dentro:  # não abriu ou não tem nada dentro: guarda o .msg como veio
            return [self._salvar_um(anexo, pasta)]
        feitos = []
        for a in de_dentro:
            if a.Type == OL_EMBEDDED_ITEM:
                feitos += self._salvar_de_dentro(a, pasta)
            else:
                feitos.append(self._salvar_um(a, pasta))
        return feitos

    def cnpjs_dos_anexos(self, entry_id: str, indices: list[int]) -> dict[int, set[str]]:
        """{índice do anexo: CNPJs no texto}. Grava num temporário e apaga em seguida."""
        from . import anexos
        item = self.ns.GetItemFromID(entry_id, self.loja.StoreID)
        tmp = Path(tempfile.mkdtemp(prefix="malote_cnpj_"))
        achados: dict[int, set[str]] = {}
        try:
            for a in item.Attachments:
                if a.Index not in indices or a.Type == OL_EMBEDDED_ITEM:
                    continue
                arquivo = tmp / f"{a.Index}{Path(a.FileName).suffix}"
                a.SaveAsFile(str(arquivo))
                achados[a.Index] = anexos.cnpjs(arquivo, a.FileName)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        return achados

    # -------------------------------------------------------------- responder
    def ja_respondido(self, entry_id: str) -> bool:
        item = self.ns.GetItemFromID(entry_id, self.loja.StoreID)
        return _prop(item, PR_LAST_VERB, 0) in (102, 103)  # responder / responder a todos

    def responder(self, entry_id: str, texto: str, modelo: Modelo, so_testar: bool = False):
        """Responde o e-mail do cliente na mesma conversa, pela conta escolhida, com a assinatura do modelo.

        so_testar: monta a resposta, devolve um resumo e descarta (não envia nem salva).
        """
        item = self.ns.GetItemFromID(entry_id, self.loja.StoreID)
        r = item.Reply()
        self._definir_conta(r)
        corpo = r.HTMLBody
        # A conta pode ter assinatura automática de resposta (texto + imagem), que o Outlook põe antes da
        # mensagem citada. O que vai antes da citação é trocado pelo texto aprovado; a imagem da assinatura
        # automática (já anexada pelo Outlook) é reaproveitada.
        inicio = re.search(r"<div class=\"?WordSection1\"?>|<body[^>]*>", corpo, re.I)
        citacao = corpo.find('<a name="_MailOriginal">')
        fim = corpo.rfind("<div", 0, citacao) if citacao > 0 else -1
        if inicio and fim > inicio.end():
            antes = corpo[inicio.end():fim]
            imagem = _IMG_CID.search(antes)
        else:
            antes, imagem = "", None
        if imagem is None:  # sem assinatura automática: usa a do modelo do envio
            imagem = _IMG_CID.search(modelo.html)
            for arquivo, cid in modelo.imagens:
                if cid:
                    anexo = r.Attachments.Add(arquivo, OL_BY_VALUE, 0)
                    anexo.PropertyAccessor.SetProperty(PR_CONTENT_ID, cid)
        corpo = _inserir_resposta(corpo, texto, imagem.group(0) if imagem else "", inicio, fim)
        r.HTMLBody = corpo
        r.SaveSentMessageFolder = self.enviados()
        if so_testar:
            resumo = {"conta": r.SendUsingAccount.SmtpAddress, "assunto": r.Subject, "para": r.To,
                      "anexos": [(a.FileName, _prop(a, PR_CONTENT_ID, "")) for a in r.Attachments],
                      "texto": r.Body[:600], "html": r.HTMLBody}
            r.Close(1)  # olDiscard
            return resumo
        r.Send()

    def respostas_antigas(self) -> list[tuple[str, str]]:
        """(mensagem do cliente, o que a equipe respondeu) dos "RES: Serviços tomados" de Itens Enviados."""
        pares = []
        for m in self.enviados().Items:
            assunto = getattr(m, "Subject", "") or ""
            corpo = (getattr(m, "Body", "") or "").replace("\r", "")
            par = par_de_resposta_antiga(assunto, corpo)
            if par:
                pares.append(par)
        return pares


def _inserir_resposta(corpo: str, texto: str, imagem: str, inicio, fim: int) -> str:
    """Põe o texto aprovado (e a imagem da assinatura) antes da mensagem citada no HTML da resposta."""
    estilo = "margin:0;font-family:Aptos,Calibri,sans-serif;font-size:11pt"
    nosso = "".join(f'<p class=MsoNormal style="{estilo}">{html_lib.escape(l) or "&nbsp;"}</p>'
                    for l in texto.splitlines())
    nosso += f"<p class=MsoNormal>{imagem}</p><p class=MsoNormal>&nbsp;</p>"
    if inicio and fim > inicio.end():
        return corpo[:inicio.end()] + nosso + corpo[fim:]
    if inicio:
        return corpo[:inicio.end()] + nosso + corpo[inicio.end():]
    return nosso + corpo


def par_de_resposta_antiga(assunto: str, corpo: str) -> tuple[str, str] | None:
    """De um "RES: Serviços tomados..." enviado pela equipe: (o que o cliente escreveu, o que a equipe
    respondeu), para a base de respostas aprender. Encaminhamentos (ENC/FW) não são resposta ao cliente."""
    if not re.match(r"\s*(res|re|enc|fw|fwd)\s*:", assunto, re.I) or "tomad" not in assunto.lower():
        return None
    if re.match(r"\s*(enc|fw|fwd)\s*:", assunto, re.I):
        return None
    corpo = (corpo or "").replace("\r", "")
    partes = _CITACAO.split(corpo, maxsplit=1)
    if len(partes) < 2:
        return None
    nosso = partes[0]
    depois = corpo[len(nosso):]
    cab = re.search(r"^\s*(Assunto|Subject):.*$", depois, re.M)
    cliente = depois[cab.end():] if cab else depois.split("\n", 1)[-1]
    cliente = _CITACAO.split(cliente, maxsplit=1)[0]
    cliente = re.sub(r"\s+", " ", cliente).strip()
    if cliente and nosso.strip():
        return cliente[:600], nosso
    return None


def _nome_seguro(nome: str, pasta: Path) -> str:
    nome = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nome or "anexo").strip().rstrip(".") or "anexo"
    sobra = 250 - len(str(pasta)) - 6  # o Outlook não grava caminho com mais de 260 caracteres
    if len(nome) > sobra:
        sufixo = Path(nome).suffix[:10]
        nome = Path(nome).stem[: max(sobra - len(sufixo), 8)] + sufixo
    return nome
