"""Cadastro próprio de contatos (cadastro.json): quem recebe a pergunta e de qual empresa é cada e-mail.

Regras:
- a planilha de clientes e os grupos do Outlook ATUALIZAM o cadastro a cada abertura: e-mail
  novo entra; e-mail que some da planilha só é sinalizado (nunca apagado);
- o que o usuário muda à mão (tirar, voltar, acrescentar, trocar e-mail) VALE MAIS que a planilha;
- pedido de cliente para trocar e-mail ou sair da lista vira mudança pronta na aba Responder e só
  é aplicada quando o usuário aceita (aprendizado supervisionado);
- empresa com vários e-mails recebe só em quem responde (últimos 6 meses) ou no principal fixado
  à mão; sem histórico, em todos.

Cada contato:
    email, codigos (empresas), nome, situacao ("ativo" | "fora"), motivo, origens ("planilha",
    "grupo:<nome>", "manual"), forcar_envio (marcado à mão mesmo sendo ex-cliente), sumiu_em
    (deixou de vir na planilha), historico [{quando, o_que, por}]
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

_EMAIL = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")


def email_valido(email: str) -> bool:
    return bool(_EMAIL.match((email or "").strip().lower()))


def _agora() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


@dataclass
class Contato:
    email: str
    codigos: list[str] = field(default_factory=list)
    nome: str = ""
    situacao: str = "ativo"
    motivo: str = ""
    origens: list[str] = field(default_factory=list)
    forcar_envio: bool = False
    sumiu_em: str = ""
    historico: list[dict] = field(default_factory=list)
    respondeu_em: str = ""     # AAAA-MM-DD da última resposta à pergunta de tomados
    principal: bool = False    # fixado à mão: a empresa recebe só neste (e nos outros principais)

    @property
    def fora(self) -> bool:
        return self.situacao == "fora"

    @property
    def manual(self) -> bool:
        return "manual" in self.origens

    def anotar(self, o_que: str, por: str) -> None:
        self.historico.append({"quando": _agora(), "o_que": o_que, "por": por})


class Cadastro:
    def __init__(self, pasta: Path):
        self.caminho = Path(pasta) / "cadastro.json"
        try:
            dados = json.loads(self.caminho.read_text(encoding="utf-8"))
        except FileNotFoundError:
            dados = {}
        self.contatos: dict[str, Contato] = {e: Contato(**c) for e, c in dados.get("contatos", {}).items()}
        self.ultima_atualizacao: str = dados.get("ultima_atualizacao", "")
        # até quando a Caixa de entrada já foi varrida atrás de quem responde (AAAA-MM-DD HH:MM)
        self.respostas_lidas_ate: str = dados.get("respostas_lidas_ate", "")

    def gravar(self) -> None:
        tmp = self.caminho.with_suffix(".tmp")
        tmp.write_text(json.dumps({"ultima_atualizacao": self.ultima_atualizacao,
                                   "respostas_lidas_ate": self.respostas_lidas_ate,
                                   "contatos": {e: asdict(c) for e, c in sorted(self.contatos.items())}},
                                  ensure_ascii=False, indent=1), encoding="utf-8")
        os.replace(tmp, self.caminho)

    # -------------------------------------------------------------- quem responde
    def anotar_respostas(self, quando_por_email: dict[str, datetime], lido_ate: datetime | None = None) -> int:
        """Guarda a data da última resposta de cada e-mail. Devolve quantos mudaram."""
        mudou = 0
        for email, quando in quando_por_email.items():
            c = self.contatos.get(email.lower())
            dia = quando.strftime("%Y-%m-%d")
            if c and dia > c.respondeu_em:
                c.respondeu_em = dia
                mudou += 1
        if lido_ate:
            self.respostas_lidas_ate = lido_ate.strftime("%Y-%m-%d %H:%M")
        self.gravar()
        return mudou

    def respondentes(self, meses: int = 6) -> set[str]:
        """Quem respondeu nos últimos `meses` meses e está na lista."""
        limite = datetime.now()
        ano, mes = limite.year, limite.month - meses
        while mes <= 0:
            ano, mes = ano - 1, mes + 12
        corte = f"{ano:04d}-{mes:02d}-{limite.day:02d}" if limite.day <= 28 else f"{ano:04d}-{mes:02d}-28"
        return {e for e, c in self.contatos.items() if c.respondeu_em >= corte and not c.fora}

    def principais(self) -> set[str]:
        return {e for e, c in self.contatos.items() if c.principal and not c.fora}

    def fixar_principal(self, email: str, sim: bool, por: str = "você") -> None:
        c = self._garantir(email)
        c.principal = sim
        c.anotar("fixado como contato principal" if sim else "deixou de ser o contato principal", por)
        self.gravar()

    def __getitem__(self, email: str) -> Contato | None:
        return self.contatos.get((email or "").lower().strip())

    # -------------------------------------------------------------- atualizar pela planilha/grupos
    def atualizar(self, empresas, grupos: dict[str, list[str]], planilha_nome: str) -> dict[str, list[str]]:
        """Traz o que é novo da planilha e dos grupos. Devolve {"novos": [...], "sumiram": [...]}."""
        da_planilha: dict[str, list[str]] = {}
        nomes: dict[str, str] = {}
        for e in empresas:
            if e.ex_cliente:
                continue
            for m in e.emails:
                da_planilha.setdefault(m, [])
                if e.codigo not in da_planilha[m]:
                    da_planilha[m].append(e.codigo)
                nomes.setdefault(m, e.razao)
        novos, sumiram = [], []
        for email, codigos in da_planilha.items():
            c = self.contatos.get(email)
            if c is None:
                c = self.contatos[email] = Contato(email, list(codigos), nomes.get(email, ""))
                c.anotar(f"entrou pela planilha ({', '.join(codigos)})", planilha_nome)
                novos.append(email)
            else:
                for cod in codigos:
                    if cod not in c.codigos:
                        c.codigos.append(cod)
                        c.anotar(f"empresa {cod} acrescentada pela planilha", planilha_nome)
                if c.sumiu_em:
                    c.anotar("voltou a aparecer na planilha", planilha_nome)
                    c.sumiu_em = ""
            if "planilha" not in c.origens:
                c.origens.append("planilha")
        for email, nomes_grupos in grupos.items():
            c = self.contatos.get(email)
            if c is None:
                c = self.contatos[email] = Contato(email)
                c.anotar(f"entrou pelo grupo {nomes_grupos[0]}", "Outlook")
                novos.append(email)
            for g in nomes_grupos:
                if f"grupo:{g}" not in c.origens:
                    c.origens.append(f"grupo:{g}")
        for email, c in self.contatos.items():
            if "planilha" in c.origens and email not in da_planilha and not c.sumiu_em:
                c.sumiu_em = _agora()[:10]
                c.anotar("não veio na planilha", planilha_nome)
                sumiram.append(email)
        self.ultima_atualizacao = f"{planilha_nome} em {_agora()}"
        self.gravar()
        return {"novos": novos, "sumiram": sumiram}

    # -------------------------------------------------------------- mudanças do usuário
    def tirar(self, email: str, motivo: str, por: str = "você") -> None:
        c = self._garantir(email)
        c.situacao, c.motivo, c.forcar_envio = "fora", motivo, False
        c.anotar(f"tirado da lista: {motivo}", por)
        self.gravar()

    def voltar(self, email: str, por: str = "você") -> None:
        c = self._garantir(email)
        c.situacao, c.motivo = "ativo", ""
        c.anotar("voltou para a lista", por)
        self.gravar()

    def marcar(self, email: str, recebe: bool, padrao: bool) -> None:
        """Clique no ✔ da aba 1. `padrao`: se receberia sem ninguém mexer."""
        c = self._garantir(email)
        if recebe:
            c.situacao, c.motivo = "ativo", ""
            c.forcar_envio = not padrao
            c.anotar("marcado para receber", "você")
        else:
            c.situacao, c.motivo, c.forcar_envio = "fora", "desmarcado na aba 1", False
            c.anotar("desmarcado na aba 1", "você")
        self.gravar()

    def acrescentar(self, email: str, codigos: list[str], nome: str = "", por: str = "você") -> Contato:
        email = email.lower().strip()
        if not email_valido(email):
            raise ValueError(f"E-mail inválido: {email}")
        c = self._garantir(email)
        for cod in codigos:
            if cod not in c.codigos:
                c.codigos.append(cod)
        if nome:
            c.nome = nome
        if "manual" not in c.origens:
            c.origens.append("manual")
        c.situacao, c.motivo = "ativo", ""
        c.anotar(f"acrescentado ({', '.join(codigos) or 'sem empresa'})", por)
        self.gravar()
        return c

    def trocar(self, antigo: str, novo: str, por: str = "você") -> Contato:
        """O novo herda as empresas do antigo; o antigo sai da lista (fica no histórico)."""
        antigo, novo = antigo.lower().strip(), novo.lower().strip()
        if not email_valido(novo):
            raise ValueError(f"E-mail inválido: {novo}")
        velho = self._garantir(antigo)
        c = self.acrescentar(novo, list(velho.codigos), velho.nome, por)
        c.anotar(f"substitui {antigo}", por)
        velho.situacao, velho.motivo, velho.forcar_envio = "fora", f"trocado por {novo}", False
        velho.anotar(f"trocado por {novo}", por)
        self.gravar()
        return c

    def _garantir(self, email: str) -> Contato:
        email = email.lower().strip()
        if email not in self.contatos:
            self.contatos[email] = Contato(email)
        return self.contatos[email]

    # -------------------------------------------------------------- para a lista de envio
    def fora_da_lista(self) -> dict[str, str]:
        return {e: c.motivo or "fora da lista" for e, c in self.contatos.items() if c.fora}

    def forcados(self) -> set[str]:
        return {e for e, c in self.contatos.items() if c.forcar_envio and not c.fora}

    def manuais(self) -> dict[str, list[str]]:
        """{e-mail: códigos} acrescentados à mão (recebem mesmo sem estar na planilha nem nos grupos)."""
        return {e: list(c.codigos) for e, c in self.contatos.items() if c.manual and not c.fora}

    def migrar(self, fora_da_lista: dict[str, str], incluir: set[str]) -> None:
        """Traz as marcações da versão anterior (registro.json), uma vez só."""
        mexeu = False
        for email, motivo in fora_da_lista.items():
            c = self._garantir(email)
            if not c.fora:
                c.situacao, c.motivo = "fora", motivo
                c.anotar(f"tirado da lista: {motivo}", "versão anterior")
                mexeu = True
        for email in incluir:
            c = self._garantir(email)
            if not c.forcar_envio:
                c.forcar_envio = True
                mexeu = True
        if mexeu:
            self.gravar()


# ------------------------------------------------------------------ pedido do cliente
def pedido_de_troca(remetente: str, texto: str, cadastro: Cadastro, dominio_interno: str = "") -> tuple[str, str] | None:
    """(antigo, novo) quando o cliente pede para trocar o e-mail; None se não der para saber.

    Ex.: "Solicito mudança do e-mail financeiro@empresa.example para
    contabil@empresa.example". O antigo é o que está ativo no cadastro; o novo, o que não está.
    """
    citados = []
    for e in re.findall(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", texto or ""):
        e = e.lower().strip(".")
        if e not in citados and not (dominio_interno and e.endswith("@" + dominio_interno.lower())):
            citados.append(e)
    candidatos = citados + ([remetente.lower()] if remetente and remetente.lower() not in citados else [])
    conhecidos = [e for e in candidatos if cadastro[e] and not cadastro[e].fora]
    desconhecidos = [e for e in candidatos if not cadastro[e]]
    if len(conhecidos) == 1 and len(desconhecidos) == 1:
        return conhecidos[0], desconhecidos[0]
    if len(citados) == 1 and citados[0] != remetente.lower() and cadastro[remetente]:
        return remetente.lower(), citados[0]
    return None
