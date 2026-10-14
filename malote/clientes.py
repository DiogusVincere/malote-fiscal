"""Pastas de clientes: quem é o remetente e em que pasta o anexo dele vai parar.

A raiz é a pasta onde cada cliente tem a sua subpasta, nomeada "<código> - <nome>" (o código é o que
liga a pasta à planilha de clientes). O destino de cada anexo é:

    <cliente>\\NFE\\<AAAA>\\<MM-AAAA>\\Tomados\\Tomados e-mail

Cada nível é procurado com as grafias que costumam existir numa rede antiga e criado só se faltar:
- NFE: "NFE", "NFe", "NF-E";
- competência: "08-2026", "08.2026", "082026" (cria "08-2026");
- Tomados: a pasta que já existir com "tomad" no nome ("Tomados", "Serviços tomados",
  "NFSE Tomados"...), preferindo "Tomados"; senão cria "Tomados".
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from rapidfuzz import fuzz, process

from .competencia import Competencia, sem_acento

PASTA_FINAL = "Tomados e-mail"

# provedores de e-mail pessoal: o domínio não diz nada sobre a empresa
PROVEDORES = {
    "gmail.com", "hotmail.com", "hotmail.com.br", "outlook.com", "outlook.com.br", "live.com",
    "msn.com", "yahoo.com", "yahoo.com.br", "terra.com.br", "uol.com.br", "bol.com.br",
    "icloud.com", "me.com", "globo.com", "ig.com.br", "r7.com", "adv.oabsp.org.br",
}

# palavras que não ajudam a distinguir uma empresa de outra
_VAZIAS = {"LTDA", "ME", "EPP", "EIRELI", "EIRE", "SA", "S", "A", "SS", "DE", "DA", "DO", "DAS",
           "DOS", "E", "EM", "CIA", "SOCIEDADE", "SIMPLES", "LIMITADA", "UNIPESSOAL", "MEI"}


def normalizar(texto: str) -> str:
    """'J&J SOFA CENTER LTDA' -> 'J J SOFA CENTER'"""
    palavras = re.sub(r"[^A-Z0-9]+", " ", sem_acento(texto or "").upper()).split()
    return " ".join(p for p in palavras if p not in _VAZIAS)


@dataclass(frozen=True)
class Cliente:
    pasta: str          # nome da pasta, que é o que se guarda (o código pode se repetir na rede)
    codigo: str | None
    nome: str
    chave: str          # nome normalizado, para comparar

    @property
    def rotulo(self) -> str:
        return self.pasta


def indexar(raiz: str | Path) -> list[Cliente]:
    clientes = []
    raiz = Path(raiz)
    if not raiz.is_dir():
        return clientes
    for item in sorted(raiz.iterdir(), key=lambda p: p.name.upper()):
        if not item.is_dir():
            continue
        m = re.match(r"^\s*(\d+)\s*[-–.]*\s*(.*)$", item.name)
        codigo, nome = (m.group(1), m.group(2)) if m else (None, item.name)
        clientes.append(Cliente(item.name, codigo, nome.strip(), normalizar(nome)))
    return clientes


@dataclass
class Palpite:
    pasta: str | None
    como: str           # "planilha", "confirmado", "mesmo domínio", "nome do contato", "domínio",
                        # "nome no e-mail", "várias empresas" ou ""
    certeza: int = 0    # 0-100, só para ordenar e explicar
    candidatos: list[str] = field(default_factory=list)  # e-mail de várias empresas: as pastas possíveis

    @property
    def certo(self) -> bool:
        """Pode salvar sem perguntar: veio da planilha ou o usuário confirmou."""
        return self.como in ("planilha", "confirmado") and bool(self.pasta)


def dominio(email: str) -> str:
    return email.rsplit("@", 1)[-1].lower().strip() if "@" in (email or "") else ""


class Identificador:
    """Descobre de qual cliente é um e-mail.

    planilha: {e-mail: [Empresa]} da planilha de clientes — manda em tudo.
    aprendidos: {e-mail: pasta} confirmados pelo usuário (registro.json), para quem não está na planilha.
    nomes: {e-mail: nome} tirado dos contatos do Outlook e dos envios antigos (último recurso).
    """

    def __init__(self, clientes: list[Cliente], aprendidos: dict[str, str], nomes: dict[str, str],
                 planilha: dict[str, list] | None = None):
        self.clientes = clientes
        self.por_pasta = {c.pasta: c for c in clientes}
        self.aprendidos = {k.lower(): v for k, v in aprendidos.items()}
        self.nomes = {k.lower(): v for k, v in nomes.items()}
        self.planilha = planilha or {}
        self._chaves = [c.chave for c in clientes]
        self._cache: dict[tuple[str, str], str | None] = {}
        self._por_codigo: dict[str, list[Cliente]] = {}
        for c in clientes:
            if c.codigo:
                self._por_codigo.setdefault(c.codigo.lstrip("0") or "0", []).append(c)

    def pasta_da_empresa(self, empresa) -> str | None:
        """Código da planilha -> pasta. Código repetido na rede ou sem pasta: decide pelo nome."""
        chave = (empresa.codigo, empresa.razao)
        if chave not in self._cache:
            self._cache[chave] = self._pasta_da_empresa(empresa)
        return self._cache[chave]

    def _pasta_da_empresa(self, empresa) -> str | None:
        mesmos = self._por_codigo.get(empresa.codigo.lstrip("0") or "0", []) if empresa.codigo.isdigit() else []
        if len(mesmos) == 1:
            return mesmos[0].pasta
        alvo = normalizar(empresa.razao)
        if mesmos:
            return max(mesmos, key=lambda c: fuzz.token_sort_ratio(alvo, c.chave)).pasta
        palpite = self._pelo_nome(empresa.razao, "planilha")
        return palpite.pasta if palpite else None

    def identificar(self, email: str, nome_exibicao: str = "") -> Palpite:
        email = (email or "").lower().strip()
        empresas = self.planilha.get(email, [])
        if empresas:
            pastas = []
            for e in empresas:
                p = self.pasta_da_empresa(e)
                if p and p not in pastas:
                    pastas.append(p)
            if len(pastas) == 1:
                return Palpite(pastas[0], "planilha", 100, pastas)
            if len(pastas) > 1:
                return Palpite(None, "várias empresas", 0, pastas)
        if email in self.aprendidos and self.aprendidos[email] in self.por_pasta:
            return Palpite(self.aprendidos[email], "confirmado", 100)

        dom = dominio(email)
        if dom and dom not in PROVEDORES:
            pastas = {p for e, p in self.aprendidos.items() if dominio(e) == dom and p in self.por_pasta}
            for e, lista in self.planilha.items():
                if dominio(e) == dom:
                    pastas |= {p for p in (self.pasta_da_empresa(x) for x in lista) if p}
            if len(pastas) == 1:
                return Palpite(pastas.pop(), "mesmo domínio", 95)

        for texto, como in ((self.nomes.get(email, ""), "nome do contato"), (nome_exibicao, "nome no e-mail")):
            palpite = self._pelo_nome(texto, como)
            if palpite:
                return palpite

        if dom and dom not in PROVEDORES:
            palpite = self._pelo_dominio(dom)
            if palpite:
                return palpite
        return Palpite(None, "", 0)

    def _pelo_nome(self, texto: str, como: str) -> Palpite | None:
        texto = re.sub(r"\(.*?@.*?\)", "", texto or "")   # "EMPRESA (x@y.com)" -> "EMPRESA"
        chave = normalizar(texto)
        if len(chave) < 5 or "@" in texto:
            return None
        achados = process.extract(chave, self._chaves, scorer=fuzz.token_sort_ratio, limit=3)
        if not achados or achados[0][1] < 88:
            return None
        melhor = achados[0]
        # segundo colocado quase empatado e com outro nome: ambíguo (matriz e filial com o mesmo nome)
        rivais = [a for a in achados[1:] if a[1] >= melhor[1] - 4 and a[0] != melhor[0]]
        if rivais:
            return None
        iguais = [c for c in self.clientes if c.chave == melhor[0]]
        if len(iguais) != 1:
            return None
        return Palpite(iguais[0].pasta, como, int(melhor[1]))

    def _pelo_dominio(self, dom: str) -> Palpite | None:
        """'acmeusinagem.com.br' -> o único cliente cuja 1ª palavra (>= 5 letras) está no domínio."""
        rotulo = re.sub(r"[^a-z0-9]", "", dom.split(".")[0])
        if len(rotulo) < 5:
            return None
        candidatos = []
        for c in self.clientes:
            palavras = c.chave.split()
            if not palavras:
                continue
            junto = "".join(palavras).lower()
            primeira = palavras[0].lower()
            if (len(primeira) >= 5 and primeira in rotulo) or (len(junto) >= 6 and (junto in rotulo or rotulo in junto)):
                candidatos.append(c)
        if len(candidatos) == 1:
            return Palpite(candidatos[0].pasta, "domínio", 80)
        return None


# ------------------------------------------------------------------ destino

def _so_letras(nome: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", sem_acento(nome).upper())


def _filho(pai: Path, escolher) -> Path | None:
    if not pai.is_dir():
        return None
    filhos = sorted((p for p in pai.iterdir() if p.is_dir()), key=lambda p: p.name.upper())
    return escolher(filhos)


def _nfe(filhos):
    return next((p for p in filhos if _so_letras(p.name) in ("NFE", "NFES")), None)


def _mes(comp: Competencia):
    ano, mes = comp
    alvo = f"{mes:02d}{ano}"

    def escolher(filhos):
        return next((p for p in filhos
                     if re.fullmatch(r"\s*\d{2}\s*[-.]?\s*\d{4}\s*", p.name) and re.sub(r"\D", "", p.name) == alvo),
                    None)
    return escolher


def _tomados(filhos):
    com = [p for p in filhos if "TOMAD" in _so_letras(p.name)]
    exato = [p for p in com if _so_letras(p.name) == "TOMADOS"]
    return (exato or com or [None])[0]


@dataclass
class Destino:
    pasta: Path
    criar: list[str]    # níveis que ainda não existem e serão criados

    @property
    def descricao(self) -> str:
        return str(self.pasta) + (f"  (cria: {' > '.join(self.criar)})" if self.criar else "")


def destino(raiz: str | Path, pasta_cliente: str, comp: Competencia) -> Destino:
    """Onde salvar, sem criar nada. Use `garantir()` na hora de gravar."""
    atual = Path(raiz) / pasta_cliente
    if not atual.is_dir():
        raise FileNotFoundError(f"Pasta do cliente não existe: {atual}")
    criar: list[str] = []
    passos = [(_nfe, "NFE"), (lambda f: next((p for p in f if p.name.strip() == str(comp[0])), None), str(comp[0])),
              (_mes(comp), f"{comp[1]:02d}-{comp[0]}"), (_tomados, "Tomados"),
              (lambda f: next((p for p in f if _so_letras(p.name) == _so_letras(PASTA_FINAL)), None), PASTA_FINAL)]
    for escolher, nome_novo in passos:
        achado = None if criar else _filho(atual, escolher)
        if achado is None:
            criar.append(nome_novo)
            atual = atual / nome_novo
        else:
            atual = achado
    return Destino(atual, criar)


def garantir(d: Destino) -> Path:
    d.pasta.mkdir(parents=True, exist_ok=True)
    return d.pasta
