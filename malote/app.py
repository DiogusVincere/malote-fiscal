"""Malote Fiscal — a tela.

1 · Perguntar: manda, pela conta do Outlook escolhida, o e-mail "Serviços tomados - <Mês> de <ano>" para
    os clientes da planilha e dos grupos marcados, em Cco, em lotes, com o texto do último envio (ou o
    texto padrão, na primeira vez).
2 · Respostas e notas: lê tudo o que chegou de fora para a competência (Caixa de entrada, subpastas e
    lixo eletrônico), descobre de qual cliente é cada e-mail e salva os anexos em
    <cliente>\\NFE\\<ano>\\<mês>\\Tomados\\Tomados e-mail.
3 · Responder: sugere a resposta para cada cliente; você corrige e aceita, e ela aprende.
4 · Cadastro: quem recebe e de qual empresa é cada e-mail; a planilha atualiza, você corrige.

Tudo roda no Outlook desta máquina; nada sai daqui além dos e-mails que você mandar.
`--demo` abre a mesma tela sobre um Outlook simulado e um escritório fictício (malote/falso.py).
"""
from __future__ import annotations

import argparse
import html as html_lib
import os
import queue
import re
import sys
import tempfile
import threading
import traceback
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import VERSAO, anexos, controle, falso, outlook, planilha, sugestao, tema
from . import cadastro as cadastro_mod
from . import clientes as K
from . import competencia as C
from . import layouts as L
from .registro import Config, Registro
from .tema import AMBAR, PAINEL, TINTA

TITULO = "Malote Fiscal"
DOWNLOADS = Path(os.path.expanduser("~")) / "Downloads"
GRUPO_PADRAO, TAMANHO_LOTE = outlook.GRUPO_PADRAO, outlook.TAMANHO_LOTE
LARGURA_TEXTO = 540   # quebra de linha dos textos do cartão da aba 1 (a tabela da esquerda tem largura mínima)


def pasta_de_dados_padrao() -> Path:
    """Onde ficam config, cadastro, aprendizado e planilhas de controle: %LOCALAPPDATA%\\Malote Fiscal."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_DATA_HOME") or str(Path.home())
    return Path(base) / "Malote Fiscal"


def argumentos(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(prog="malote", description="Malote Fiscal — serviços tomados por e-mail.")
    p.add_argument("--demo", action="store_true", help="abre com o Outlook simulado e clientes fictícios")
    p.add_argument("--dados", help="pasta de dados (padrão: %%LOCALAPPDATA%%\\Malote Fiscal)")
    return p.parse_args(argv)


class App(tk.Tk):
    def __init__(self, args: argparse.Namespace | None = None):
        super().__init__()
        self.args = args or argumentos([])
        self.title(f"{TITULO} — serviços tomados por e-mail")
        self.geometry("1320x900")
        self.minsize(1100, 720)
        self.configure(bg=tema.FUNDO)
        tema.estilos(self)
        s = ttk.Style(self)
        s.configure("Painel.TCheckbutton", background=PAINEL, foreground=TINTA)
        s.map("Painel.TCheckbutton", background=[("active", PAINEL)])

        self.pasta_dados_real = Path(self.args.dados) if self.args.dados else pasta_de_dados_padrao()
        self.demo: falso.Demonstracao | None = None
        self._abrir_dados(self.pasta_dados_real)
        self._por_codigo: dict[str, planilha.Empresa] = {}
        self.aviso_cadastro = ""
        self.aviso_outlook = ""
        self.mudanca: tuple | None = None     # aba 3: ("tirar", e-mail) ou ("trocar", antigo, novo)
        self.grupos: list[outlook.Grupo] = []
        self.conta: str = self.config_.get("conta") or ""
        self.dominio: str = self.conta.rsplit("@", 1)[-1].lower() if "@" in self.conta else ""
        self.contas: list[outlook.Conta] = []
        self.marcar_reenvio: dict[str, bool] = {}   # aba "Não entregues": quem vai no reenvio
        self.fila: queue.Queue = queue.Queue()
        self.trabalhando = False

        self.caminho_planilha: Path | None = None
        self.empresas: list[planilha.Empresa] = []
        self.mapa: dict[str, list[planilha.Empresa]] = {}
        self.lista: planilha.ListaDeEnvio | None = None
        self.envios: dict = {}
        self.modelo: outlook.Modelo | None = None
        self.nomes: dict[str, str] = {}
        self.clientes: list[K.Cliente] = []
        self.ident: K.Identificador | None = None
        self.respostas: list[outlook.Resposta] = []
        self.palpites: dict[str, K.Palpite] = {}
        self.sugeridos: dict[str, dict[int, str | None]] = {}    # CNPJ achado no anexo
        self.atribuicoes: dict[str, dict[int, str]] = {}         # escolhido na tela, anexo a anexo
        self.sem_resposta: list[tuple[planilha.Empresa, list[str]]] = []
        self.atual: outlook.Resposta | None = None
        self.sugestao_atual: sugestao.Sugestao | None = None

        comps = C.ultimas(13)
        self.var_comp = tk.StringVar(value=C.texto(comps[0]))
        self.var_status = tk.StringVar(value="Abrindo o Outlook…")
        self.var_planilha = tk.StringVar(value="Planilha de clientes: procurando…")
        self.var_raiz = tk.StringVar(value="Pasta dos clientes: —")
        self.var_resumo_envio = tk.StringVar()
        self.var_aviso_envio = tk.StringVar()
        self.var_porque = tk.StringVar()
        self.var_saudacao = tk.StringVar()
        self.var_aprender = tk.BooleanVar(value=True)
        self.var_tirar = tk.BooleanVar(value=False)        # aplicar a mudança proposta no cadastro
        self.var_texto_mudanca = tk.StringVar(value="Tirar este e-mail da lista de envio")
        self.var_busca_cad = tk.StringVar()
        self.var_filtro_cad = tk.StringVar(value="Todos")
        self.var_info_cad = tk.StringVar()
        self.var_mostrar_respondidos = tk.BooleanVar(value=False)
        self.var_cabecalho = tk.StringVar()
        self.var_filtro = tk.StringVar(value="Todos")
        self.var_layout = tk.StringVar()
        self.var_assunto = tk.StringVar()
        self.var_pronta = tk.StringVar()
        self.var_conta = tk.StringVar(value=self.conta)
        self.var_mostrar = tk.StringVar(value="Tudo")
        self.var_demo = tk.StringVar()

        self._montar([C.texto(c) for c in comps])
        self.after(200, self._listar_contas)

    # ================================================================ dados (pasta real ou da demonstração)
    def _abrir_dados(self, pasta: Path):
        pasta.mkdir(parents=True, exist_ok=True)
        self.pasta_dados = pasta
        self.config_ = Config(pasta)
        self.registro = Registro(pasta)
        self.base = sugestao.Base(pasta)
        self.layouts = L.Layouts(pasta)
        self.cadastro = cadastro_mod.Cadastro(pasta)
        self.cadastro.migrar(self.registro.fora_da_lista, self.registro.incluir)
        self.marcados: set[str] = set(self.config_.get("grupos", []))

    @property
    def pasta_controle(self) -> Path:
        return self.pasta_dados / "controle"

    # ================================================================ tela
    def _montar(self, competencias: list[str]):
        tema.faixa(self, TITULO,
                   "Pergunta aos clientes pela conta escolhida, guarda as notas na pasta de cada um e ajuda a responder",
                   f"versão {VERSAO}")
        rodape = tema.barra_status(self, self.var_status)
        self.progresso = ttk.Progressbar(rodape, mode="indeterminate", length=180)

        topo = ttk.Frame(self, padding=(20, 12, 20, 0))
        topo.pack(fill="x")
        ttk.Label(topo, text="Competência (mês das notas)", style="Passo.TLabel").pack(side="left")
        cb = ttk.Combobox(topo, textvariable=self.var_comp, values=competencias, width=10, state="readonly")
        cb.pack(side="left", padx=(10, 0))
        cb.bind("<<ComboboxSelected>>", lambda _e: self._mudou_competencia())
        ttk.Label(topo, text="Conta do Outlook", style="Passo.TLabel").pack(side="left", padx=(24, 0))
        self.cb_conta = ttk.Combobox(topo, textvariable=self.var_conta, values=[self.conta] if self.conta else [],
                                     width=40, state="readonly")
        self.cb_conta.pack(side="left", padx=(10, 0))
        self.cb_conta.bind("<<ComboboxSelected>>", lambda _e: self._mudou_conta())
        ttk.Button(topo, text="Trocar planilha…", command=self._trocar_planilha).pack(side="right")
        ttk.Button(topo, text="Pasta dos clientes…", command=self._escolher_raiz).pack(side="right", padx=6)

        info = ttk.Frame(self, padding=(20, 4, 20, 0))
        info.pack(fill="x")
        self.lb_demo = tk.Label(info, textvariable=self.var_demo, bg=tema.AMBAR_T, fg=AMBAR,
                                font=(tema.FONTE, 9, "bold"), padx=10, pady=2)
        self.lb_demo.pack(side="left", padx=(0, 16))
        self.lb_demo.pack_forget()
        ttk.Label(info, textvariable=self.var_planilha, style="Apagado.TLabel").pack(side="left")
        ttk.Label(info, textvariable=self.var_raiz, style="Apagado.TLabel").pack(side="left", padx=(24, 0))

        self.abas = ttk.Notebook(self)
        self.abas.pack(fill="both", expand=True, padx=20, pady=12)
        self._aba_perguntar()
        self._aba_respostas()
        self._aba_responder()
        self._aba_cadastro()

    # ---------------------------------------------------------------- aba 1
    def _aba_perguntar(self):
        aba = ttk.Frame(self.abas, padding=14)
        self.abas.add(aba, text="1 · Perguntar aos clientes")
        aba.columnconfigure(0, weight=5)
        aba.columnconfigure(1, weight=6)
        aba.rowconfigure(3, weight=1)

        # ---- esquerda: quem recebe (planilha + grupos)
        ttk.Label(aba, text="Grupos do Outlook somados à planilha (clique para marcar)", style="Passo.TLabel").grid(
            row=0, column=0, sticky="w")
        self.tab_grupos = tema.Tabela(aba, [("marca", "", 30, "center", False), ("grupo", "Grupo", 260, "w", True),
                                            ("qtd", "E-mails", 70, "e", False)])
        self.tab_grupos.arvore.configure(height=5)
        self.tab_grupos.grid(row=1, column=0, sticky="nsew", padx=(0, 14), pady=(2, 8))
        self.tab_grupos.arvore.bind("<ButtonRelease-1>", self._clicou_grupo)

        filtro = ttk.Frame(aba)
        filtro.grid(row=2, column=0, sticky="ew", padx=(0, 14))
        ttk.Label(filtro, text="Quem recebe", style="Passo.TLabel").pack(side="left")
        cb = ttk.Combobox(filtro, textvariable=self.var_filtro, state="readonly", width=24,
                          values=["Todos", "Recebem", "Não recebem", "Só nos grupos", "Só na planilha",
                                  "Planilha + grupo"])
        cb.pack(side="right")
        cb.bind("<<ComboboxSelected>>", lambda _e: self._preencher_destinatarios())
        ttk.Label(filtro, text="mostrar:", style="Apagado.TLabel").pack(side="right", padx=4)
        self.tab_dest = tema.Tabela(aba, [("marca", "", 30, "center", False), ("email", "E-mail", 200, "w", True),
                                          ("origem", "De onde", 105, "w", False),
                                          ("emp", "Empresa(s)", 180, "w", True), ("obs", "Observação", 150, "w", True)])
        self.tab_dest.grid(row=3, column=0, sticky="nsew", padx=(0, 14))
        self.tab_dest.arvore.tag_configure("fora", foreground=tema.APAGADO)
        self.tab_dest.arvore.bind("<ButtonRelease-1>", self._clicou_destinatario)

        # ---- direita: a mensagem
        lado = ttk.Frame(aba, style="Cartao.TFrame", padding=14)
        lado.grid(row=0, column=1, rowspan=4, sticky="nsew")
        lado.columnconfigure(1, weight=1)
        lado.rowconfigure(5, weight=1)
        ttk.Label(lado, textvariable=self.var_resumo_envio, style="Painel.TLabel", justify="left",
                  wraplength=LARGURA_TEXTO).grid(row=0, column=0, columnspan=2, sticky="w")
        self.lb_aviso = tk.Label(lado, textvariable=self.var_aviso_envio, bg=tema.AMBAR_T, fg=AMBAR,
                                 font=(tema.FONTE, 9, "bold"), justify="left", wraplength=LARGURA_TEXTO - 20,
                                 padx=10, pady=5, anchor="w")
        self.lb_aviso.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.lb_aviso.grid_remove()

        linha = ttk.Frame(lado, style="Painel.TFrame")
        linha.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 6))
        ttk.Label(linha, text="Layout", style="Painel.TLabel").pack(side="left")
        self.cb_layout = ttk.Combobox(linha, textvariable=self.var_layout, state="readonly", width=30)
        self.cb_layout.pack(side="left", padx=6)
        self.cb_layout.bind("<<ComboboxSelected>>", lambda _e: self._carregar_layout())
        ttk.Button(linha, text="Salvar", command=self._salvar_layout).pack(side="left")
        ttk.Button(linha, text="Salvar como…", command=lambda: self._salvar_layout(novo=True)).pack(side="left", padx=4)
        ttk.Button(linha, text="Excluir", command=self._excluir_layout).pack(side="left")

        ttk.Label(lado, text="Assunto", style="Painel.TLabel").grid(row=3, column=0, sticky="w")
        ttk.Entry(lado, textvariable=self.var_assunto).grid(row=3, column=1, sticky="ew", padx=(8, 0))
        ttk.Label(lado, text="Texto", style="Painel.TLabel").grid(row=4, column=0, sticky="nw", pady=(8, 0))
        self.txt_mensagem = tk.Text(lado, height=14, wrap="word", font=(tema.FONTE, 10), relief="solid",
                                    borderwidth=1, padx=10, pady=8, undo=True)
        self.txt_mensagem.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(2, 4))
        ajuda = "  ·  ".join(f"{c} = {d}" for c, d in L.CAMPOS[:4])
        ttk.Label(lado, text=f"{ajuda}\nLinha em branco separa parágrafo · “- ” vira item de lista · **texto** sai em "
                             "negrito · a assinatura do último envio entra sozinha no fim.",
                  style="ApagadoPainel.TLabel", wraplength=LARGURA_TEXTO, justify="left").grid(
            row=6, column=0, columnspan=2, sticky="w")

        botoes = ttk.Frame(aba)
        botoes.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(botoes, text="Conferência da planilha (Excel)", command=self._conferencia).pack(side="left")
        ttk.Button(botoes, text="Ver como fica", command=self._ver_como_fica).pack(side="left", padx=6)
        ttk.Button(botoes, text="Enviar agora", style="Primario.TButton",
                   command=lambda: self._enviar(True)).pack(side="right")
        ttk.Button(botoes, text="Criar rascunhos no Outlook", command=lambda: self._enviar(False)).pack(
            side="right", padx=6)

    # ---------------------------------------------------------------- aba 2
    def _aba_respostas(self):
        aba = ttk.Frame(self.abas, padding=14)
        self.abas.add(aba, text="2 · Respostas e notas")
        barra = ttk.Frame(aba)
        barra.pack(fill="x")
        ttk.Button(barra, text="Buscar respostas", style="Primario.TButton", command=self._buscar).pack(side="left")
        ttk.Button(barra, text="Escolher o cliente…", command=self._escolher).pack(side="left", padx=(10, 0))
        ttk.Button(barra, text="Confirmar todas as sugestões", command=self._confirmar_sugestoes).pack(
            side="left", padx=6)
        ttk.Label(barra, text="mostrar:", style="Apagado.TLabel").pack(side="left", padx=(16, 4))
        cb_mostrar = ttk.Combobox(barra, textvariable=self.var_mostrar, state="readonly", width=24,
                                  values=["Tudo", "Respostas à pergunta", "Contatos do cadastro", "Outros assuntos",
                                          "Com anexo", "Desta competência", "Outra competência", "Lixo eletrônico"])
        cb_mostrar.pack(side="left")
        cb_mostrar.bind("<<ComboboxSelected>>", lambda _e: self._preencher_respostas())
        ttk.Button(barra, text="Planilha de controle", command=self._planilha_controle).pack(side="right")
        ttk.Button(barra, text="Salvar anexos nas pastas", style="Primario.TButton", command=self._salvar).pack(
            side="right", padx=8)

        self.sub = ttk.Notebook(aba)
        self.sub.pack(fill="both", expand=True, pady=(10, 0))
        f1, f2, f3 = (ttk.Frame(self.sub, style="Painel.TFrame") for _ in range(3))
        self.sub.add(f1, text="Respostas")
        self.sub.add(f2, text="Sem resposta")
        self.sub.add(f3, text="Não entregues")

        self.tab_resp = tema.Tabela(f1, [
            ("rec", "Recebido", 100, "w", False), ("nome", "Remetente", 180, "w", False),
            ("cli", "Cliente", 290, "w", True), ("como", "Como achou", 115, "w", False),
            ("rel", "Por quê", 130, "w", False), ("tipo", "Situação", 100, "w", False),
            ("anx", "Anexos", 55, "center", False), ("comp", "Comp.", 60, "center", False),
            ("pasta", "Salvo", 55, "w", False), ("resp", "Respondido", 75, "center", False)])
        self.tab_resp.pack(fill="both", expand=True)
        self.tab_resp.arvore.tag_configure("ambar", background=tema.AMBAR_T)
        self.tab_resp.arvore.tag_configure("outra", foreground=tema.APAGADO)
        self.tab_resp.arvore.bind("<<TreeviewSelect>>", lambda _e: self._mostrar_detalhe())
        self.tab_resp.arvore.bind("<Double-1>", lambda _e: self._escolher())
        self.txt_detalhe = tk.Text(f1, height=8, wrap="word", font=(tema.FONTE, 9), bg=tema.ZEBRA, relief="flat",
                                   padx=10, pady=6, state="disabled")
        self.txt_detalhe.pack(fill="x", pady=(8, 0))
        ttk.Label(f1, text="Verde: salvo.  Amarelo: cliente só sugerido ou várias empresas — confira e confirme "
                           "(dois cliques).  Vermelho: mandou anexo e não sei de quem é.  Cinza: outro assunto ou "
                           "outra competência (mostrado para você decidir; nada é escondido).",
                  style="ApagadoPainel.TLabel").pack(anchor="w", pady=(4, 0))

        self.tab_sem = tema.Tabela(f2, [("cli", "Empresa", 420, "w", True), ("fiscal", "Resp. fiscal", 200, "w", False),
                                        ("email", "E-mails perguntados", 420, "w", True)])
        self.tab_sem.pack(fill="both", expand=True)
        ttk.Label(f2, text="Empresas que receberam a pergunta deste mês e não responderam por nenhum dos e-mails.",
                  style="ApagadoPainel.TLabel").pack(anchor="w", pady=6)

        self.tab_dev = tema.Tabela(f3, [("marca", "", 30, "center", False), ("rec", "Recebido", 95, "w", False),
                                        ("email", "E-mail que voltou", 250, "w", True),
                                        ("motivo", "Motivo", 280, "w", True), ("cli", "Cliente", 280, "w", True),
                                        ("sit", "Situação", 170, "w", False)])
        self.tab_dev.pack(fill="both", expand=True)
        self.tab_dev.arvore.configure(selectmode="extended")
        self.tab_dev.arvore.bind("<ButtonRelease-1>", self._clicou_devolvido)
        b3 = ttk.Frame(f3, style="Painel.TFrame")
        b3.pack(fill="x", pady=6)
        ttk.Button(b3, text="Reenviar só para os marcados", style="Primario.TButton",
                   command=self._reenviar).pack(side="left")
        ttk.Button(b3, text="Tirar da lista de envio", command=self._tirar_devolvido).pack(side="left", padx=(12, 0))
        ttk.Button(b3, text="Voltar para a lista", command=self._voltar_devolvido).pack(side="left", padx=6)
        ttk.Label(f3, text="✔ = vai no reenvio (mesmo texto e assunto da aba 1). Já vem marcado o que vale tentar de novo "
                           "(caixa cheia, recusa por spam, falha temporária); endereço que não existe vem desmarcado — "
                           "esse é melhor tirar da lista e corrigir na planilha.",
                  style="ApagadoPainel.TLabel", wraplength=1150, justify="left").pack(anchor="w")

    # ---------------------------------------------------------------- aba 3
    def _aba_responder(self):
        aba = ttk.Frame(self.abas, padding=14)
        self.abas.add(aba, text="3 · Responder")
        aba.columnconfigure(1, weight=1)
        aba.rowconfigure(1, weight=1)

        cab = ttk.Frame(aba)
        cab.grid(row=0, column=0, sticky="ew")
        ttk.Label(cab, text="A responder", style="Passo.TLabel").pack(side="left")
        ttk.Checkbutton(cab, text="mostrar já respondidas", variable=self.var_mostrar_respondidos,
                        command=self._preencher_responder).pack(side="right")
        self.tab_fila = tema.Tabela(aba, [("rec", "Recebido", 95, "w", False),
                                          ("quem", "Cliente / remetente", 250, "w", True),
                                          ("cat", "Assunto", 110, "w", False)])
        self.tab_fila.grid(row=1, column=0, sticky="nsew", padx=(0, 14))
        self.tab_fila.arvore.tag_configure("ambar", background=tema.AMBAR_T)
        self.tab_fila.arvore.bind("<<TreeviewSelect>>", lambda _e: self._abrir_resposta())

        lado = ttk.Frame(aba, style="Cartao.TFrame", padding=16)
        lado.grid(row=0, column=1, rowspan=3, sticky="nsew")
        lado.columnconfigure(0, weight=1)
        ttk.Label(lado, textvariable=self.var_cabecalho, style="Titulo.TLabel", wraplength=760).grid(
            row=0, column=0, sticky="w")
        ttk.Label(lado, text="O que o cliente escreveu", style="ApagadoPainel.TLabel").grid(
            row=1, column=0, sticky="w", pady=(8, 2))
        self.txt_cliente = tk.Text(lado, height=10, wrap="word", font=(tema.FONTE, 10), bg=tema.ZEBRA, relief="flat",
                                   padx=10, pady=8, state="disabled")
        self.txt_cliente.grid(row=2, column=0, sticky="nsew")
        lado.rowconfigure(2, weight=1)

        ttk.Label(lado, textvariable=self.var_porque, style="ApagadoPainel.TLabel", wraplength=760,
                  justify="left").grid(row=3, column=0, sticky="w", pady=(10, 4))
        linha = ttk.Frame(lado, style="Painel.TFrame")
        linha.grid(row=4, column=0, sticky="ew")
        ttk.Label(linha, text="Saudação", style="Painel.TLabel").pack(side="left")
        ttk.Entry(linha, textvariable=self.var_saudacao, width=34).pack(side="left", padx=8)
        ttk.Label(linha, text=f"+ “{sugestao.ABERTURA}”", style="ApagadoPainel.TLabel").pack(side="left")
        cab_resp = ttk.Frame(lado, style="Painel.TFrame")
        cab_resp.grid(row=5, column=0, sticky="ew", pady=(8, 2))
        ttk.Label(cab_resp, text="Resposta", style="Painel.TLabel").pack(side="left")
        ttk.Button(cab_resp, text="Excluir", command=self._excluir_pronta).pack(side="right")
        ttk.Button(cab_resp, text="Guardar como pronta…", command=self._salvar_pronta).pack(side="right", padx=4)
        ttk.Button(cab_resp, text="Usar", command=self._usar_pronta).pack(side="right")
        self.cb_pronta = ttk.Combobox(cab_resp, textvariable=self.var_pronta, state="readonly", width=30)
        self.cb_pronta.pack(side="right", padx=6)
        self.cb_pronta.bind("<<ComboboxSelected>>", lambda _e: self._usar_pronta())
        ttk.Label(cab_resp, text="Respostas prontas:", style="ApagadoPainel.TLabel").pack(side="right")
        self.txt_resposta = tk.Text(lado, height=6, wrap="word", font=(tema.FONTE, 10), relief="solid",
                                    borderwidth=1, padx=10, pady=8)
        self.txt_resposta.grid(row=6, column=0, sticky="nsew")
        ttk.Label(lado, text=f"+ “{sugestao.FECHO}” e a assinatura da conta", style="ApagadoPainel.TLabel").grid(
            row=7, column=0, sticky="w", pady=(2, 6))

        opcoes = ttk.Frame(lado, style="Painel.TFrame")
        opcoes.grid(row=8, column=0, sticky="w")
        ttk.Checkbutton(opcoes, text="Aprender: usar esta resposta quando chegar mensagem parecida",
                        variable=self.var_aprender, style="Painel.TCheckbutton").pack(anchor="w")
        self.ck_mudanca = ttk.Checkbutton(opcoes, textvariable=self.var_texto_mudanca, variable=self.var_tirar,
                                          style="Painel.TCheckbutton")
        self.ck_mudanca.pack(anchor="w")

        botoes = ttk.Frame(lado, style="Painel.TFrame")
        botoes.grid(row=9, column=0, sticky="e", pady=(10, 0))
        ttk.Button(botoes, text="Respostas aprendidas…", command=lambda: BaseDeRespostas(self, self.base)).pack(
            side="left", padx=(0, 24))
        ttk.Button(botoes, text="Pular", command=self._proxima).pack(side="left")
        ttk.Button(botoes, text="Só ensinar (não envia)", command=lambda: self._aceitar(False)).pack(side="left", padx=6)
        ttk.Button(botoes, text="Aceitar e enviar", style="Primario.TButton",
                   command=lambda: self._aceitar(True)).pack(side="left")

    # ---------------------------------------------------------------- aba 4
    def _aba_cadastro(self):
        aba = ttk.Frame(self.abas, padding=14)
        self.abas.add(aba, text="4 · Cadastro")
        barra = ttk.Frame(aba)
        barra.pack(fill="x")
        ttk.Label(barra, text="Procurar", style="Passo.TLabel").pack(side="left")
        busca = ttk.Entry(barra, textvariable=self.var_busca_cad, width=32)
        busca.pack(side="left", padx=8)
        self.var_busca_cad.trace_add("write", lambda *_: self._preencher_cadastro())
        cb = ttk.Combobox(barra, textvariable=self.var_filtro_cad, state="readonly", width=22,
                          values=["Todos", "Na lista", "Fora da lista", "Responderam", "Principais",
                                  "Acrescentados à mão", "Sumiram da planilha", "Mudados por pedido"])
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _e: self._preencher_cadastro())
        ttk.Button(barra, text="Acrescentar…", style="Primario.TButton", command=self._acrescentar).pack(side="right")
        ttk.Button(barra, text="Trocar e-mail…", command=self._trocar_email).pack(side="right", padx=6)
        ttk.Button(barra, text="★ Principal", command=self._principal).pack(side="right")
        ttk.Button(barra, text="Voltar para a lista", command=lambda: self._mudar_cad(True)).pack(side="right")
        ttk.Button(barra, text="Tirar da lista", command=lambda: self._mudar_cad(False)).pack(side="right", padx=6)

        self.tab_cad = tema.Tabela(aba, [("email", "E-mail", 250, "w", True), ("emp", "Empresa(s)", 310, "w", True),
                                         ("sit", "Situação", 200, "w", True), ("resp", "Respondeu", 85, "center", False),
                                         ("orig", "Origem", 150, "w", False), ("ult", "Última mudança", 240, "w", True)])
        self.tab_cad.pack(fill="both", expand=True, pady=(10, 0))
        self.tab_cad.arvore.configure(selectmode="extended")
        self.tab_cad.arvore.tag_configure("fora", foreground=tema.APAGADO)
        self.tab_cad.arvore.tag_configure("ambar", background=tema.AMBAR_T)
        self.tab_cad.arvore.bind("<<TreeviewSelect>>", lambda _e: self._mostrar_historico())
        self.txt_hist = tk.Text(aba, height=6, wrap="word", font=(tema.FONTE, 9), bg=tema.ZEBRA, relief="flat",
                                padx=10, pady=6, state="disabled")
        self.txt_hist.pack(fill="x", pady=(8, 0))
        ttk.Label(aba, textvariable=self.var_info_cad, style="Apagado.TLabel").pack(anchor="w", pady=(4, 0))

    # ================================================================ segundo plano
    def _abrir_outlook(self):
        """O Outlook de verdade na conta escolhida, ou o simulado da demonstração."""
        return self.demo.abrir() if self.demo else outlook.Outlook(self.conta)

    def _rodar(self, aviso: str, trabalho, depois, sem_outlook: bool = False):
        """trabalho(ol, avisar) roda numa thread com o Outlook aberto na conta escolhida; depois(resultado)
        volta na tela. sem_outlook: trabalho(None, avisar), para o que não precisa de uma conta."""
        if self.trabalhando:
            return
        self.trabalhando = True
        self.configure(cursor="watch")
        self.progresso.pack(side="right", padx=10, pady=4)
        self.progresso.start(12)
        self.var_status.set(aviso)

        def avisar(texto: str):
            self.fila.put(("status", texto))

        def alvo():
            try:
                if sem_outlook:
                    resultado = trabalho(None, avisar)
                else:
                    with self._abrir_outlook() as ol:
                        resultado = trabalho(ol, avisar)
                self.fila.put(("fim", (depois, resultado)))
            except Exception as erro:
                with (self.pasta_dados / "erros.log").open("a", encoding="utf-8") as log:
                    log.write(f"\n[{datetime.now():%d/%m/%Y %H:%M}]\n{traceback.format_exc()}")
                self.fila.put(("erro", str(erro) or erro.__class__.__name__))

        threading.Thread(target=alvo, daemon=True).start()
        self.after(80, self._ler_fila)

    def _ler_fila(self):
        try:
            while True:
                tipo, dados = self.fila.get_nowait()
                if tipo == "status":
                    self.var_status.set(dados)
                    continue
                self.trabalhando = False
                self.configure(cursor="")
                self.progresso.stop()
                self.progresso.pack_forget()
                if tipo == "erro":
                    self.var_status.set("Deu erro: " + dados)
                    messagebox.showerror(TITULO, "Deu erro:\n\n" + dados +
                                         f"\n\nDetalhes em {self.pasta_dados / 'erros.log'}")
                else:
                    depois, resultado = dados
                    depois(resultado)
                return
        except queue.Empty:
            pass
        self.after(80, self._ler_fila)

    # ================================================================ conta do Outlook
    def _listar_contas(self):
        def trabalho(_ol, avisar):
            avisar("Lendo as contas do Outlook…")
            try:
                return outlook.contas(), ""
            except Exception as erro:  # sem pywin32, sem Outlook ou Outlook travado: a demonstração continua valendo
                return [], str(erro) or erro.__class__.__name__

        def depois(res):
            contas, erro = res
            self.contas = contas
            enderecos = [c.endereco for c in contas] + [falso.ROTULO]
            self.cb_conta.configure(values=enderecos)
            self.aviso_outlook = f"Outlook indisponível nesta máquina ({erro}). " if erro else ""
            if self.args.demo or not contas:
                escolha = falso.ROTULO
            elif self.conta in enderecos:
                escolha = self.conta
            else:
                escolha = enderecos[0]
            self._definir_conta(escolha, carregar=True)

        self._rodar("Abrindo o Outlook…", trabalho, depois, sem_outlook=True)

    def _definir_conta(self, conta: str, carregar: bool):
        if conta == falso.ROTULO:
            if self.demo is None:
                self.demo = falso.Demonstracao()
                self.demo.preparar(limpar=True)
            self._abrir_dados(self.demo.dados)
            self.conta = falso.CONTA
            self.var_demo.set("DEMONSTRAÇÃO · Outlook simulado · nada sai da máquina")
            self.lb_demo.pack(side="left", padx=(0, 16), before=self.lb_demo.master.winfo_children()[1])
        else:
            self.demo = None
            self._abrir_dados(self.pasta_dados_real)
            self.conta = conta
            self.config_.set("conta", conta)
            self.lb_demo.pack_forget()
        self.dominio = self.conta.rsplit("@", 1)[-1].lower()
        self.var_conta.set(conta)
        self._atualizar_raiz_na_tela()
        if carregar:
            self._carregar()

    def _conta_na_tela(self) -> str:
        return falso.ROTULO if self.demo else self.conta

    def _mudou_conta(self):
        nova = self.var_conta.get()
        if nova == self._conta_na_tela() or self.trabalhando:
            self.var_conta.set(self._conta_na_tela())
            return
        self.respostas, self.palpites, self.sem_resposta = [], {}, []
        self.sugeridos, self.atribuicoes, self.envios, self.modelo = {}, {}, {}, None
        for t in (self.tab_resp, self.tab_sem, self.tab_dev, self.tab_fila):
            t.limpar()
        self._definir_conta(nova, carregar=True)

    # ================================================================ carregar
    def _pastas_da_planilha(self, guardada: str | None) -> list:
        if self.demo:
            return [self.demo.pasta]
        return [DOWNLOADS, self.pasta_dados, Path(guardada).parent if guardada else None]

    def _carregar(self):
        raiz = self._raiz()
        semear = not self.base.exemplos
        guardada = self.config_.get("planilha")
        caminho = (Path(guardada) if guardada and Path(guardada).is_file() else None)
        mais_nova = planilha.achar(self._pastas_da_planilha(guardada))
        if mais_nova and (caminho is None or mais_nova.name > caminho.name):
            caminho = mais_nova  # chegou planilha mais nova na pasta: usa ela
        lidas_ate = self.cadastro.respostas_lidas_ate

        def trabalho(ol, avisar):
            avisar("Lendo o último envio (modelo do texto)…")
            modelo = ol.modelo()
            avisar("Lendo os envios anteriores…")
            envios = ol.envios()
            agora = datetime.now()
            if lidas_ate:
                desde = datetime.strptime(lidas_ate, "%Y-%m-%d %H:%M")
            else:  # primeira vez: desde o envio mais antigo que existe em Itens Enviados
                desde = min((e.enviado_em for lista in envios.values() for e in lista), default=agora)
            avisar("Vendo quem costuma responder em cada empresa…")
            quem_respondeu = (ol.quem_respondeu(desde, envios), agora)
            avisar("Lendo os grupos de contatos…")
            grupos = ol.grupos()
            avisar("Lendo os contatos do Outlook (para reconhecer quem não está na planilha)…")
            nomes = ol.nomes_dos_contatos()
            antigas = []
            if semear:
                avisar("Aprendendo com as respostas que a equipe já mandou…")
                antigas = ol.respostas_antigas()
            avisar("Lendo as pastas dos clientes e a planilha…")
            clientes = K.indexar(raiz) if raiz else []
            empresas = planilha.ler(caminho) if caminho else []
            return modelo, envios, grupos, nomes, antigas, clientes, empresas, quem_respondeu

        def depois(res):
            self.modelo, self.envios, self.grupos, self.nomes, antigas, self.clientes, empresas, quem = res
            if not self.marcados:
                self.marcados = {g.nome for g in self.grupos if GRUPO_PADRAO.match(g.nome)}
            for g in self.grupos:
                for nome, email in g.membros:
                    self.nomes.setdefault(email.lower(), nome)
            self.layouts.garantir_padrao(self.modelo.html, self.modelo.competencia,
                                         texto=None if self.modelo.de_envio else L.TEXTO_PADRAO)
            self._preencher_grupos()
            self._atualizar_layouts(self.layouts.pergunta_atual)
            self._atualizar_prontas()
            for lista in self.envios.values():
                for envio in lista:
                    for nome, email in envio.destinatarios:
                        self.nomes.setdefault(email.lower(), nome)
            for cliente_msg, nosso in antigas:
                self.base.ensinar(cliente_msg, sugestao.miolo_de_resposta_antiga(nosso),
                                  sugestao.categoria(cliente_msg, "Ler"), "histórico")
            self._usar_planilha(caminho, empresas)
            self.cadastro.anotar_respostas(*quem)   # depois do cadastro existir: quem responde em cada empresa
            self._atualizar_envio()
            self._preencher_cadastro()
            self._atualizar_raiz_na_tela()
            extra = f" Aprendi {len(antigas)} respostas antigas da equipe." if antigas else ""
            if not self.modelo.de_envio:
                extra += (" Esta conta nunca mandou a pergunta: o texto da aba 1 é o padrão da ferramenta — "
                          "ajuste e salve.")
            self.var_status.set(f"{self.aviso_outlook}Pronto. {len(self.empresas)} empresas na planilha, "
                                f"{len(self.clientes)} pastas de clientes, {len(self.base.exemplos)} respostas "
                                f"aprendidas.{extra} {self.aviso_cadastro}")
            if not caminho:
                messagebox.showwarning("Planilha de clientes",
                                       "Não achei a planilha de clientes (“Clientes AAAA-MM-DD.xlsx”) em Downloads.\n\n"
                                       "Clique em “Trocar planilha…” e aponte onde ela está. A planilha precisa das "
                                       "colunas Código, Razão Social, E-mails e Regime municipal.")
            elif not raiz:
                messagebox.showinfo("Pasta dos clientes",
                                    "Clique em “Pasta dos clientes…” e escolha a pasta onde cada cliente tem a sua "
                                    "subpasta (“<código> - <nome>”). É lá que as notas serão salvas.")

        self._rodar("Abrindo o Outlook…", trabalho, depois)

    def _usar_planilha(self, caminho: Path | None, empresas: list[planilha.Empresa]):
        self.caminho_planilha = caminho
        self.empresas = empresas
        self._por_codigo = {e.codigo: e for e in empresas}
        if caminho:
            self.config_.set("planilha", str(caminho))
            recebem = sum(1 for e in empresas if e.recebe)
            self.var_planilha.set(f"Planilha: {caminho.name}  ·  {len(empresas)} empresas, {recebem} com tomados")
            mudou = self.cadastro.atualizar(empresas, self._emails_dos_grupos(), caminho.name)
            if mudou["novos"] or mudou["sumiram"]:
                self.aviso_cadastro = (f"Cadastro atualizado: {len(mudou['novos'])} e-mail(s) novo(s), "
                                       f"{len(mudou['sumiram'])} sumiram da planilha (continuam no cadastro).")
        else:
            self.var_planilha.set("Planilha de clientes: não encontrada")
        self._refazer_mapa()
        self._atualizar_envio()
        self._preencher_cadastro()

    def _empresas_do_contato(self, contato) -> list[planilha.Empresa]:
        return [self._por_codigo.get(cod) or planilha.Empresa(cod, contato.nome or "(fora da planilha)", "")
                for cod in contato.codigos]

    def _manuais(self) -> dict[str, list[planilha.Empresa]]:
        return {e: self._empresas_do_contato(self.cadastro[e]) for e in self.cadastro.manuais()}

    def _refazer_mapa(self):
        """E-mail -> empresas: planilha + o que foi acrescentado/trocado à mão; e-mail fora não identifica nada."""
        self.mapa = planilha.por_email(self.empresas)
        for email, empresas in self._manuais().items():
            self.mapa[email] = empresas
        self.ident = K.Identificador(self.clientes, self.registro.remetentes, self.nomes, self.mapa)

    def _trocar_planilha(self):
        inicial = self.caminho_planilha.parent if self.caminho_planilha else (self.demo.pasta if self.demo else DOWNLOADS)
        arquivo = filedialog.askopenfilename(title="Planilha de clientes", initialdir=str(inicial),
                                             filetypes=[("Excel", "*.xlsx")])
        if not arquivo:
            return
        try:
            empresas = planilha.ler(Path(arquivo))
        except Exception as erro:
            messagebox.showerror("Planilha de clientes", f"Não consegui ler a planilha:\n\n{erro}")
            return
        self._usar_planilha(Path(arquivo), empresas)
        if self.respostas:
            self._identificar_todos()
            self._preencher_respostas()

    # ---- pasta dos clientes
    def _raiz(self) -> str:
        return self.config_.get("raiz_clientes", "") or ""

    def _atualizar_raiz_na_tela(self):
        raiz = self._raiz()
        if not raiz:
            self.var_raiz.set("Pasta dos clientes: não definida — clique em “Pasta dos clientes…”")
        elif not Path(raiz).is_dir():
            self.var_raiz.set(f"Pasta dos clientes: {raiz} (não encontrada)")
        else:
            self.var_raiz.set(f"Pasta dos clientes: {raiz}  ·  {len(self.clientes)} pastas")

    def _escolher_raiz(self):
        atual = self._raiz()
        pasta = filedialog.askdirectory(title="Pasta onde cada cliente tem a sua subpasta (“<código> - <nome>”)",
                                        initialdir=atual if atual and Path(atual).is_dir() else str(Path.home()))
        if not pasta:
            return
        self.config_.set("raiz_clientes", pasta)
        self.clientes = K.indexar(pasta)
        self._refazer_mapa()
        self._atualizar_raiz_na_tela()
        if self.respostas:
            self._identificar_todos()
            self._preencher_respostas()
        self.var_status.set(f"{len(self.clientes)} pastas de clientes em {pasta}." if self.clientes else
                            f"Nenhuma subpasta em {pasta}. Cada cliente precisa de uma pasta “<código> - <nome>”.")

    def _comp(self) -> C.Competencia:
        return C.do_texto(self.var_comp.get())

    def _mudou_competencia(self):
        self.respostas, self.palpites, self.sem_resposta = [], {}, []
        self.sugeridos, self.atribuicoes = {}, {}
        for t in (self.tab_resp, self.tab_sem, self.tab_dev, self.tab_fila):
            t.limpar()
        self._atualizar_envio()
        self.var_status.set(f"Competência {self.var_comp.get()}. Na aba 2, clique em “Buscar respostas”.")

    # ================================================================ aba 1: enviar
    def _preencher_grupos(self):
        self.tab_grupos.limpar()
        for g in self.grupos:
            marcado = g.nome in self.marcados
            self.tab_grupos.inserir(("✔" if marcado else "", g.nome, len(g.membros)), ("bom",) if marcado else ())

    def _clicou_grupo(self, evento):
        linha = self.tab_grupos.arvore.identify_row(evento.y)
        if not linha:
            return
        self.marcados.symmetric_difference_update({self.tab_grupos.arvore.set(linha, "grupo")})
        self.config_.set("grupos", sorted(self.marcados))
        self._preencher_grupos()
        self._atualizar_envio()

    def _emails_dos_grupos(self) -> dict[str, list[str]]:
        emails: dict[str, list[str]] = {}
        for g in self.grupos:
            if g.nome in self.marcados:
                for _nome, email in g.membros:
                    emails.setdefault(email.lower(), []).append(g.nome)
        return emails

    def _atualizar_envio(self):
        if not self.modelo:
            return
        comp = self._comp()
        self.lista = planilha.lista_de_envio(self.empresas, self.cadastro.fora_da_lista(), TAMANHO_LOTE,
                                             self._emails_dos_grupos(), self.cadastro.forcados(),
                                             self._manuais(), self.cadastro.respondentes(), self.cadastro.principais(),
                                             dominio_interno=self.dominio)
        lista = self.lista
        recebem = [d for d in lista.destinatarios if d.incluido]
        origens = {o: sum(1 for d in recebem if d.origem == o) for o in ("planilha + grupo", "só planilha", "só grupo")}
        linhas = [f"De {self.conta} para ela mesma, clientes em Cco: {len(recebem)} e-mails em "
                  f"{len(lista.lotes)} mensagens de até {TAMANHO_LOTE}.",
                  f"Nos dois: {origens['planilha + grupo']}  ·  só na planilha: {origens['só planilha']}  ·  "
                  f"só nos grupos: {origens['só grupo']}  ·  não recebem: {len(lista.destinatarios) - len(recebem)} "
                  "(clique no ✔ da lista para mudar)."]
        poupados = sum(1 for d in lista.destinatarios if not d.incluido
                       and d.observacao.startswith(("a empresa responde", "principal da empresa")))
        if poupados:
            linhas.append(f"{poupados} e-mail(s) não recebem porque a empresa deles responde por outro e-mail "
                          "(ou tem contato principal). Empresa sem histórico recebe em todos.")
        if lista.sem_email:
            linhas.append(f"{len(lista.sem_email)} empresa(s) da planilha sem e-mail (ver “Conferência da planilha”).")
        self.var_resumo_envio.set("\n".join(linhas))
        ja = self.envios.get(comp, [])
        if ja:
            quando = min(e.enviado_em for e in ja)
            self.var_aviso_envio.set(f"Atenção: {C.assunto(comp)} já foi enviado em {quando:%d/%m/%Y %H:%M} "
                                     f"({len(ja)} mensagens, {sum(len(e.destinatarios) for e in ja)} endereços). "
                                     "Enviar de novo manda tudo em dobro para os clientes.")
            self.lb_aviso.grid()
        else:
            self.lb_aviso.grid_remove()
        self._preencher_destinatarios()

    def _preencher_destinatarios(self):
        if not self.lista:
            return
        regra = {"Recebem": lambda d: d.incluido, "Não recebem": lambda d: not d.incluido,
                 "Só nos grupos": lambda d: d.origem == "só grupo", "Só na planilha": lambda d: d.origem == "só planilha",
                 "Planilha + grupo": lambda d: d.origem == "planilha + grupo"}.get(self.var_filtro.get(), lambda d: True)
        self.tab_dest.limpar()
        for d in self.lista.destinatarios:
            if not regra(d):
                continue
            empresas = "; ".join(e.rotulo for e in d.empresas[:3]) + \
                (f" (+{len(d.empresas) - 3})" if len(d.empresas) > 3 else "")
            self.tab_dest.arvore.insert("", "end", iid=d.email, tags=() if d.incluido else ("fora",), values=(
                "✔" if d.incluido else "", d.email, d.origem, empresas or "—", d.observacao))
            self.tab_dest._linhas += 1

    def _clicou_destinatario(self, evento):
        linha = self.tab_dest.arvore.identify_row(evento.y)
        if not linha or self.tab_dest.arvore.identify_column(evento.x) != "#1" or not self.lista:
            return
        d = next((x for x in self.lista.destinatarios if x.email == linha), None)
        if d is not None:
            self.cadastro.marcar(d.email, not d.incluido, d.padrao)
            self._atualizar_envio()

    # ---- layouts da pergunta
    def _atualizar_layouts(self, escolher: str | None = None):
        nomes = list(self.layouts.pergunta)
        self.cb_layout.configure(values=nomes)
        if escolher in nomes:
            self.var_layout.set(escolher)
            self._carregar_layout()

    def _carregar_layout(self):
        layout = self.layouts.pergunta.get(self.var_layout.get())
        if not layout:
            return
        self.var_assunto.set(layout["assunto"])
        self.txt_mensagem.delete("1.0", "end")
        self.txt_mensagem.insert("1.0", layout["texto"])
        self.layouts.pergunta_atual = self.var_layout.get()
        self.layouts.gravar()

    def _texto_mensagem(self) -> str:
        return self.txt_mensagem.get("1.0", "end").rstrip()

    def _salvar_layout(self, novo: bool = False):
        nome = self.var_layout.get()
        if novo or not nome:
            nome = (simpledialog.askstring("Salvar layout", "Nome do novo layout:", parent=self) or "").strip()
            if not nome:
                return
            if nome in self.layouts.pergunta and not messagebox.askyesno("Salvar layout",
                                                                         f"Já existe “{nome}”. Substituir?"):
                return
        self.layouts.salvar_pergunta(nome, self.var_assunto.get().strip(), self._texto_mensagem())
        self._atualizar_layouts(nome)
        self.var_status.set(f"Layout “{nome}” salvo.")

    def _excluir_layout(self):
        nome = self.var_layout.get()
        if len(self.layouts.pergunta) <= 1:
            messagebox.showinfo("Layout", "É o único layout: altere e salve em vez de excluir.")
            return
        if nome and messagebox.askyesno("Layout", f"Excluir o layout “{nome}”?"):
            self.layouts.excluir_pergunta(nome)
            self._atualizar_layouts(self.layouts.pergunta_atual)

    def _mensagem_pronta(self, comp: C.Competencia) -> tuple[str, str]:
        """(assunto, html) com os campos preenchidos agora — a saudação depende da hora do envio."""
        agora = datetime.now()
        assunto = L.preencher(self.var_assunto.get().strip(), comp, agora)
        corpo = L.html(L.preencher(self._texto_mensagem(), comp, agora), self.modelo.assinatura)
        return assunto, corpo

    def _ver_como_fica(self):
        if not self.modelo:
            return
        assunto, corpo = self._mensagem_pronta(self._comp())
        for arquivo, cid in self.modelo.imagens:  # no navegador o cid vira o arquivo da imagem
            if cid:
                corpo = corpo.replace(f"cid:{cid}", Path(arquivo).as_uri())
        cabecalho = (f"<p style=\"font:bold 14px 'Segoe UI';color:{tema.ACENTO_D}\">Assunto: {html_lib.escape(assunto)}"
                     "</p><hr>")
        corpo = corpo.replace("<body lang=PT-BR>", "<body lang=PT-BR>" + cabecalho, 1)
        previa = Path(tempfile.gettempdir()) / "malote_previa.html"
        previa.write_text(corpo, encoding="utf-8")
        _abrir_arquivo(previa)

    def _enviar(self, de_verdade: bool, enderecos: list[str] | None = None, titulo: str = "Perguntar aos clientes"):
        """Sem `enderecos`: a lista do mês. Com `enderecos`: reenvio só para quem não recebeu."""
        if not self.modelo or not self.lista:
            return
        comp = self._comp()
        lotes = (self.lista.lotes if enderecos is None
                 else [enderecos[i:i + TAMANHO_LOTE] for i in range(0, len(enderecos), TAMANHO_LOTE)])
        if not lotes:
            messagebox.showinfo(titulo, "Ninguém para receber.")
            return
        assunto, corpo = self._mensagem_pronta(comp)
        if re.search(r"\{\w+\}", assunto + re.sub(r"<[^>]+>", "", corpo)):
            messagebox.showwarning(titulo, "Sobrou um campo entre chaves que a ferramenta não conhece. "
                                           "Use só {saudacao}, {mes}, {mes_nome} e {ano}.")
            return
        total = sum(len(l) for l in lotes)
        ja = self.envios.get(comp) if enderecos is None else None
        acao = "ENVIAR AGORA" if de_verdade else "Criar como rascunho"
        texto = f"{acao}:\n“{assunto}”\n\n{len(lotes)} mensagem(ns), {total} e-mail(s) em Cco, pela {self.conta}."
        if self.demo:
            texto += "\n\n(Demonstração: nada sai de verdade; o envio fica só no Outlook simulado.)"
        if C.do_assunto(assunto) is None:
            texto += ("\n\nAtenção: o assunto está fora do padrão “Serviços tomados - <Mês> de <ano>”. A ferramenta "
                      "usa esse padrão para reconhecer o envio e achar as respostas.")
        if ja:
            texto = "ESTE MÊS JÁ FOI ENVIADO. Os clientes vão receber em dobro.\n\n" + texto
        if not messagebox.askyesno(titulo, texto + "\n\nConfirma?", icon="warning" if (ja or de_verdade) else "question"):
            return
        imagens = self.modelo.imagens

        def trabalho(ol, avisar):
            feitos = []
            for i, lote in enumerate(lotes, start=1):
                avisar(f"Mensagem {i} de {len(lotes)} ({len(lote)} e-mails)…")
                feitos.append((len(lote), ol.montar(assunto, corpo, imagens, lote, de_verdade)))
            return feitos, (ol.envios() if de_verdade else None)

        def depois(res):
            feitos, envios = res
            if envios is not None:
                self.envios = envios
            if enderecos is not None and de_verdade:
                self.registro.marcar_reenviado(enderecos, C.texto(comp))
                self.marcar_reenvio = {}
                self._preencher_respostas()
            self._atualizar_envio()
            onde = "enviada(s)" if de_verdade else f"salva(s) em {feitos[0][1]}"
            self.var_status.set(f"{len(feitos)} mensagem(ns) {onde} ({sum(f[0] for f in feitos)} e-mails em Cco).")
            messagebox.showinfo(titulo, f"Pronto: {len(feitos)} mensagem(ns) {onde}.")

        self._rodar("Montando as mensagens…", trabalho, depois)

    def _conferencia(self):
        if not self.empresas or not self.ident or not self.lista:
            return
        lista = self.lista
        sem_email = [[e.codigo, e.razao, e.resp_fiscal] for e in lista.sem_email]
        internos = [[d.email, "; ".join(e.rotulo for e in d.empresas)] for d in lista.destinatarios
                    if d.observacao == "e-mail interno"]
        compart = sorted(([m, len(es), "; ".join(e.rotulo for e in es)] for m, es in self.mapa.items() if len(es) > 1),
                         key=lambda x: -x[1])
        sem_pasta = [[e.codigo, e.razao, e.resp_fiscal] for e in self.empresas
                     if e.recebe and not self.ident.pasta_da_empresa(e)]
        fora = [[d.email, d.origem, d.observacao] for d in lista.destinatarios if not d.incluido]
        so_grupo = [[d.email, ", ".join(d.grupos), d.observacao] for d in lista.destinatarios if d.origem == "só grupo"]
        so_planilha = [[d.email, "; ".join(e.rotulo for e in d.empresas)] for d in lista.destinatarios
                       if d.origem == "só planilha"]
        arquivo = self.pasta_controle / f"Conferência da planilha {datetime.now():%Y-%m-%d}.xlsx"
        try:
            controle.gerar_conferencia(arquivo, sem_email, internos, compart, sem_pasta, fora, so_grupo, so_planilha)
        except PermissionError:
            messagebox.showerror("Conferência", f"{arquivo.name} está aberto no Excel. Feche e clique de novo.")
            return
        self.var_status.set(f"Conferência gravada: {arquivo}")
        _abrir_arquivo(arquivo)

    # ================================================================ aba 2: respostas
    def _buscar(self):
        if not self.ident:
            return
        comp = self._comp()
        envios = self.envios
        perguntados = {e.lower() for envio in envios.get(comp, []) for _n, e in envio.destinatarios}
        ident, mapa = self.ident, self.mapa
        conhecidos = set(self.cadastro.contatos) | set(mapa) | {e.lower() for e in self.registro.remetentes}

        def trabalho(ol, avisar):
            avisar(f"Lendo o que chegou na {self.conta} para {C.texto(comp)} (Caixa de entrada, subpastas e "
                   "lixo eletrônico)…")
            respostas = ol.respostas(comp, envios, perguntados, conhecidos)
            palpites, sugeridos = {}, {}
            for r in respostas:
                if r.tipo in ("Não entregue", "Interno"):
                    palpites[r.entry_id] = K.Palpite(None, "", 0)
                    continue
                p = ident.identificar(r.email, r.nome)
                palpites[r.entry_id] = p
                if len(p.candidatos) > 1 and r.anexos:
                    avisar(f"Lendo os anexos de {r.email} para achar o CNPJ…")
                    cnpjs: dict[str, list[str]] = {}
                    for e in mapa.get(r.email, []):
                        pasta = ident.pasta_da_empresa(e)
                        if pasta:
                            cnpjs.setdefault(pasta, []).append(e.cnpj)
                    achados = ol.cnpjs_dos_anexos(r.entry_id, [a.indice for a in r.anexos])
                    sugeridos[r.entry_id] = {i: anexos.qual_empresa(c, cnpjs) for i, c in achados.items()}
            return respostas, palpites, sugeridos, perguntados

        self._rodar("Buscando respostas…", trabalho, self._buscou)

    def _buscou(self, res):
        self.respostas, self.palpites, self.sugeridos, perguntados = res
        respondeu = {}
        for r in self.respostas:
            if r.tipo not in ("Não entregue", "Interno") and r.email and r.relacao == "resposta":
                respondeu[r.email] = max(respondeu.get(r.email, r.recebido), r.recebido)
        if self.cadastro.anotar_respostas(respondeu):
            self._atualizar_envio()
            self._preencher_cadastro()
        self.atribuicoes = {}
        self._calcular_sem_resposta(perguntados)
        self._preencher_respostas()
        self._preencher_responder()
        desta = [r for r in self.respostas if r.da_competencia and r.tipo != "Não entregue"]
        com = sum(1 for r in desta if r.tipo == "Mandou anexo")
        outros = sum(1 for r in desta if r.relacao == "outro")
        varias = sum(1 for p in self.palpites.values() if p.como == "várias empresas")
        outras = sum(1 for r in self.respostas if not r.da_competencia)
        self.var_status.set(f"{len(desta)} e-mails de {C.texto(self._comp())}: {com} com anexo, {varias} de e-mail "
                            f"usado por várias empresas, {outros} de outros assuntos; {outras} de outra competência "
                            f"(em cinza). {len(self.sem_resposta)} empresas sem resposta.")

    def _calcular_sem_resposta(self, perguntados: set[str]):
        respondeu = {r.email for r in self.respostas if r.tipo not in ("Não entregue", "Interno")
                     and r.da_competencia and (r.relacao == "resposta" or r.anexos)}
        pastas = set()
        for r in self.respostas:
            pastas |= set(self._destinos(r).values())
        self.sem_resposta = []
        for e in self.empresas:
            if not e.recebe:
                continue
            perguntado = [m for m in e.emails if m in perguntados]
            if not perguntado or any(m in respondeu for m in e.emails):
                continue
            if self.ident and self.ident.pasta_da_empresa(e) in pastas:
                continue
            self.sem_resposta.append((e, perguntado))

    def _identificar_todos(self):
        self.ident = K.Identificador(self.clientes, self.registro.remetentes, self.nomes, self.mapa)
        for r in self.respostas:
            if r.tipo not in ("Não entregue", "Interno"):
                self.palpites[r.entry_id] = self.ident.identificar(r.email, r.nome)

    def _destinos(self, r: outlook.Resposta) -> dict[int, str]:
        """Anexo -> pasta, só o que já pode ser salvo (planilha, confirmado ou escolhido na tela)."""
        if r.entry_id in self.atribuicoes:
            return self.atribuicoes[r.entry_id]
        p = self.palpites.get(r.entry_id)
        if p and p.certo:
            return {a.indice: p.pasta for a in r.anexos}
        return {}

    def _descricao_cliente(self, r: outlook.Resposta) -> tuple[str, str]:
        p = self.palpites.get(r.entry_id) or K.Palpite(None, "")
        if r.tipo == "Interno":
            return "(interno)", ""
        escolhidos = self.atribuicoes.get(r.entry_id)
        if escolhidos:
            pastas = sorted(set(escolhidos.values()))
            return " + ".join(pastas), "escolhido"
        if p.como == "várias empresas":
            sug = self.sugeridos.get(r.entry_id, {})
            achou = sorted({v for v in sug.values() if v})
            if achou:
                return " + ".join(achou) + f"  (de {len(p.candidatos)})", "CNPJ no anexo"
            return f"várias empresas ({len(p.candidatos)}) — escolher", "várias empresas"
        if p.pasta:
            return p.pasta, ("planilha" if p.como == "planilha" else
                             "confirmado" if p.como == "confirmado" else f"sugestão: {p.como}")
        return "?", "—"

    def _preencher_respostas(self):
        sel = set(self.tab_resp.arvore.selection())
        self.tab_resp.limpar()
        regra = {"Respostas à pergunta": lambda r: r.relacao == "resposta",
                 "Contatos do cadastro": lambda r: r.relacao == "contato",
                 "Outros assuntos": lambda r: r.relacao == "outro",
                 "Com anexo": lambda r: bool(r.anexos),
                 "Desta competência": lambda r: r.da_competencia,
                 "Outra competência": lambda r: not r.da_competencia,
                 "Lixo eletrônico": lambda r: r.pasta_outlook.lower().startswith("lixo"),
                 }.get(self.var_mostrar.get(), lambda r: True)
        total = 0
        for r in self.respostas:
            if r.tipo == "Não entregue":
                continue
            total += 1
            if not regra(r):
                continue
            cliente, como = self._descricao_cliente(r)
            salvo = self.registro.salvos.get(r.entry_id)
            respondido = r.respondido or r.entry_id in self.registro.respondidos
            pronto = bool(self._destinos(r)) or (self.palpites.get(r.entry_id) and self.palpites[r.entry_id].certo)
            if salvo:
                tags = ("bom",)
            elif r.tipo == "Mandou anexo" and cliente == "?":
                tags = ("alerta",)
            elif not pronto and r.tipo != "Interno" and cliente != "?":
                tags = ("ambar",)
            else:
                tags = ()
            if not r.da_competencia or r.relacao == "outro":
                tags = tags + ("outra",)
            porque = outlook.RELACOES.get(r.relacao, r.relacao)
            if r.pasta_outlook.lower().startswith("lixo"):
                porque += " · lixo eletrônico"
            elif "/" in r.pasta_outlook:
                porque += " · subpasta"
            self.tab_resp.arvore.insert("", "end", iid=r.entry_id, tags=tags, values=(
                f"{r.recebido:%d/%m %H:%M}", r.nome, cliente, como, porque, r.tipo, len(r.anexos) or "",
                "" if r.da_competencia or not r.competencia else C.texto(r.competencia),
                "salvo" if salvo else "", "sim" if respondido else ""))
            self.tab_resp._linhas += 1
        for iid in sel & set(self.tab_resp.arvore.get_children()):
            self.tab_resp.arvore.selection_add(iid)
        mostradas = len(self.tab_resp.arvore.get_children())
        self.sub.tab(0, text=f"Respostas ({mostradas} de {total})" if mostradas != total else f"Respostas ({total})")

        self.tab_sem.limpar()
        for e, emails in self.sem_resposta:
            self.tab_sem.inserir((e.rotulo, e.resp_fiscal, ", ".join(emails)))
        self.sub.tab(1, text=f"Sem resposta ({len(self.sem_resposta)})")

        self.tab_dev.limpar()
        for email, (recebido, motivo, reenviar) in self._devolvidos().items():
            p = self.ident.identificar(email, self.nomes.get(email, ""))
            cliente = p.pasta or (f"várias ({len(p.candidatos)})" if p.candidatos else "?")
            if self.cadastro[email] and self.cadastro[email].fora:
                situacao = "fora da lista"
            elif email in self.registro.reenviados:
                situacao = "reenviado " + self.registro.reenviados[email].split(" em ")[-1]
            else:
                situacao = ""
            marcado = self.marcar_reenvio.get(email, reenviar and not situacao)
            self.tab_dev.arvore.insert("", "end", iid=email, tags=() if reenviar else ("alerta",), values=(
                "✔" if marcado else "", f"{recebido:%d/%m %H:%M}", email, motivo, cliente, situacao))
            self.tab_dev._linhas += 1
        self.sub.tab(2, text=f"Não entregues ({len(self.tab_dev.arvore.get_children())})")

    def _devolvidos(self) -> dict[str, tuple[datetime, str, bool]]:
        """{e-mail: (quando voltou, motivo, vale reenviar)} — a devolução mais recente de cada endereço."""
        achados: dict[str, tuple[datetime, str, bool]] = {}
        for r in sorted((r for r in self.respostas if r.tipo == "Não entregue"), key=lambda r: r.recebido):
            for email in r.nao_entregues:
                motivo, reenviar = r.motivos.get(email, ("não entregue", True))
                achados[email] = (r.recebido, motivo, reenviar)
        return achados

    def _resposta(self, entry_id: str) -> outlook.Resposta | None:
        return next((r for r in self.respostas if r.entry_id == entry_id), None)

    def _mostrar_detalhe(self):
        sel = self.tab_resp.arvore.selection()
        r = self._resposta(sel[0]) if sel else None
        texto = ""
        if r:
            texto = f"{r.nome} <{r.email}>  ·  {r.assunto}\n"
            texto += f"Pasta no Outlook: {r.pasta_outlook}  ·  {outlook.RELACOES.get(r.relacao, r.relacao)}"
            if not r.da_competencia and r.competencia:
                texto += f"  ·  atribuído à competência {C.texto(r.competencia)}"
            texto += "\n"
            empresas = self.mapa.get(r.email, [])
            if empresas:
                texto += "Na planilha: " + "; ".join(e.rotulo for e in empresas[:6]) + \
                         (f" … (+{len(empresas) - 6})" if len(empresas) > 6 else "") + "\n"
            destinos = self._destinos(r)
            sug = self.sugeridos.get(r.entry_id, {})
            for a in r.anexos:
                onde = destinos.get(a.indice) or (f"CNPJ de {sug[a.indice]}" if sug.get(a.indice) else "")
                texto += f"   📎 {a.nome}" + (f"  →  {onde}" if onde else "") + "\n"
            if r.embutidas:
                texto += f"   ({r.embutidas} imagem(ns) de assinatura ficam de fora)\n"
            salvo = self.registro.salvos.get(r.entry_id)
            if salvo:
                texto += f"Salvo em: {salvo['destino']}\n"
            elif destinos and self._raiz():
                pasta = next(iter(destinos.values()))
                try:
                    texto += "Vai para: " + K.destino(self._raiz(), pasta, self._comp()).descricao + "\n"
                except OSError as erro:
                    texto += f"Pasta: {erro}\n"
            if r.como_competencia != "assunto":
                origem = "do nome do anexo" if r.como_competencia == "anexo" else "da data"
                texto += f"Competência {C.texto(r.competencia)} tirada {origem}.\n"
            texto += "\n" + r.texto
        self.txt_detalhe.configure(state="normal")
        self.txt_detalhe.delete("1.0", "end")
        self.txt_detalhe.insert("1.0", texto)
        self.txt_detalhe.configure(state="disabled")

    def _escolher(self):
        sel = self.tab_resp.arvore.selection()
        if not sel:
            messagebox.showinfo("Cliente", "Selecione uma resposta na lista.")
            return
        r = self._resposta(sel[0])
        if r is None or r.tipo in ("Interno", "Não entregue"):
            return
        p = self.palpites.get(r.entry_id) or K.Palpite(None, "")
        if len(p.candidatos) > 1:
            if not r.anexos:
                messagebox.showinfo("Cliente", "Este e-mail serve a várias empresas e não trouxe anexo: "
                                               "não há o que salvar. Ele conta como resposta de todas elas.")
                return
            atual = self.atribuicoes.get(r.entry_id) or {i: v for i, v in self.sugeridos.get(r.entry_id, {}).items() if v}
            escolha = DistribuirAnexos(self, r, p.candidatos, atual, self.clientes).resultado
            if escolha is not None:
                self.atribuicoes[r.entry_id] = escolha
        else:
            escolhido = EscolherCliente(self, r, p, self.clientes).resultado
            if not escolhido:
                return
            self.registro.aprender(r.email, escolhido)
            self._identificar_todos()
            self.var_status.set(f"{r.email} → {escolhido}. Fica guardado para os próximos meses.")
        self._preencher_respostas()
        self._mostrar_detalhe()

    def _confirmar_sugestoes(self):
        simples, varias = [], []
        for r in self.respostas:
            p = self.palpites.get(r.entry_id)
            if not p or r.tipo in ("Interno", "Não entregue") or r.entry_id in self.atribuicoes:
                continue
            if p.pasta and not p.certo:
                simples.append((r, p))
            sug = self.sugeridos.get(r.entry_id, {})
            if p.como == "várias empresas" and r.anexos and all(sug.get(a.indice) for a in r.anexos):
                varias.append((r, sug))
        if not simples and not varias:
            messagebox.showinfo("Cliente", "Não há sugestão para confirmar.")
            return
        linhas = [f"• {r.email} → {p.pasta}" for r, p in simples[:10]]
        linhas += [f"• {r.email} → {' + '.join(sorted(set(s.values())))} (CNPJ no anexo)" for r, s in varias[:6]]
        if not messagebox.askyesno(
                "Confirmar sugestões",
                f"Confirmar {len(simples) + len(varias)} sugestões (linhas em amarelo)?\n\n" + "\n".join(linhas) +
                "\n\nConfira antes. Remetente fora da planilha confirmado fica guardado para os próximos meses. "
                "A que estiver errada, troque com dois cliques."):
            return
        for r, p in simples:
            self.registro.remetentes[r.email.lower()] = p.pasta
        self.registro.gravar()
        for r, sug in varias:
            self.atribuicoes[r.entry_id] = {i: v for i, v in sug.items() if v}
        self._identificar_todos()
        self._preencher_respostas()
        self.var_status.set(f"{len(simples) + len(varias)} sugestões confirmadas.")

    def _salvar(self):
        comp = self._comp()
        raiz = self._raiz()
        if not raiz or not Path(raiz).is_dir():
            messagebox.showinfo("Salvar anexos", "Escolha antes a pasta dos clientes (botão “Pasta dos clientes…”).")
            return
        fila, faltam = [], 0
        for r in self.respostas:
            if r.tipo != "Mandou anexo" or r.entry_id in self.registro.salvos:
                continue
            destinos = self._destinos(r)
            if not destinos or len(destinos) < len(r.anexos):
                faltam += 1
                continue
            por_pasta: dict[str, list[int]] = {}
            for indice, pasta in destinos.items():
                por_pasta.setdefault(pasta, []).append(indice)
            fila.append((r.entry_id, por_pasta, r.competencia or comp))
        if not fila:
            messagebox.showinfo("Salvar anexos", "Nada pronto para salvar." +
                                (f"\n\n{faltam} resposta(s) com anexo esperam você escolher ou confirmar o cliente "
                                 "(linhas em amarelo ou vermelho)." if faltam else ""))
            return
        pastas = {p for _e, pp, _c in fila for p in pp}
        de_outra = sum(1 for _e, _pp, c in fila if c != comp)
        try:
            exemplo = K.destino(raiz, next(iter(pastas)), comp).descricao
        except OSError as erro:
            messagebox.showerror("Salvar anexos", str(erro))
            return
        aviso = f"\n\nFicam esperando: {faltam} resposta(s) sem cliente definido." if faltam else ""
        if not messagebox.askyesno("Salvar anexos",
                                   f"Salvar os anexos de {len(fila)} e-mail(s) em {len(pastas)} cliente(s), em "
                                   f"NFE\\{comp[0]}\\{comp[1]:02d}-{comp[0]}\\Tomados\\{K.PASTA_FINAL}?\n\n"
                                   f"Exemplo:\n{exemplo}\n\nPastas que faltarem são criadas. Nada é sobrescrito."
                                   + (f"\n\n{de_outra} e-mail(s) citam outra competência e vão para a pasta dela."
                                      if de_outra else "") + aviso):
            return

        def trabalho(ol, avisar):
            feitos = []
            for i, (entry_id, por_pasta, comp_r) in enumerate(fila, start=1):
                for pasta, indices in por_pasta.items():
                    avisar(f"[{i}/{len(fila)}] {pasta}…")
                    try:
                        alvo = K.garantir(K.destino(raiz, pasta, comp_r))
                        feitos.append((entry_id, pasta, str(alvo), ol.salvar_anexos(entry_id, indices, alvo), None,
                                       C.texto(comp_r)))
                    except Exception as erro:
                        feitos.append((entry_id, pasta, "", [], str(erro), C.texto(comp_r)))
            return feitos

        def depois(feitos):
            erros, salvos, por_email, comps = [], 0, {}, {}
            for entry_id, pasta, alvo, arquivos, erro, comp_txt in feitos:
                comps[entry_id] = comp_txt
                if erro:
                    erros.append(f"{pasta}: {erro}")
                    continue
                salvos += sum(1 for _n, s in arquivos if s == "salvo")
                junto = por_email.setdefault(entry_id, ([], [], []))
                junto[0].append(pasta)
                junto[1].append(alvo)
                junto[2].extend(n for n, _s in arquivos)
            com_erro = {f[0] for f in feitos if f[4]}
            for entry_id, (pastas_, alvos, nomes) in por_email.items():
                if entry_id not in com_erro:  # e-mail com erro em alguma pasta fica para tentar de novo
                    self.registro.marcar_salvo(entry_id, " + ".join(pastas_), comps[entry_id], " ; ".join(alvos), nomes)
            self._preencher_respostas()
            msg = f"{salvos} arquivo(s) salvos de {len(por_email)} e-mail(s)."
            self.var_status.set(msg)
            if erros:
                messagebox.showwarning("Salvar anexos", msg + "\n\nNão deu:\n" + "\n".join(erros[:15]))
            else:
                messagebox.showinfo("Salvar anexos", msg)

        self._rodar("Salvando anexos…", trabalho, depois)

    def _planilha_controle(self):
        if not self.respostas and not self.sem_resposta:
            messagebox.showinfo("Planilha", "Busque as respostas primeiro.")
            return
        comp = self._comp()
        respostas, devolvidos = [], []
        for r in self.respostas:
            if r.tipo == "Não entregue":
                for email in r.nao_entregues:
                    q = self.ident.identificar(email, self.nomes.get(email, ""))
                    devolvidos.append([r.recebido.strftime("%d/%m/%Y %H:%M"), email, q.pasta or ""])
                continue
            cliente, _como = self._descricao_cliente(r)
            salvo = self.registro.salvos.get(r.entry_id)
            respondido = r.respondido or r.entry_id in self.registro.respondidos
            respostas.append([r.recebido.strftime("%d/%m/%Y %H:%M"), cliente, r.nome, r.email, r.tipo,
                              len(r.anexos), salvo["destino"] if salvo else "", "sim" if respondido else "",
                              re.sub(r"\s+", " ", r.texto)[:500]])
        sem = [[e.rotulo, e.resp_fiscal, ", ".join(emails)] for e, emails in self.sem_resposta]
        arquivo = self.pasta_controle / f"Tomados {comp[1]:02d}-{comp[0]}.xlsx"
        try:
            controle.gerar(arquivo, respostas, sem, devolvidos)
        except PermissionError:
            messagebox.showerror("Planilha", f"{arquivo.name} está aberto no Excel. Feche e clique de novo.")
            return
        self.var_status.set(f"Planilha gravada: {arquivo}")
        _abrir_arquivo(arquivo)

    def _clicou_devolvido(self, evento):
        linha = self.tab_dev.arvore.identify_row(evento.y)
        if not linha or self.tab_dev.arvore.identify_column(evento.x) != "#1":
            return
        self.marcar_reenvio[linha] = not bool(self.tab_dev.arvore.set(linha, "marca"))
        self._preencher_respostas()

    def _reenviar(self):
        marcados = [i for i in self.tab_dev.arvore.get_children() if self.tab_dev.arvore.set(i, "marca")]
        if not marcados:
            messagebox.showinfo("Reenviar", "Marque (✔) os endereços que devem receber de novo.")
            return
        fora = [e for e in marcados if self.cadastro[e] and self.cadastro[e].fora]
        if fora:
            messagebox.showinfo("Reenviar", "Estes estão fora da lista; volte-os para a lista antes:\n" + "\n".join(fora))
            return
        self._enviar(True, enderecos=marcados, titulo="Reenviar para quem não recebeu")

    def _devolvidos_selecionados(self) -> list[str]:
        sel = list(self.tab_dev.arvore.selection())
        if not sel:
            messagebox.showinfo("Lista", "Selecione um ou mais endereços na lista.")
        return sel

    def _tirar_devolvido(self):
        devolvidos = self._devolvidos()
        for email in self._devolvidos_selecionados():
            self.cadastro.tirar(email, f"não entregue: {devolvidos.get(email, (0, 'voltou'))[1]}")
        self._preencher_respostas()
        self._atualizar_envio()

    def _voltar_devolvido(self):
        for email in self._devolvidos_selecionados():
            self.cadastro.voltar(email)
        self._preencher_respostas()
        self._atualizar_envio()

    # ================================================================ aba 3: responder
    def _a_responder(self) -> list[outlook.Resposta]:
        todas = self.var_mostrar_respondidos.get()
        return [r for r in self.respostas if r.tipo not in ("Não entregue", "Interno") and r.relevante
                and (todas or not (r.respondido or r.entry_id in self.registro.respondidos))]

    def _preencher_responder(self):
        self.tab_fila.limpar()
        fila = self._a_responder()
        for r in fila:
            cliente, _como = self._descricao_cliente(r)
            quem = r.nome if cliente in ("?", "") or cliente.startswith("várias") else cliente
            cat = sugestao.categoria(r.texto, r.tipo)
            feito = r.respondido or r.entry_id in self.registro.respondidos
            self.tab_fila.arvore.insert("", "end", iid=r.entry_id, values=(f"{r.recebido:%d/%m %H:%M}", quem, cat),
                                        tags=("bom",) if feito else (("ambar",) if cat == "Pergunta" else ()))
            self.tab_fila._linhas += 1
        self.abas.tab(2, text=f"3 · Responder ({len(fila)})")

    def _abrir_resposta(self):
        sel = self.tab_fila.arvore.selection()
        r = self._resposta(sel[0]) if sel else None
        self.atual = r
        if r is None:
            return
        s = self.base.sugerir(r.texto, r.tipo)
        self.sugestao_atual = s
        cliente, _como = self._descricao_cliente(r)
        self.var_cabecalho.set(f"{r.nome}  <{r.email}>" + (f"   ·   {cliente}" if cliente != "?" else ""))
        self.txt_cliente.configure(state="normal")
        self.txt_cliente.delete("1.0", "end")
        extra = ("\n\n[anexos: " + "; ".join(a.nome for a in r.anexos) + "]") if r.anexos else ""
        self.txt_cliente.insert("1.0", (r.texto or "(sem texto)") + extra)
        self.txt_cliente.configure(state="disabled")
        self.var_saudacao.set(sugestao.saudacao(r.nome))
        self.txt_resposta.delete("1.0", "end")
        self.txt_resposta.insert("1.0", s.resposta)
        feito = self.registro.respondidos.get(r.entry_id)
        porque = f"{s.categoria} · {s.porque}"
        if r.respondido or feito:
            quando = f" pela ferramenta em {feito['quando'][:16].replace('T', ' ')}" if feito else " no Outlook"
            porque = f"JÁ RESPONDIDA{quando}.   " + porque
        self.var_porque.set(porque)
        self.var_aprender.set(True)
        self._propor_mudanca(r, s.categoria)
        self.txt_resposta.focus_set()

    def _propor_mudanca(self, r: outlook.Resposta, categoria: str):
        """Pedido de sair da lista / trocar e-mail vira mudança pronta no cadastro (aplicada ao aceitar)."""
        self.mudanca = None
        if categoria == "Trocar e-mail":
            troca = cadastro_mod.pedido_de_troca(r.email, r.texto, self.cadastro, self.dominio)
            if troca:
                antigo, novo = troca
                empresas = ", ".join(self.cadastro[antigo].codigos) if self.cadastro[antigo] else ""
                self.mudanca = ("trocar", antigo, novo)
                self.var_texto_mudanca.set(f"Cadastro: trocar {antigo} por {novo}"
                                           + (f" (empresas {empresas})" if empresas else ""))
            else:
                self.var_texto_mudanca.set("Pede troca de e-mail, mas não achei o endereço novo no texto — "
                                           "faça na aba 4 · Cadastro.")
        elif categoria == "Sair da lista":
            self.mudanca = ("tirar", r.email)
            self.var_texto_mudanca.set(f"Cadastro: tirar {r.email} da lista de envio (pediu para sair)")
        else:
            self.mudanca = ("tirar", r.email)
            self.var_texto_mudanca.set(f"Tirar {r.email} da lista de envio")
        self.var_tirar.set(categoria in ("Trocar e-mail", "Sair da lista") and self.mudanca is not None)
        self.ck_mudanca.state(["!disabled"] if self.mudanca else ["disabled"])

    def _aplicar_mudanca(self, r: outlook.Resposta, mudanca: tuple | None):
        if not mudanca:
            return
        por = f"pedido de {r.email} em {r.recebido:%d/%m/%Y}"
        if mudanca[0] == "trocar":
            self.cadastro.trocar(mudanca[1], mudanca[2], por)
        else:
            self.cadastro.tirar(mudanca[1], "pediu para sair", por)
        self._refazer_mapa()
        self._atualizar_envio()
        self._preencher_cadastro()

    # ---- respostas prontas (layouts de resposta)
    def _atualizar_prontas(self):
        self.cb_pronta.configure(values=sorted(self.layouts.resposta))

    def _usar_pronta(self):
        texto = self.layouts.resposta.get(self.var_pronta.get())
        if texto is None:
            return
        nome = sugestao.primeiro_nome(self.atual.nome) if self.atual else ""
        self.txt_resposta.delete("1.0", "end")
        self.txt_resposta.insert("1.0", L.preencher(texto, self._comp(), nome=nome))
        self.txt_resposta.focus_set()

    def _salvar_pronta(self):
        texto = self.txt_resposta.get("1.0", "end").strip()
        if not texto:
            messagebox.showinfo("Resposta pronta", "Escreva a resposta antes de guardar.")
            return
        nome = (simpledialog.askstring("Resposta pronta", "Nome desta resposta pronta:", parent=self,
                                       initialvalue=self.var_pronta.get()) or "").strip()
        if not nome:
            return
        if nome in self.layouts.resposta and not messagebox.askyesno("Resposta pronta", f"Substituir “{nome}”?"):
            return
        self.layouts.salvar_resposta(nome, texto)
        self._atualizar_prontas()
        self.var_pronta.set(nome)
        self.var_status.set(f"Resposta pronta “{nome}” guardada. Dá para usar {{mes}} e {{nome}} no texto.")

    def _excluir_pronta(self):
        nome = self.var_pronta.get()
        if nome in self.layouts.resposta and messagebox.askyesno("Resposta pronta", f"Excluir “{nome}”?"):
            self.layouts.excluir_resposta(nome)
            self.var_pronta.set("")
            self._atualizar_prontas()

    def _proxima(self):
        antes = list(self.tab_fila.arvore.get_children())
        atual = self.atual.entry_id if self.atual else None
        self._preencher_responder()
        agora = list(self.tab_fila.arvore.get_children())
        if not agora:
            self.atual = None
            return
        seguinte = agora[0]
        if atual in antes:
            seguinte = next((f for f in antes[antes.index(atual) + 1:] if f in agora), agora[0])
        self.tab_fila.arvore.selection_set(seguinte)
        self.tab_fila.arvore.see(seguinte)

    def _aceitar(self, enviar: bool):
        r = self.atual
        if r is None or self.modelo is None:
            return
        miolo = self.txt_resposta.get("1.0", "end").strip()
        if not miolo:
            messagebox.showinfo("Responder", "Escreva a resposta antes.")
            return
        texto = sugestao.montar_texto(self.var_saudacao.get(), miolo)
        aprender, tirar = self.var_aprender.get(), self.var_tirar.get()
        mudanca = self.mudanca  # guardado agora: durante o envio a tela pode mudar de resposta
        sug = self.sugestao_atual
        cat = sug.categoria if sug else "Outro"
        if enviar and (r.respondido or r.entry_id in self.registro.respondidos):
            if not messagebox.askyesno("Responder", "Esta mensagem já foi respondida. Responder de novo?"):
                return

        def registrar():
            if aprender:
                self.base.ensinar(r.texto, miolo, cat, "aceito" if sug and miolo == sug.resposta else "ensinado")
            if tirar:
                self._aplicar_mudanca(r, mudanca)

        if not enviar:
            registrar()
            self.var_status.set("Resposta guardada como exemplo (nada foi enviado).")
            self._proxima()
            return
        modelo, entry_id = self.modelo, r.entry_id

        def trabalho(ol, avisar):
            avisar(f"Respondendo {r.email}…")
            ol.responder(entry_id, texto, modelo)

        def depois(_res):
            self.registro.marcar_respondido(entry_id, r.email, texto)
            registrar()
            self._preencher_respostas()
            self.var_status.set(f"Resposta enviada para {r.email}.")
            self._proxima()

        self._rodar("Enviando a resposta…", trabalho, depois)

    # ================================================================ aba 4: cadastro
    def _preencher_cadastro(self):
        if not hasattr(self, "tab_cad"):
            return
        termo = K.normalizar(self.var_busca_cad.get())
        filtro = self.var_filtro_cad.get()
        regra = {"Na lista": lambda c: not c.fora, "Fora da lista": lambda c: c.fora,
                 "Responderam": lambda c: bool(c.respondeu_em), "Principais": lambda c: c.principal,
                 "Acrescentados à mão": lambda c: c.manual, "Sumiram da planilha": lambda c: bool(c.sumiu_em),
                 "Mudados por pedido": lambda c: any(h["por"].startswith("pedido de") for h in c.historico)
                 }.get(filtro, lambda c: True)
        self.tab_cad.limpar()
        total = 0
        for email, c in sorted(self.cadastro.contatos.items()):
            total += 1
            empresas = "; ".join(e.rotulo for e in self._empresas_do_contato(c))
            if not regra(c) or (termo and termo not in K.normalizar(f"{email} {empresas} {c.nome}")):
                continue
            situacao = f"fora: {c.motivo}" if c.fora else ("na lista" + (" (marcado à mão)" if c.forcar_envio else ""))
            if c.principal:
                situacao = "★ principal · " + situacao
            if c.sumiu_em and not c.fora:
                situacao += f" · sumiu da planilha {c.sumiu_em}"
            origem = ", ".join("grupo" if o.startswith("grupo:") else o for o in dict.fromkeys(c.origens)) or "—"
            ultima = f"{c.historico[-1]['quando']} {c.historico[-1]['o_que']}" if c.historico else ""
            tags = ("fora",) if c.fora else (("ambar",) if c.sumiu_em else ())
            self.tab_cad.arvore.insert("", "end", iid=email, tags=tags,
                                       values=(email, empresas or "—", situacao,
                                               "/".join(reversed(c.respondeu_em.split("-"))) if c.respondeu_em else "",
                                               origem, ultima))
            self.tab_cad._linhas += 1
        fora = sum(1 for c in self.cadastro.contatos.values() if c.fora)
        self.var_info_cad.set(f"{total} e-mails no cadastro, {fora} fora da lista. Mostrando "
                              f"{len(self.tab_cad.arvore.get_children())}. Atualizado com "
                              f"{self.cadastro.ultima_atualizacao or '—'}. O que você muda aqui vale mais que a "
                              "planilha.")

    def _cad_selecionados(self) -> list[str]:
        sel = list(self.tab_cad.arvore.selection())
        if not sel:
            messagebox.showinfo("Cadastro", "Selecione um ou mais e-mails na lista.")
        return sel

    def _depois_de_mudar_cadastro(self, aviso: str):
        self._refazer_mapa()
        self._atualizar_envio()
        self._preencher_cadastro()
        if self.respostas:
            self._identificar_todos()
            self._preencher_respostas()
        self.var_status.set(aviso)

    def _mostrar_historico(self):
        sel = self.tab_cad.arvore.selection()
        c = self.cadastro[sel[0]] if sel else None
        texto = ""
        if c:
            texto = f"{c.email}" + (f"  ·  {c.nome}" if c.nome else "") + "\n"
            texto += "\n".join(f"{h['quando']}  {h['o_que']}  ({h['por']})" for h in reversed(c.historico))
        self.txt_hist.configure(state="normal")
        self.txt_hist.delete("1.0", "end")
        self.txt_hist.insert("1.0", texto)
        self.txt_hist.configure(state="disabled")

    def _mudar_cad(self, voltar: bool):
        emails = self._cad_selecionados()
        if not emails:
            return
        if voltar:
            for e in emails:
                self.cadastro.voltar(e)
        else:
            motivo = simpledialog.askstring("Tirar da lista", "Motivo (fica no histórico):", parent=self,
                                            initialvalue="pedido do cliente")
            if motivo is None:
                return
            for e in emails:
                self.cadastro.tirar(e, motivo.strip() or "tirado à mão")
        self._depois_de_mudar_cadastro(f"{len(emails)} e-mail(s) {'voltaram para a' if voltar else 'saíram da'} lista.")

    def _principal(self):
        emails = self._cad_selecionados()
        if not emails:
            return
        fixar = not all(self.cadastro[e] and self.cadastro[e].principal for e in emails)
        for e in emails:
            self.cadastro.fixar_principal(e, fixar)
        self._depois_de_mudar_cadastro(
            f"{len(emails)} e-mail(s) {'fixados como contato principal' if fixar else 'deixaram de ser principais'}. "
            "Empresa com principal recebe só nele.")

    def _trocar_email(self):
        emails = self._cad_selecionados()
        if len(emails) != 1:
            if emails:
                messagebox.showinfo("Trocar e-mail", "Selecione um e-mail só.")
            return
        novo = simpledialog.askstring("Trocar e-mail", f"Novo e-mail no lugar de {emails[0]}:", parent=self)
        if not novo:
            return
        try:
            c = self.cadastro.trocar(emails[0], novo)
        except ValueError as erro:
            messagebox.showerror("Trocar e-mail", str(erro))
            return
        self._depois_de_mudar_cadastro(f"{emails[0]} trocado por {c.email} (empresas {', '.join(c.codigos) or '—'}).")

    def _acrescentar(self):
        dados = AcrescentarContato(self, self.empresas).resultado
        if not dados:
            return
        email, nome, codigos = dados
        try:
            self.cadastro.acrescentar(email, codigos, nome)
        except ValueError as erro:
            messagebox.showerror("Acrescentar", str(erro))
            return
        self._depois_de_mudar_cadastro(f"{email} acrescentado ({', '.join(codigos) or 'sem empresa'}).")


def _abrir_arquivo(caminho: Path) -> None:
    """Abre o arquivo no programa padrão (Excel, navegador). Fora do Windows, só avisa onde ficou."""
    try:
        os.startfile(str(caminho))  # type: ignore[attr-defined]
    except (AttributeError, OSError):
        print(f"Arquivo gravado em {caminho}", file=sys.stderr)


class AcrescentarContato(tk.Toplevel):
    """E-mail + nome + empresa(s) da planilha. .resultado = (email, nome, [códigos]) ou None."""

    def __init__(self, pai: App, empresas: list[planilha.Empresa]):
        super().__init__(pai)
        self.title("Acrescentar e-mail")
        self.configure(bg=tema.FUNDO)
        self.geometry("680x560")
        self.transient(pai)
        self.resultado = None
        self.empresas = [e for e in empresas if not e.ex_cliente]
        self.escolhidas: list[str] = []
        self.var_email, self.var_nome, self.var_busca = tk.StringVar(), tk.StringVar(), tk.StringVar()
        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(1, weight=1)
        ttk.Label(corpo, text="E-mail").grid(row=0, column=0, sticky="w")
        entrada = ttk.Entry(corpo, textvariable=self.var_email)
        entrada.grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Label(corpo, text="Nome (opcional)").grid(row=1, column=0, sticky="w")
        ttk.Entry(corpo, textvariable=self.var_nome).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Label(corpo, text="Empresa(s) — procure pelo código ou nome e dê dois cliques:",
                  style="Apagado.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 2))
        ttk.Entry(corpo, textvariable=self.var_busca).grid(row=3, column=0, columnspan=2, sticky="ew")
        self.lista = tk.Listbox(corpo, font=(tema.FONTE, 10), activestyle="none", selectbackground=tema.ACENTO,
                                relief="flat", highlightthickness=1, highlightcolor=tema.LINHA, height=10)
        self.lista.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=6)
        corpo.rowconfigure(4, weight=1)
        self.var_escolhidas = tk.StringVar(value="Nenhuma empresa escolhida.")
        ttk.Label(corpo, textvariable=self.var_escolhidas, wraplength=620, justify="left").grid(
            row=5, column=0, columnspan=2, sticky="w")
        botoes = ttk.Frame(corpo)
        botoes.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(botoes, text="Limpar empresas", command=self._limpar).pack(side="left")
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(botoes, text="Acrescentar", style="Primario.TButton", command=self._ok).pack(side="right", padx=8)
        self.var_busca.trace_add("write", lambda *_: self._filtrar())
        self.lista.bind("<Double-1>", lambda _e: self._escolher())
        self.bind("<Escape>", lambda _e: self.destroy())
        self._filtrar()
        entrada.focus_set()
        self.grab_set()
        self.wait_window()

    def _filtrar(self):
        termo = K.normalizar(self.var_busca.get())
        digitado = self.var_busca.get().strip()
        self.visiveis = [e for e in self.empresas
                         if not termo or e.codigo == digitado or termo in K.normalizar(e.rotulo)][:400]
        self.lista.delete(0, "end")
        for e in self.visiveis:
            self.lista.insert("end", e.rotulo)

    def _escolher(self):
        sel = self.lista.curselection()
        if sel and self.visiveis[sel[0]].codigo not in self.escolhidas:
            self.escolhidas.append(self.visiveis[sel[0]].codigo)
            rotulos = [e.rotulo for e in self.empresas if e.codigo in self.escolhidas]
            self.var_escolhidas.set("Empresas: " + "; ".join(rotulos))

    def _limpar(self):
        self.escolhidas = []
        self.var_escolhidas.set("Nenhuma empresa escolhida.")

    def _ok(self):
        email = self.var_email.get().strip().lower()
        if not cadastro_mod.email_valido(email):
            messagebox.showwarning("Acrescentar", "Digite um e-mail válido.", parent=self)
            return
        if not self.escolhidas and not messagebox.askyesno(
                "Acrescentar", "Sem empresa, a ferramenta não vai saber de quem são as notas desse e-mail. "
                               "Acrescentar assim mesmo?", parent=self):
            return
        self.resultado = (email, self.var_nome.get().strip(), list(self.escolhidas))
        self.destroy()


class EscolherCliente(tk.Toplevel):
    """Lista de pastas de clientes com busca; devolve o nome da pasta escolhida em .resultado."""

    def __init__(self, pai: App, r: outlook.Resposta, palpite: K.Palpite, clientes: list[K.Cliente]):
        super().__init__(pai)
        self.title("De qual cliente é este e-mail?")
        self.configure(bg=tema.FUNDO)
        self.geometry("640x520")
        self.transient(pai)
        self.resultado: str | None = None
        self.clientes = clientes
        self.visiveis: list[str] = []
        self.var = tk.StringVar()
        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text=f"{r.nome}  <{r.email}>", style="Resumo.TLabel").pack(anchor="w")
        if pai.nomes.get(r.email):
            ttk.Label(corpo, text=f"No Outlook: {pai.nomes[r.email]}", style="Apagado.TLabel").pack(anchor="w")
        ttk.Label(corpo, text="Procure pelo código ou pelo nome:", style="Apagado.TLabel").pack(anchor="w", pady=(10, 2))
        busca = ttk.Entry(corpo, textvariable=self.var)
        busca.pack(fill="x")
        self.lista = tk.Listbox(corpo, font=(tema.FONTE, 10), activestyle="none", selectbackground=tema.ACENTO,
                                relief="flat", highlightthickness=1, highlightcolor=tema.LINHA)
        self.lista.pack(fill="both", expand=True, pady=8)
        botoes = ttk.Frame(corpo)
        botoes.pack(fill="x")
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(botoes, text="Usar este cliente", style="Primario.TButton", command=self._ok).pack(side="right", padx=8)
        ttk.Label(botoes, text="Fica guardado para os próximos meses.", style="Apagado.TLabel").pack(side="left")
        self.var.trace_add("write", lambda *_: self._filtrar())
        self.lista.bind("<Double-1>", lambda _e: self._ok())
        self.bind("<Return>", lambda _e: self._ok())
        self.bind("<Escape>", lambda _e: self.destroy())
        self._filtrar(palpite.pasta)
        busca.focus_set()
        self.grab_set()
        self.wait_window()

    def _filtrar(self, marcar: str | None = None):
        termo = K.normalizar(self.var.get())
        digitado = self.var.get().strip()
        self.visiveis = [c.pasta for c in self.clientes
                         if not termo or termo in K.normalizar(c.pasta) or (c.codigo and c.codigo == digitado)]
        self.lista.delete(0, "end")
        for pasta in self.visiveis[:500]:
            self.lista.insert("end", pasta)
        if marcar and marcar in self.visiveis:
            i = self.visiveis.index(marcar)
            self.lista.selection_set(i)
            self.lista.see(i)
        elif self.visiveis:
            self.lista.selection_set(0)

    def _ok(self):
        sel = self.lista.curselection()
        if sel:
            self.resultado = self.visiveis[sel[0]]
        self.destroy()


class DistribuirAnexos(tk.Toplevel):
    """E-mail de várias empresas: escolhe a empresa de cada anexo. .resultado = {índice: pasta} ou None."""

    def __init__(self, pai: App, r: outlook.Resposta, candidatos: list[str], atual: dict[int, str],
                 clientes: list[K.Cliente]):
        super().__init__(pai)
        self.title("De qual empresa é cada anexo?")
        self.configure(bg=tema.FUNDO)
        self.geometry("980x560")
        self.transient(pai)
        self.resultado: dict[int, str] | None = None
        self.r = r
        self.pai = pai
        self.clientes = clientes
        self.opcoes = list(candidatos)
        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text=f"{r.nome}  <{r.email}>", style="Resumo.TLabel").pack(anchor="w")
        ttk.Label(corpo, text=f"Este e-mail está em {len(candidatos)} empresas na planilha. "
                              "Onde achei o CNPJ da empresa no anexo, já deixei escolhido.",
                  style="Apagado.TLabel").pack(anchor="w", pady=(2, 10))
        todos = ttk.Frame(corpo)
        todos.pack(fill="x", pady=(0, 8))
        ttk.Label(todos, text="Todos para:").pack(side="left")
        self.var_todos = tk.StringVar()
        self.cb_todos = ttk.Combobox(todos, textvariable=self.var_todos, values=self.opcoes, state="readonly", width=70)
        self.cb_todos.pack(side="left", padx=8)
        self.cb_todos.bind("<<ComboboxSelected>>", lambda _e: self._todos())
        ttk.Button(todos, text="Outra empresa…", command=self._outra).pack(side="left")

        quadro = ttk.Frame(corpo, style="Cartao.TFrame", padding=10)
        quadro.pack(fill="both", expand=True)
        quadro.columnconfigure(1, weight=1)
        self.vars: dict[int, tk.StringVar] = {}
        self.combos = []
        for linha, a in enumerate(r.anexos):
            ttk.Label(quadro, text=a.nome, style="Painel.TLabel", width=48).grid(row=linha, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=atual.get(a.indice, ""))
            cb = ttk.Combobox(quadro, textvariable=var, values=self.opcoes, state="readonly")
            cb.grid(row=linha, column=1, sticky="ew", padx=(10, 0), pady=3)
            self.vars[a.indice] = var
            self.combos.append(cb)
        botoes = ttk.Frame(corpo)
        botoes.pack(fill="x", pady=(10, 0))
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(botoes, text="Usar", style="Primario.TButton", command=self._ok).pack(side="right", padx=8)
        self.bind("<Escape>", lambda _e: self.destroy())
        self.grab_set()
        self.wait_window()

    def _todos(self):
        for var in self.vars.values():
            var.set(self.var_todos.get())

    def _outra(self):
        escolhido = EscolherCliente(self.pai, self.r, K.Palpite(None, ""), self.clientes).resultado
        if escolhido and escolhido not in self.opcoes:
            self.opcoes.append(escolhido)
            for cb in self.combos + [self.cb_todos]:
                cb.configure(values=self.opcoes)
        if escolhido:
            self.var_todos.set(escolhido)

    def _ok(self):
        escolha = {i: v.get() for i, v in self.vars.items() if v.get()}
        if len(escolha) < len(self.vars) and not messagebox.askyesno(
                "Anexos", f"{len(self.vars) - len(escolha)} anexo(s) sem empresa não serão salvos. Continuar?",
                parent=self):
            return
        self.resultado = escolha
        self.destroy()


class BaseDeRespostas(tk.Toplevel):
    """Ver e apagar o que a ferramenta aprendeu."""

    def __init__(self, pai: App, base: sugestao.Base):
        super().__init__(pai)
        self.title("Respostas aprendidas")
        self.configure(bg=tema.FUNDO)
        self.geometry("1100x600")
        self.transient(pai)
        self.base = base
        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text="Cada linha é uma mensagem de cliente e a resposta aprovada. Mensagem parecida recebe "
                              "a mesma resposta como sugestão.", style="Apagado.TLabel").pack(anchor="w")
        self.tab = tema.Tabela(corpo, [("msg", "Mensagem do cliente", 470, "w", True),
                                       ("resp", "Resposta", 420, "w", True), ("cat", "Tipo", 100, "w", False),
                                       ("orig", "Origem", 80, "w", False)])
        self.tab.pack(fill="both", expand=True, pady=8)
        botoes = ttk.Frame(corpo)
        botoes.pack(fill="x")
        ttk.Button(botoes, text="Apagar a selecionada", command=self._apagar).pack(side="left")
        ttk.Button(botoes, text="Fechar", command=self.destroy).pack(side="right")
        self._preencher()

    def _preencher(self):
        self.tab.limpar()
        for i in range(len(self.base.exemplos) - 1, -1, -1):
            e = self.base.exemplos[i]
            self.tab.arvore.insert("", "end", iid=str(i), values=(
                re.sub(r"\s+", " ", e.mensagem)[:200], e.resposta.replace("\n", " / ")[:200], e.categoria, e.origem))

    def _apagar(self):
        sel = self.tab.arvore.selection()
        if sel and messagebox.askyesno("Apagar", "Esquecer esta resposta?", parent=self):
            self.base.esquecer(self.base.exemplos[int(sel[0])])
            self._preencher()


def main(argv: list[str] | None = None) -> int:
    App(argumentos(argv)).mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
