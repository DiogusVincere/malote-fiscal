"""Tema visual do Malote Fiscal: paleta, estilos ttk e a classe Tabela num módulo só."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

# Paleta: azul-ardósia como acento sobre cinzas neutros
ESCURO, ACENTO, ACENTO_D, ACENTO_DD = "#1F2A37", "#2F5D8A", "#264B70", "#1E3C5A"
FUNDO, PAINEL, TINTA, APAGADO, LINHA = "#F4F5F7", "#FFFFFF", "#1C2430", "#5F6B7A", "#E1E5EA"
CLARO, ACENTO_CLARO, ZEBRA = "#EAEEF3", "#9DB8D4", "#F9FAFB"

# Estado: pronto / atenção / problema
VERDE, VERDE_T = "#1F7A4D", "#DFF2E6"
AMBAR, AMBAR_T = "#8C6914", "#FBF0D9"
VERMELHO, VERMELHO_T = "#A63D30", "#F7E2DF"

FONTE = "Segoe UI"
MARCA = "MALOTE FISCAL"

# As mesmas cores em ARGB, para quem grava Excel com openpyxl
EXCEL_CABECALHO_FUNDO = "FF1F2A37"
EXCEL_CABECALHO_TEXTO = "FFEAEEF3"
EXCEL_ZEBRA = "FFF4F5F7"
EXCEL_VERMELHO_T = "FFF7E2DF"
EXCEL_AMBAR_T = "FFFBF0D9"
EXCEL_VERDE_T = "FFDFF2E6"


def estilos(janela: tk.Misc) -> ttk.Style:
    """Aplica os estilos ttk da suíte. Chame uma vez, logo depois de criar a janela."""
    s = ttk.Style(janela)
    s.theme_use("clam")
    s.configure(".", background=FUNDO, foreground=TINTA, font=(FONTE, 10))
    s.configure("TFrame", background=FUNDO)
    s.configure("Painel.TFrame", background=PAINEL)
    s.configure("Cartao.TFrame", background=PAINEL, relief="solid", borderwidth=1, bordercolor=LINHA)
    s.configure("TLabel", background=FUNDO, foreground=TINTA)
    s.configure("Painel.TLabel", background=PAINEL, foreground=TINTA)
    s.configure("Passo.TLabel", foreground=ACENTO_D, font=(FONTE, 10, "bold"))
    s.configure("Apagado.TLabel", foreground=APAGADO, font=(FONTE, 9))
    s.configure("ApagadoPainel.TLabel", background=PAINEL, foreground=APAGADO, font=(FONTE, 9))
    s.configure("Titulo.TLabel", background=PAINEL, foreground=TINTA, font=(FONTE, 12, "bold"))
    s.configure("Resumo.TLabel", font=(FONTE, 10, "bold"))
    s.configure("Status.TLabel", background=CLARO, foreground=APAGADO, font=(FONTE, 9), padding=(12, 5))
    s.configure("TButton", padding=(12, 6), background=PAINEL, bordercolor="#CBD2DA")
    s.map("TButton", background=[("active", CLARO)])
    s.configure("Primario.TButton", background=ACENTO_D, foreground="#FFFFFF", bordercolor=ACENTO_D,
                font=(FONTE, 10, "bold"), padding=(18, 7))
    s.map("Primario.TButton",
          background=[("disabled", "#B3BEC9"), ("pressed", ACENTO_DD), ("active", ACENTO_DD)],
          foreground=[("disabled", "#F4F5F7")], bordercolor=[("disabled", "#B3BEC9")])
    s.configure("TEntry", fieldbackground=PAINEL, padding=5, bordercolor="#CBD2DA")
    s.configure("TCombobox", fieldbackground=PAINEL, padding=5, bordercolor="#CBD2DA")
    s.map("TCombobox", fieldbackground=[("readonly", PAINEL)], selectbackground=[("readonly", PAINEL)],
          selectforeground=[("readonly", TINTA)])
    s.configure("Treeview", background=PAINEL, fieldbackground=PAINEL, foreground=TINTA, rowheight=24,
                bordercolor=LINHA, font=(FONTE, 9))
    s.configure("Treeview.Heading", background=ESCURO, foreground=CLARO, relief="flat",
                font=(FONTE, 9, "bold"), padding=(6, 5))
    s.map("Treeview.Heading", background=[("active", ACENTO_DD)])
    s.map("Treeview", background=[("selected", ACENTO)], foreground=[("selected", "#FFFFFF")])
    s.configure("TNotebook", background=FUNDO, borderwidth=0, tabmargins=(0, 4, 0, 0))
    s.configure("TNotebook.Tab", padding=(14, 6), background=CLARO, foreground=APAGADO, font=(FONTE, 10))
    s.map("TNotebook.Tab", background=[("selected", PAINEL)], foreground=[("selected", TINTA)])
    s.configure("Horizontal.TProgressbar", background=ACENTO, troughcolor=CLARO, bordercolor=CLARO)
    s.configure("TCheckbutton", background=FUNDO, foreground=TINTA)
    s.map("TCheckbutton", background=[("active", FUNDO)])
    s.configure("TRadiobutton", background=FUNDO, foreground=TINTA)
    s.map("TRadiobutton", background=[("active", FUNDO)])
    return s


def faixa(janela: tk.Misc, titulo: str, subtitulo: str = "", canto: str = "",
          marca: str = MARCA) -> tk.Frame:
    """Faixa escura do topo, com o nome da suíte e o título da tela."""
    topo = tk.Frame(janela, bg=ESCURO)
    topo.pack(fill="x")
    tk.Label(topo, text=marca, bg=ESCURO, fg=ACENTO_CLARO,
             font=(FONTE, 9, "bold")).pack(anchor="w", padx=20, pady=(12, 0))
    tk.Label(topo, text=titulo, bg=ESCURO, fg="#FFFFFF",
             font=(FONTE, 15, "bold")).pack(anchor="w", padx=20, pady=(0, 2 if subtitulo else 12))
    if subtitulo:
        tk.Label(topo, text=subtitulo, bg=ESCURO, fg=ACENTO_CLARO,
                 font=(FONTE, 9)).pack(anchor="w", padx=20, pady=(0, 12))
    if canto:
        tk.Label(topo, text=canto, bg=ESCURO, fg=ACENTO_CLARO,
                 font=(FONTE, 9)).place(relx=1, x=-20, y=14, anchor="ne")
    return topo


def barra_status(janela: tk.Misc, variavel: tk.StringVar) -> ttk.Frame:
    """Barra cinza do rodapé com o texto de status."""
    base = ttk.Frame(janela, style="TFrame")
    base.pack(fill="x", side="bottom")
    ttk.Label(base, textvariable=variavel, style="Status.TLabel", wraplength=1100,
              justify="left").pack(side="left", fill="x", expand=True)
    return base


def etiqueta(pai: tk.Misc, texto: str, cor_fundo: str, cor_texto: str) -> tk.Label:
    """Selo colorido curto (verde, âmbar ou vermelho) para marcar um estado."""
    return tk.Label(pai, text=texto, bg=cor_fundo, fg=cor_texto, font=(FONTE, 8, "bold"),
                    padx=8, pady=2)


class Tabela(ttk.Frame):
    """Treeview com barras de rolagem.

    colunas = lista de (id, título, largura, alinhamento, estica).
    """

    def __init__(self, pai, colunas):
        super().__init__(pai, style="Painel.TFrame")
        ids = [c[0] for c in colunas]
        self.arvore = ttk.Treeview(self, columns=ids, show="headings", selectmode="browse")
        for ident, titulo, largura, ancora, estica in colunas:
            self.arvore.heading(ident, text=titulo, anchor=ancora)
            self.arvore.column(ident, width=largura, minwidth=40, anchor=ancora, stretch=estica)
        vy = ttk.Scrollbar(self, orient="vertical", command=self.arvore.yview)
        vx = ttk.Scrollbar(self, orient="horizontal", command=self.arvore.xview)
        self.arvore.configure(yscrollcommand=vy.set, xscrollcommand=vx.set)
        self.arvore.grid(row=0, column=0, sticky="nsew")
        vy.grid(row=0, column=1, sticky="ns")
        vx.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.arvore.tag_configure("par", background=ZEBRA)
        self.arvore.tag_configure("alerta", background=VERMELHO_T)
        self.arvore.tag_configure("bom", background=VERDE_T)
        self._linhas = 0

    def limpar(self):
        self.arvore.delete(*self.arvore.get_children())
        self._linhas = 0

    def inserir(self, valores, tags=()):
        # o contador é guardado de propósito: pedir get_children() a cada linha faz a
        # inserção virar O(n²) e a janela congela num relatório de milhares de linhas
        tags = tuple(tags) or (("par",) if self._linhas % 2 else ())
        self.arvore.insert("", "end", values=valores, tags=tags)
        self._linhas += 1

    @property
    def vazia(self) -> bool:
        return self._linhas == 0
