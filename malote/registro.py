"""Memória da ferramenta, em JSON na pasta de dados.

registro.json
  remetentes: {e-mail: pasta do cliente} — confirmados na tela; é o que faz o mês seguinte
              sair sozinho.
  salvos:     {EntryID do e-mail: {cliente, competencia, destino, arquivos, quando}} — para
              não salvar duas vezes a mesma resposta.
  respondidos: {EntryID: {para, texto, quando}} — respostas enviadas pela ferramenta.
  reenviados:  {e-mail: "MM/AAAA em dd/mm/aaaa"} — reenvio depois de "não entregue".
config.json
  conta do Outlook, grupos marcados, pasta raiz dos clientes, planilha.
"""
from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path


def _ler(caminho: Path) -> dict:
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}


def _gravar(caminho: Path, dados: dict) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    tmp = caminho.with_suffix(".tmp")
    tmp.write_text(json.dumps(dados, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    os.replace(tmp, caminho)


class Registro:
    def __init__(self, pasta: Path):
        self.caminho = Path(pasta) / "registro.json"
        dados = _ler(self.caminho)
        self.remetentes: dict[str, str] = dados.get("remetentes", {})
        self.salvos: dict[str, dict] = dados.get("salvos", {})
        self.respondidos: dict[str, dict] = dados.get("respondidos", {})
        # versões antigas guardavam aqui quem estava fora da lista; hoje isso é do cadastro
        self.fora_da_lista: dict[str, str] = dados.get("fora_da_lista", {})
        self.incluir: set[str] = set(dados.get("incluir", []))
        self.reenviados: dict[str, str] = dados.get("reenviados", {})

    def aprender(self, email: str, pasta: str) -> None:
        self.remetentes[email.lower().strip()] = pasta
        self.gravar()

    def marcar_respondido(self, entry_id: str, para: str, texto: str) -> None:
        self.respondidos[entry_id] = {"para": para, "texto": texto,
                                      "quando": datetime.now().isoformat(timespec="seconds")}
        self.gravar()

    def marcar_reenviado(self, emails: list[str], competencia: str) -> None:
        for e in emails:
            self.reenviados[e.lower()] = f"{competencia} em {datetime.now():%d/%m/%Y %H:%M}"
        self.gravar()

    def marcar_salvo(self, entry_id: str, cliente: str, competencia: str, destino: str, arquivos: list[str]) -> None:
        self.salvos[entry_id] = {"cliente": cliente, "competencia": competencia, "destino": destino,
                                 "arquivos": arquivos, "quando": datetime.now().isoformat(timespec="seconds")}
        self.gravar()

    def gravar(self) -> None:
        _gravar(self.caminho, {"remetentes": self.remetentes, "salvos": self.salvos,
                               "respondidos": self.respondidos, "fora_da_lista": self.fora_da_lista,
                               "incluir": sorted(self.incluir), "reenviados": self.reenviados})


class Config:
    def __init__(self, pasta: Path):
        self.caminho = Path(pasta) / "config.json"
        self.dados = _ler(self.caminho)

    def get(self, chave, padrao=None):
        return self.dados.get(chave, padrao)

    def set(self, chave, valor) -> None:
        self.dados[chave] = valor
        _gravar(self.caminho, self.dados)
