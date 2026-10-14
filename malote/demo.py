"""O fluxo inteiro sem tela, sobre o Outlook simulado: para ver o que a ferramenta faz e para a
integração contínua (roda em Linux, sem Outlook e sem Tkinter).

    python -m malote.demo                  roda numa pasta temporária e imprime o resumo
    python -m malote.demo --pasta ./demo   deixa a pasta para você abrir (clientes, planilha, dados)
    python -m malote.demo --hoje 2026-10-14

O que ele faz, na ordem da tela: monta a lista de envio (planilha + grupos + cadastro), lê tudo o que
chegou para a competência, identifica de quem é cada e-mail, decide o CNPJ dos anexos de e-mail
compartilhado, salva na pasta de cada cliente o que já está certo, sugere as respostas e grava a
planilha de controle. Devolve um dicionário com os números (os testes conferem cada um).
"""
from __future__ import annotations

import argparse
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

from . import anexos, falso, planilha, sugestao
from . import cadastro as cadastro_mod
from . import clientes as K
from . import competencia as C
from . import controle
from . import layouts as L
from .outlook import GRUPO_PADRAO, TAMANHO_LOTE
from .registro import Registro


def rodar(pasta: str | Path | None = None, hoje: date | None = None, mostrar=print) -> dict:
    demo = falso.Demonstracao(pasta, hoje)
    demo.preparar(limpar=True)
    comp = demo.comp
    dados = demo.dados
    registro, cadastro = Registro(dados), cadastro_mod.Cadastro(dados)
    base, layouts = sugestao.Base(dados), L.Layouts(dados)
    resumo: dict = {"competencia": C.texto(comp), "pasta": str(demo.pasta)}
    mostrar(f"Malote Fiscal — demonstração · competência {C.texto(comp)} · pasta {demo.pasta}")

    with demo.abrir() as ol:
        modelo, envios, grupos = ol.modelo(), ol.envios(), ol.grupos()
        nomes = ol.nomes_dos_contatos()
        layouts.garantir_padrao(modelo.html, modelo.competencia, texto=None if modelo.de_envio else L.TEXTO_PADRAO)
        for cliente_msg, nosso in ol.respostas_antigas():
            base.ensinar(cliente_msg, sugestao.miolo_de_resposta_antiga(nosso), sugestao.categoria(cliente_msg, "Ler"),
                         "histórico")
        empresas = planilha.ler(demo.planilha)
        clientes = K.indexar(demo.raiz)
        marcados = {g.nome for g in grupos if GRUPO_PADRAO.match(g.nome)}
        emails_grupos: dict[str, list[str]] = {}
        for g in grupos:
            if g.nome in marcados:
                for _n, email in g.membros:
                    emails_grupos.setdefault(email.lower(), []).append(g.nome)
                    nomes.setdefault(email.lower(), _n)
        cadastro.atualizar(empresas, emails_grupos, demo.planilha.name)
        desde = min(e.enviado_em for lista in envios.values() for e in lista)
        cadastro.anotar_respostas(ol.quem_respondeu(desde, envios), datetime.now())

        # ---- 1 · quem recebe
        lista = planilha.lista_de_envio(empresas, cadastro.fora_da_lista(), TAMANHO_LOTE, emails_grupos,
                                        cadastro.forcados(), {}, cadastro.respondentes(), cadastro.principais(),
                                        dominio_interno=ol.dominio)
        recebem = [d for d in lista.destinatarios if d.incluido]
        resumo["envio"] = {"recebem": len(recebem), "nao_recebem": len(lista.destinatarios) - len(recebem),
                           "lotes": len(lista.lotes), "sem_email": [e.codigo for e in lista.sem_email],
                           "ja_enviado": comp in envios, "aprendidas": len(base.exemplos),
                           "poupados": {d.email: d.observacao for d in lista.destinatarios
                                        if d.observacao.startswith(("a empresa responde", "principal"))},
                           "observacoes": {d.email: d.observacao for d in lista.destinatarios if not d.incluido}}
        mostrar(f"\n1 · Perguntar: {len(recebem)} e-mails em {len(lista.lotes)} mensagem(ns); "
                f"{len(lista.destinatarios) - len(recebem)} não recebem; {len(lista.sem_email)} empresa(s) sem e-mail; "
                f"{len(base.exemplos)} respostas antigas aprendidas"
                + ("; a pergunta deste mês JÁ FOI enviada (a tela avisa)" if comp in envios else ""))
        for d in lista.destinatarios:
            if not d.incluido:
                mostrar(f"      não recebe  {d.email:42} {d.observacao}")

        # ---- 2 · respostas
        mapa = planilha.por_email(empresas)
        ident = K.Identificador(clientes, registro.remetentes, nomes, mapa)
        perguntados = {e.lower() for envio in envios.get(comp, []) for _n, e in envio.destinatarios}
        conhecidos = set(cadastro.contatos) | set(mapa)
        respostas = ol.respostas(comp, envios, perguntados, conhecidos)
        palpites, sugeridos, destinos = {}, {}, {}
        for r in respostas:
            if r.tipo in ("Não entregue", "Interno"):
                continue
            p = ident.identificar(r.email, r.nome)
            palpites[r.entry_id] = p
            if len(p.candidatos) > 1 and r.anexos:
                cnpjs: dict[str, list[str]] = {}
                for e in mapa.get(r.email, []):
                    pasta_c = ident.pasta_da_empresa(e)
                    if pasta_c:
                        cnpjs.setdefault(pasta_c, []).append(e.cnpj)
                achados = ol.cnpjs_dos_anexos(r.entry_id, [a.indice for a in r.anexos])
                sugeridos[r.entry_id] = {i: anexos.qual_empresa(c, cnpjs) for i, c in achados.items()}
                if all(sugeridos[r.entry_id].get(a.indice) for a in r.anexos):
                    destinos[r.entry_id] = dict(sugeridos[r.entry_id])
            elif p.certo:
                destinos[r.entry_id] = {a.indice: p.pasta for a in r.anexos}
        mostrar(f"\n2 · Respostas e notas: {len(respostas)} e-mails lidos para {C.texto(comp)}")
        linhas = []
        for r in respostas:
            p = palpites.get(r.entry_id)
            if r.tipo == "Não entregue":
                quem = ", ".join(f"{e} ({m[0]}{'' if m[1] else ', não reenviar'})" for e, m in r.motivos.items())
            elif r.tipo == "Interno":
                quem = "(interno)"
            elif p and p.como == "várias empresas":
                sug = sorted({v for v in sugeridos.get(r.entry_id, {}).values() if v})
                quem = (" + ".join(sug) + " (CNPJ no anexo)") if sug else f"várias empresas ({len(p.candidatos)})"
            elif p and p.pasta:
                quem = p.pasta + ("" if p.certo else f"  [sugestão: {p.como}]")
            else:
                quem = "?"
            marca = "" if r.da_competencia else f"  ← competência {C.texto(r.competencia)}"
            onde = "" if r.pasta_outlook == "Caixa de entrada" else f"  [{r.pasta_outlook}]"
            linhas.append((r, quem))
            mostrar(f"      {r.recebido:%d/%m %H:%M}  {r.tipo:14} {r.relacao:9} {len(r.anexos) or '-':>2}  "
                    f"{r.nome[:26]:26} → {quem}{marca}{onde}")
        resumo["respostas"] = [{"id": r.entry_id, "email": r.email, "tipo": r.tipo, "relacao": r.relacao,
                                "anexos": len(r.anexos), "da_competencia": r.da_competencia,
                                "competencia": C.texto(r.competencia) if r.competencia else "",
                                "pasta_outlook": r.pasta_outlook, "quem": quem,
                                "nao_entregues": dict(r.motivos)} for r, quem in linhas]

        # quem não respondeu
        respondeu = {r.email for r in respostas if r.tipo not in ("Não entregue", "Interno") and r.da_competencia
                     and (r.relacao == "resposta" or r.anexos)}
        pastas_ok = {p for d in destinos.values() for p in d.values()}
        sem_resposta = [e for e in empresas if e.recebe and any(m in perguntados for m in e.emails)
                        and not any(m in respondeu for m in e.emails) and ident.pasta_da_empresa(e) not in pastas_ok]
        resumo["sem_resposta"] = [e.codigo for e in sem_resposta]
        mostrar("      sem resposta: " + (", ".join(e.rotulo for e in sem_resposta) or "ninguém"))

        # ---- 3 · salvar o que já está certo
        salvos = []
        for r in respostas:
            if r.tipo != "Mandou anexo" or r.entry_id not in destinos:
                continue
            por_pasta: dict[str, list[int]] = {}
            for indice, pasta_c in destinos[r.entry_id].items():
                por_pasta.setdefault(pasta_c, []).append(indice)
            comp_r = r.competencia or comp
            for pasta_c, indices in por_pasta.items():
                alvo = K.garantir(K.destino(demo.raiz, pasta_c, comp_r))
                feitos = ol.salvar_anexos(r.entry_id, indices, alvo)
                salvos.append({"email": r.email, "cliente": pasta_c, "destino": str(alvo.relative_to(demo.raiz)),
                               "arquivos": feitos, "competencia": C.texto(comp_r)})
                registro.marcar_salvo(r.entry_id, pasta_c, C.texto(comp_r), str(alvo), [n for n, _s in feitos])
        resumo["salvos"] = salvos
        mostrar(f"\n3 · Salvar anexos: {sum(len(s['arquivos']) for s in salvos)} arquivo(s) de {len(salvos)} e-mail/pasta")
        for s in salvos:
            for nome, situacao in s["arquivos"]:
                mostrar(f"      {situacao:10} {s['destino']}\\{nome}")

        # ---- 4 · sugestões de resposta
        sugestoes = []
        for r in respostas:
            if r.tipo in ("Não entregue", "Interno") or not r.relevante or r.respondido:
                continue
            s = base.sugerir(r.texto, r.tipo)
            sugestoes.append({"email": r.email, "categoria": s.categoria, "resposta": s.resposta, "porque": s.porque,
                              "saudacao": sugestao.saudacao(r.nome, datetime(2026, 1, 1, 9))})
        resumo["sugestoes"] = sugestoes
        mostrar(f"\n4 · Responder: {len(sugestoes)} resposta(s) sugeridas")
        for s in sugestoes:
            mostrar(f"      {s['email']:40} {s['categoria']:14} “{s['resposta'][:70]}”")

        # ---- 5 · planilha de controle
        arquivo = dados / "controle" / f"Tomados {comp[1]:02d}-{comp[0]}.xlsx"
        linhas_resp = [[r.recebido.strftime("%d/%m/%Y %H:%M"), quem, r.nome, r.email, r.tipo, len(r.anexos), "", "",
                        r.texto[:500]] for r, quem in linhas if r.tipo != "Não entregue"]
        controle.gerar(arquivo, linhas_resp, [[e.rotulo, e.resp_fiscal, ", ".join(e.emails)] for e in sem_resposta],
                       [[r.recebido.strftime("%d/%m/%Y %H:%M"), e, ""] for r in respostas for e in r.nao_entregues])
        resumo["planilha_controle"] = str(arquivo)
        mostrar(f"\n5 · Planilha de controle: {arquivo}")
    return resumo


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Roda a demonstração do Malote Fiscal sem tela.")
    p.add_argument("--pasta", help="onde montar o escritório fictício (padrão: pasta temporária)")
    p.add_argument("--hoje", help="data de referência AAAA-MM-DD (padrão: hoje)")
    args = p.parse_args(argv)
    hoje = date.fromisoformat(args.hoje) if args.hoje else None
    pasta = args.pasta or Path(tempfile.gettempdir()) / "malote-fiscal-demo"
    try:  # console cp1252 do Windows: imprime o que der em vez de parar numa seta
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    rodar(pasta, hoje)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
