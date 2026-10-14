"""A lógica de leitura do outlook.py, que não toca no COM: roda em qualquer sistema."""
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from malote import outlook as O
from malote import competencia as C

CONTA = "fiscal@contabil.example"
COMP = (2026, 9)
ENVIOS = {COMP: [O.Envio(COMP, datetime(2026, 10, 2, 14, 5), [("Cliente", "a@cliente.example")])],
          (2026, 8): [O.Envio((2026, 8), datetime(2026, 9, 2, 14, 5), [])]}


def item(email="a@cliente.example", assunto="RES: Serviços tomados - Setembro de 2026", corpo="Segue.",
         anexos=(), recebido=datetime(2026, 10, 5, 9, 0), classe="IPM.Note", nome="Cliente", html="", **extra):
    return O.ItemLido(f"id-{email}-{assunto[:12]}-{recebido:%d%H%M}", recebido, assunto, classe, corpo, html, email, nome,
                      list(anexos), **extra)


def contexto(**kw):
    return O.contexto_de(CONTA, ENVIOS, kw.pop("assinaturas", None), kw.pop("perguntados", ()),
                         kw.pop("conhecidos", ()))


class Marcas(unittest.TestCase):
    def test_marcas_e_sobre_a_pergunta(self):
        marcas = O.marcas_da_pergunta(CONTA, ENVIOS, [(2026, 7)])
        self.assertIn(CONTA, marcas)
        self.assertIn(O.FRASE_DA_PERGUNTA, marcas)
        self.assertIn("servicos tomados - julho de 2026", marcas)
        self.assertTrue(O.e_sobre_a_pergunta("RES: Serviços TOMADOS - Setembro de 2026", "", marcas))
        self.assertTrue(O.e_sobre_a_pergunta("Notas", "Sua empresa tomou (contratou) algum serviço?", marcas))
        self.assertTrue(O.e_sobre_a_pergunta("Notas", f"De: {CONTA}\nblabla", marcas))
        self.assertFalse(O.e_sobre_a_pergunta("Notas de setembro", "Seguem as notas.", marcas))

    def test_contexto_normaliza(self):
        ctx = O.contexto_de("Fiscal@Contabil.example", {}, {100}, ["A@X.example"], ["B@Y.example"])
        self.assertEqual(ctx.endereco, "fiscal@contabil.example")
        self.assertEqual(ctx.dominio, "contabil.example")
        self.assertEqual(ctx.perguntados, {"a@x.example"})
        self.assertEqual(ctx.conhecidos, {"b@y.example"})


class Assinatura(unittest.TestCase):
    def test_e_assinatura(self):
        html = '<img src="cid:image001.png@01DC">'
        self.assertTrue(O.e_assinatura(O.AnexoLido(1, "image001.png", 900_000, "image001.png@01DC"), html, set()))
        self.assertTrue(O.e_assinatura(O.AnexoLido(1, "logo.png", 50, "", oculto=True), "", set()))
        self.assertTrue(O.e_assinatura(O.AnexoLido(1, "image002.jpg", 20_000), "", set()))
        self.assertTrue(O.e_assinatura(O.AnexoLido(1, "image001.png", 1_010_000), "", {1_000_000}))
        self.assertFalse(O.e_assinatura(O.AnexoLido(1, "image001.png", 900_000), "", {100_000}))
        self.assertFalse(O.e_assinatura(O.AnexoLido(1, "nota.pdf", 20_000), "", set()))
        self.assertFalse(O.e_assinatura(O.AnexoLido(1, "foto da nota.png", 20_000), "", set()))


class LerItem(unittest.TestCase):
    def test_ignora_o_que_nao_e_de_fora(self):
        ctx = contexto()
        self.assertIsNone(O.ler_item(item(email=CONTA), ctx))
        self.assertIsNone(O.ler_item(item(classe="IPM.Schedule.Meeting.Request"), ctx))
        self.assertIsNone(O.ler_item(item(email=""), ctx))

    def test_relacao(self):
        ctx = contexto(conhecidos=["b@cliente.example"])
        self.assertEqual(O.ler_item(item(), ctx).relacao, "resposta")
        self.assertEqual(O.ler_item(item(email="b@cliente.example", assunto="Notas", corpo="Seguem."), ctx).relacao,
                         "contato")
        r = O.ler_item(item(email="z@outro.example", assunto="Promoção", corpo="Compre."), ctx)
        self.assertEqual((r.relacao, r.tipo, r.relevante), ("outro", "Ler", False))

    def test_tipos(self):
        ctx = contexto(assinaturas={900_000})
        anexos = [O.AnexoLido(1, "nota.pdf", 5000), O.AnexoLido(2, "image001.png", 20_000, "img@1")]
        r = O.ler_item(item(anexos=anexos, html='<img src="cid:img@1">'), ctx)
        self.assertEqual((r.tipo, len(r.anexos), r.embutidas, r.anexos[0].indice), ("Mandou anexo", 1, 1, 1))
        self.assertEqual(O.ler_item(item(corpo="As notas estão em https://we.tl/t-abc"), ctx).tipo, "Mandou link")
        self.assertEqual(O.ler_item(item(corpo="Não tivemos serviços tomados em setembro."), ctx).tipo, "Sem serviços")
        self.assertEqual(O.ler_item(item(corpo="Não houve."), ctx).tipo, "Sem serviços")
        self.assertEqual(O.ler_item(item(corpo="Sim, tivemos. Segue em anexo."), ctx).tipo, "Ler")
        self.assertEqual(O.ler_item(item(corpo="Tivemos sim, mas não achei a nota."), ctx).tipo, "Ler")
        # "não" fora de uma resposta à pergunta não quer dizer "sem serviços"
        self.assertEqual(O.ler_item(item(email="z@outro.example", assunto="Oi", corpo="Não."), ctx).tipo, "Ler")
        self.assertEqual(O.ler_item(item(email="ana@contabil.example", anexos=anexos[:1]), ctx).tipo, "Interno")

    def test_texto_sem_a_citacao(self):
        corpo = "Segue a nota.\n\nAtt,\nMaria\n\nDe: fiscal@contabil.example\nEnviado: ontem\n\nSua empresa tomou..."
        r = O.ler_item(item(corpo=corpo, respondido=True), contexto())
        self.assertEqual(r.texto, "Segue a nota.\nAtt,\nMaria")
        self.assertTrue(r.respondido)
        self.assertEqual(r.pasta_outlook, "Caixa de entrada")

    def test_devolucao(self):
        corpo = ("This is the mail system.\n\n<antigo@rotanorte.example>: host mx said: 550 5.1.1 User unknown\n")
        d = item(email="", nome="Mail Delivery System", assunto="Undelivered Mail Returned to Sender", corpo=corpo,
                 classe="REPORT.IPM.Note.NDR")
        # endereço perguntado: entra, com motivo e sem reenviar
        r = O.ler_item(d, contexto(perguntados=["antigo@rotanorte.example"]))
        self.assertEqual((r.tipo, r.relacao, r.nao_entregues), ("Não entregue", "resposta", ["antigo@rotanorte.example"]))
        self.assertEqual(r.motivos["antigo@rotanorte.example"], ("endereço ou domínio não existe", False))
        # endereço que não foi perguntado: devolução de outro e-mail qualquer
        self.assertIsNone(O.ler_item(d, contexto(perguntados=["outro@x.example"])))
        # sem lista de perguntados, só entra se o aviso citar a pergunta
        self.assertIsNone(O.ler_item(d, contexto()))
        d2 = item(email="postmaster@mail.example", assunto="Delivery report",
                  corpo="Final-Recipient: rfc822;x@y.example\nStatus: 5.2.2 mailbox full\nSubject: Serviços tomados - Setembro de 2026")
        r = O.ler_item(d2, contexto())
        self.assertEqual(r.motivos["x@y.example"], ("caixa do cliente cheia", True))


class Competencias(unittest.TestCase):
    def test_atribuir_competencia(self):
        inicios = [(datetime(2026, 9, 2), (2026, 8)), (datetime(2026, 10, 2), (2026, 9))]
        r = O.Resposta("1", datetime(2026, 10, 5), "n", "e", "RES: Serviços tomados - Agosto de 2026", "")
        self.assertEqual(O.atribuir_competencia(r, COMP, inicios), ((2026, 8), "assunto"))
        r = O.Resposta("2", datetime(2026, 10, 5), "n", "e", "Notas", "", anexos=[O.Anexo(1, "notas setembro.zip", 1)])
        self.assertEqual(O.atribuir_competencia(r, COMP, inicios), ((2026, 9), "anexo"))
        r = O.Resposta("3", datetime(2026, 9, 20), "n", "e", "Notas", "")
        self.assertEqual(O.atribuir_competencia(r, COMP, inicios), ((2026, 8), "data"))
        r = O.Resposta("4", datetime(2026, 8, 1), "n", "e", "Notas", "")
        self.assertEqual(O.atribuir_competencia(r, COMP, inicios), (COMP, "data"))

    def test_janela_de_leitura(self):
        self.assertEqual(O.janela_de_leitura(COMP, ENVIOS), (datetime(2026, 9, 1), datetime(2026, 10, 1)))
        cedo = {COMP: [O.Envio(COMP, datetime(2026, 9, 29, 8), [])]}
        self.assertEqual(O.janela_de_leitura(COMP, cedo), (datetime(2026, 9, 1), datetime(2026, 9, 29, 8)))
        self.assertEqual(O.janela_de_leitura((2026, 12), {}), (datetime(2026, 12, 1), datetime(2027, 1, 1)))


class Selecionar(unittest.TestCase):
    def test_regras_da_janela_e_do_lixo(self):
        ctx = contexto(conhecidos=["b@cliente.example"])
        agosto_em_setembro = item(email="b@cliente.example", assunto="RES: Serviços tomados - Agosto de 2026",
                                  recebido=datetime(2026, 9, 3))
        setembro_adiantado = item(email="b@cliente.example", assunto="RES: Serviços tomados - Setembro de 2026",
                                  recebido=datetime(2026, 9, 28))
        sem_assunto_em_setembro = item(email="b@cliente.example", assunto="Nota", recebido=datetime(2026, 9, 28))
        normal = item(recebido=datetime(2026, 10, 3))
        outra_comp = item(assunto="RES: Serviços tomados - Agosto de 2026", recebido=datetime(2026, 10, 1, 8))
        spam_no_lixo = item(email="z@spam.example", assunto="Ganhe", corpo="x", recebido=datetime(2026, 10, 4))
        cliente_no_lixo = item(email="b@cliente.example", assunto="Nota", corpo="segue", recebido=datetime(2026, 10, 4),
                               pasta="Lixo eletrônico")
        repetido = item(recebido=datetime(2026, 10, 3))  # mesmo entry_id de `normal`
        itens = [(agosto_em_setembro, False), (setembro_adiantado, False), (sem_assunto_em_setembro, False),
                 (normal, False), (outra_comp, False), (spam_no_lixo, True), (cliente_no_lixo, True), (repetido, False)]
        achadas = O.selecionar_respostas(itens, COMP, ENVIOS, ctx)
        ids = [r.entry_id for r in achadas]
        self.assertNotIn(agosto_em_setembro.entry_id, ids)
        self.assertNotIn(sem_assunto_em_setembro.entry_id, ids)
        self.assertNotIn(spam_no_lixo.entry_id, ids)
        self.assertIn(setembro_adiantado.entry_id, ids)
        self.assertIn(cliente_no_lixo.entry_id, ids)
        self.assertEqual(ids.count(normal.entry_id), 1)
        por_id = {r.entry_id: r for r in achadas}
        self.assertFalse(por_id[outra_comp.entry_id].da_competencia)
        self.assertEqual(por_id[outra_comp.entry_id].competencia, (2026, 8))
        self.assertTrue(por_id[setembro_adiantado.entry_id].da_competencia)
        # do mais novo para o mais velho
        self.assertEqual([r.recebido for r in achadas], sorted((r.recebido for r in achadas), reverse=True))


class Gravar(unittest.TestCase):
    def test_gravar_na_pasta_nunca_sobrescreve(self):
        with tempfile.TemporaryDirectory() as tmp:
            pasta, origem = Path(tmp) / "destino", Path(tmp) / "origem"
            pasta.mkdir()
            origem.mkdir()
            a = origem / "a.bin"
            a.write_bytes(b"conteudo 1")
            self.assertEqual(O.gravar_na_pasta(a, "nota: final?.pdf", pasta), ("nota_ final_.pdf", "salvo"))
            self.assertEqual(O.gravar_na_pasta(a, "nota: final?.pdf", pasta), ("nota_ final_.pdf", "já estava"))
            b = origem / "b.bin"
            b.write_bytes(b"conteudo 2!")
            self.assertEqual(O.gravar_na_pasta(b, "nota: final?.pdf", pasta), ("nota_ final_ (2).pdf", "salvo"))
            self.assertEqual(O.gravar_na_pasta(b, "nota: final?.pdf", pasta), ("nota_ final_ (2).pdf", "já estava"))
            # o mesmo conteúdo com outro nome também é "já estava"
            self.assertEqual(O.gravar_na_pasta(a, "outro nome.pdf", pasta), ("nota_ final_.pdf", "já estava"))
            self.assertEqual(sorted(p.name for p in pasta.iterdir()), ["nota_ final_ (2).pdf", "nota_ final_.pdf"])

    def test_nome_seguro(self):
        self.assertEqual(O._nome_seguro('a<b>:"c"/d\\e|f?g*.pdf', Path("C:/x")), "a_b___c__d_e_f_g_.pdf")
        self.assertEqual(O._nome_seguro("", Path("C:/x")), "anexo")
        pasta = Path("C:/" + "p" * 100)
        longo = O._nome_seguro("n" * 300 + ".pdf", pasta)
        self.assertTrue(longo.endswith(".pdf"))
        self.assertEqual(len(longo), 250 - len(str(pasta)) - 6)


class Antigas(unittest.TestCase):
    def test_par_de_resposta_antiga(self):
        corpo = ("Bom dia, Maria.\nEspero que você esteja bem.\nOk, recebido.\nAtenciosamente,\n\n"
                 "De: Maria <m@x.example>\nEnviado: ontem\nAssunto: RES: Serviços tomados - Agosto de 2026\n\n"
                 "Segue a nota do mês.\n\nDe: fiscal@contabil.example\nSua empresa tomou...")
        par = O.par_de_resposta_antiga("RES: Serviços tomados - Agosto de 2026", corpo)
        self.assertEqual(par[0], "Segue a nota do mês.")
        self.assertTrue(par[1].startswith("Bom dia, Maria."))
        self.assertIsNone(O.par_de_resposta_antiga("ENC: Serviços tomados - Agosto de 2026", corpo))
        self.assertIsNone(O.par_de_resposta_antiga("RES: Outra coisa", corpo))
        self.assertIsNone(O.par_de_resposta_antiga("RES: Serviços tomados - Agosto de 2026", "sem citação"))


class ModeloPadrao(unittest.TestCase):
    def test_modelo_padrao(self):
        m = O.modelo_padrao((2026, 9))
        self.assertFalse(m.de_envio)
        self.assertEqual(m.assunto, C.assunto((2026, 9)))
        self.assertIn(O.FRASE_DA_PERGUNTA, m.html)
        self.assertIn("09/2026", m.html)
        self.assertEqual(m.imagens, [])


if __name__ == "__main__":
    unittest.main()
