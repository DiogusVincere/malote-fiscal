"""O Outlook simulado e o fluxo inteiro da demonstração, com data fixa (14/10/2026 → competência 09/2026)."""
import io
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from malote import anexos, demo, falso
from malote import competencia as C

HOJE = date(2026, 10, 14)
COMP = (2026, 9)


class OutlookFalso(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.d = falso.Demonstracao(Path(cls.tmp.name) / "demo", HOJE)
        cls.d.preparar()

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def respostas(self, ol):
        envios = ol.envios()
        perguntados = {e.lower() for envio in envios.get(COMP, []) for _n, e in envio.destinatarios}
        conhecidos = {m for _c, _r, lista, *_ in falso.EMPRESAS for m in lista}
        return {r.email + "|" + r.assunto: r for r in ol.respostas(COMP, envios, perguntados, conhecidos)}

    def test_dados_ficticios(self):
        self.assertEqual(falso.contas()[0].endereco, falso.CONTA)
        self.assertTrue(self.d.planilha.is_file())
        self.assertTrue((self.d.raiz / "101 - PADARIA AURORA LTDA" / "NFE" / "2026").is_dir())
        self.assertTrue((self.d.raiz / "116 - LAVANDERIA BRILHO LTDA").is_dir())
        self.assertFalse((self.d.raiz / "115 - GRAFICA IMPRESSAO RAPIDA LTDA").exists())
        c = falso.cnpj_da_empresa("101")
        self.assertEqual(len(c), 14)
        self.assertEqual(falso.formatar_cnpj(c)[2], ".")
        # dígitos verificadores batem com o algoritmo oficial
        base = [int(x) for x in c[:12]]
        pesos = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        r = sum(a * b for a, b in zip(base, pesos)) % 11
        self.assertEqual(int(c[12]), 0 if r < 2 else 11 - r)
        self.assertTrue(falso.pdf_simples(["a"]).startswith(b"%PDF-1.4"))
        self.assertTrue(falso.png_minusculo().startswith(b"\x89PNG"))

    def test_modelo_envios_grupos(self):
        with self.d.abrir() as ol:
            m = ol.modelo()
            self.assertEqual((m.competencia, m.de_envio), (COMP, True))
            self.assertIn("09/2026", m.html)
            envios = ol.envios()
            self.assertEqual(set(envios), {(2026, 7), (2026, 8), COMP})
            self.assertEqual(envios[COMP][0].enviado_em, datetime(2026, 10, 2, 14, 5))
            self.assertIn("antigo@rotanorte.example", [e for _n, e in envios[COMP][0].destinatarios])
            self.assertEqual([g.nome for g in ol.grupos()], ["Clientes Fiscal 1", "Clientes Fiscal 2", "Fornecedores"])
            self.assertEqual(len(ol.respostas_antigas()), 6)
            quem = ol.quem_respondeu(datetime(2026, 8, 1), envios)
            self.assertIn("financeiro@clinicavitalis.example", quem)
            self.assertNotIn("amigofiel@petshop.example", {e for e in quem if quem[e] > datetime(2026, 10, 1)})

    def test_leitura_da_competencia(self):
        with self.d.abrir() as ol:
            rs = self.respostas(ol)
        res = "RES: " + C.assunto(COMP)
        self.assertEqual(len(rs), 18)
        por_email = {}
        for r in rs.values():
            por_email.setdefault(r.email, []).append(r)
        # casos que a leitura antiga perdia
        fala = por_email["secretaria@falamundo.example"][0]
        self.assertEqual((fala.relacao, fala.tipo, len(fala.anexos)), ("contato", "Mandou anexo", 1))
        pet = por_email["amigofiel@petshop.example"][0]
        self.assertEqual((pet.relacao, pet.pasta_outlook), ("resposta", "Lixo eletrônico"))
        atlas = por_email["adm@atlasgestao.example"][0]
        self.assertEqual(atlas.pasta_outlook, "Caixa de entrada/Clientes")
        cedo = por_email["fiscal@rotanorte.example"][0]
        self.assertTrue(cedo.recebido < datetime(2026, 10, 1) and cedo.da_competencia)
        joao = por_email["joao@dois-irmaos.example"][0]
        self.assertEqual((joao.relacao, joao.tipo), ("outro", "Mandou anexo"))
        # o que não entra
        self.assertNotIn("premio@loteria.example", por_email)
        self.assertNotIn("fiscal@boasafra.example|RES: " + C.assunto((2026, 8)) + "x", rs)
        boa = sorted(por_email["fiscal@boasafra.example"], key=lambda r: r.recebido)
        self.assertEqual([(r.da_competencia, r.tipo) for r in boa], [(False, "Mandou anexo"), (True, "Sem serviços")])
        self.assertEqual(boa[0].competencia, (2026, 8))
        # tipos
        self.assertEqual(rs["financeiro@clinicavitalis.example|" + res].tipo, "Sem serviços")
        self.assertEqual(rs["oi@pixeldesign.example|" + res].tipo, "Mandou link")
        self.assertEqual(rs[falso.COLEGA + "|ENC: " + C.assunto(COMP)].tipo, "Interno")
        padaria = rs["financeiro@padariaaurora.example|" + res]
        self.assertEqual((padaria.tipo, len(padaria.anexos), padaria.embutidas), ("Mandou anexo", 1, 1))
        # devoluções
        dev = [r for r in rs.values() if r.tipo == "Não entregue"]
        motivos = {e: m for r in dev for e, m in r.motivos.items()}
        self.assertEqual(motivos["antigo@rotanorte.example"], ("endereço ou domínio não existe", False))
        self.assertEqual(motivos["academia@corpoemmovimento.example"], ("caixa do cliente cheia", True))

    def test_cnpj_dos_anexos_decide_a_empresa(self):
        with self.d.abrir() as ol:
            rs = self.respostas(ol)
            r = rs["contabil@grupoaurora.example|RES: " + C.assunto(COMP)]
            achados = ol.cnpjs_dos_anexos(r.entry_id, [a.indice for a in r.anexos])
            cand = {"101 - PADARIA AURORA LTDA": [falso.cnpj_da_empresa("101")],
                    "108 - RESTAURANTE SABOR DA SERRA LTDA": [falso.cnpj_da_empresa("108")]}
            self.assertEqual([anexos.qual_empresa(c, cand) for _i, c in sorted(achados.items())],
                             ["101 - PADARIA AURORA LTDA", "108 - RESTAURANTE SABOR DA SERRA LTDA"])

    def test_salvar_enviar_responder(self):
        d = falso.Demonstracao(Path(self.tmp.name) / "outra", HOJE)
        d.preparar()
        with d.abrir() as ol:
            rs = self.respostas(ol)
            r = rs["adm@atlasgestao.example|RES: " + C.assunto(COMP)]
            pasta = d.raiz / "112 - CONSULTORIA ATLAS GESTAO LTDA" / "NF-E" / "2026" / "092026" / "Tomados" / "Tomados e-mail"
            self.assertEqual(ol.salvar_anexos(r.entry_id, [1], pasta), [("NFSe 919 - ti.pdf", "já estava")])
            r2 = rs["financeiro@padariaaurora.example|RES: " + C.assunto(COMP)]
            destino = d.raiz / "101 - PADARIA AURORA LTDA" / "x"
            destino.mkdir()
            self.assertEqual(ol.salvar_anexos(r2.entry_id, [a.indice for a in r2.anexos], destino),
                             [("NFSe 1234 - Limpeza Total.pdf", "salvo")])
            self.assertEqual(ol.salvar_anexos(r2.entry_id, [1], destino), [("NFSe 1234 - Limpeza Total.pdf", "já estava")])
            # enviar: entra em Itens Enviados e passa a contar como envio da competência
            antes = len(ol.envios()[COMP])
            resumo = ol.montar("Serviços tomados - Setembro de 2026", "<html><body>oi</body></html>", [], ["a@b.example"],
                               enviar=True, so_testar=True)
            self.assertEqual(resumo["cco"], ["a@b.example"])
            self.assertEqual(len(ol.envios()[COMP]), antes)
            self.assertEqual(ol.montar("Serviços tomados - Setembro de 2026", "<html><body>oi</body></html>", [],
                                       ["a@b.example"], enviar=True), "enviado")
            self.assertEqual(len(ol.envios()[COMP]), antes + 1)
            self.assertEqual(ol.montar("x", "<html></html>", [], [], enviar=False), "Rascunhos (simulado)")
            # responder
            self.assertFalse(ol.ja_respondido(r2.entry_id))
            ol.responder(r2.entry_id, "Bom dia, Maria.\nOk.\nAtenciosamente,", ol.modelo())
            self.assertTrue(ol.ja_respondido(r2.entry_id))
            self.assertEqual(ol.respostas_antigas()[-1][0], "Bom dia! Segue a nota do mês. Att, Maria")


class FluxoCompleto(unittest.TestCase):
    def test_demo_rodar(self):
        with tempfile.TemporaryDirectory() as tmp:
            saida = io.StringIO()
            resumo = demo.rodar(Path(tmp) / "d", HOJE, mostrar=lambda *a: print(*a, file=saida))
            texto = saida.getvalue()
            self.assertEqual(resumo["competencia"], "09/2026")
            envio = resumo["envio"]
            self.assertEqual((envio["recebem"], envio["nao_recebem"], envio["lotes"]), (14, 4, 1))
            self.assertEqual(envio["sem_email"], ["106"])
            self.assertTrue(envio["ja_enviado"])
            self.assertEqual(envio["aprendidas"], 6)
            self.assertEqual(envio["observacoes"]["dra.helena@clinicavitalis.example"],
                             "a empresa responde por financeiro@clinicavitalis.example (102)")
            self.assertEqual(envio["observacoes"][falso.COLEGA], "e-mail interno")
            self.assertEqual(envio["observacoes"]["moveis@antiga.example"], "ex-cliente na planilha")
            self.assertEqual(len(resumo["respostas"]), 18)
            quem = {r["email"]: r["quem"] for r in resumo["respostas"] if r["da_competencia"]}
            self.assertEqual(quem["contabil@grupoaurora.example"],
                             "101 - PADARIA AURORA LTDA + 108 - RESTAURANTE SABOR DA SERRA LTDA (CNPJ no anexo)")
            self.assertIn("sugestão: mesmo domínio", quem["joao@dois-irmaos.example"])
            self.assertIn("sugestão: nome no e-mail", quem["corpoemmovimento.academia@provedor.example"])
            self.assertEqual(resumo["sem_resposta"], ["104", "107", "115"])
            salvos = {(s["cliente"], n, sit) for s in resumo["salvos"] for n, sit in s["arquivos"]}
            self.assertIn(("112 - CONSULTORIA ATLAS GESTAO LTDA", "NFSe 919 - ti.pdf", "já estava"), salvos)
            self.assertIn(("101 - PADARIA AURORA LTDA", "NFSe 88 - marketing.xml", "salvo"), salvos)
            self.assertIn(("108 - RESTAURANTE SABOR DA SERRA LTDA", "NFSe 301 - ti.xml", "salvo"), salvos)
            self.assertIn(("110 - ESCOLA DE IDIOMAS FALA MUNDO LTDA", "notas setembro.zip", "salvo"), salvos)
            agosto = next(s for s in resumo["salvos"] if s["competencia"] == "08/2026")
            self.assertIn("08-2026", agosto["destino"])
            self.assertTrue((Path(tmp) / "d" / "Clientes" / "101 - PADARIA AURORA LTDA" / "NFE" / "2026" / "09-2026"
                             / "Tomados" / "Tomados e-mail" / "NFSe 1234 - Limpeza Total.pdf").is_file())
            sug = {s["email"]: s for s in resumo["sugestoes"]}
            self.assertEqual(sug["amigofiel@petshop.example"]["categoria"], "Pergunta")
            self.assertTrue(sug["amigofiel@petshop.example"]["resposta"].startswith("Pedágio e estacionamento não entram"))
            self.assertEqual(sug["moveis@antiga.example"]["categoria"], "Sair da lista")
            self.assertEqual(sug["farmacia@bemestar.example"]["categoria"], "Trocar e-mail")
            self.assertEqual(sug["financeiro@padariaaurora.example"]["saudacao"], "Bom dia, Maria.")
            self.assertTrue(Path(resumo["planilha_controle"]).is_file())
            self.assertIn("5 · Planilha de controle", texto)
            # rodar de novo na mesma pasta recomeça do zero (a demonstração sempre reinicia)
            resumo2 = demo.rodar(Path(tmp) / "d", HOJE, mostrar=lambda *a: None)
            self.assertEqual(len(resumo2["salvos"]), len(resumo["salvos"]))


if __name__ == "__main__":
    unittest.main()
