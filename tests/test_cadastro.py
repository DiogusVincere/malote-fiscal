import json
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

from malote import cadastro as M
from malote.planilha import Empresa

EMPRESAS = [Empresa("101", "PADARIA AURORA LTDA", "", ["a@aurora.example", "b@aurora.example"]),
            Empresa("102", "CLINICA LTDA", "", ["fin@clinica.example"]),
            Empresa("000 EX-CLIENTE", "ANTIGA", "", ["antiga@antiga.example"], ex_cliente=True)]


class Cadastro(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.pasta = Path(self.tmp.name)
        self.cad = M.Cadastro(self.pasta)

    def tearDown(self):
        self.tmp.cleanup()

    def test_atualizar_pela_planilha_e_grupos(self):
        mudou = self.cad.atualizar(EMPRESAS, {"g@grupo.example": ["Clientes 1"], "a@aurora.example": ["Clientes 1"]},
                                   "Clientes 2026-10-01.xlsx")
        self.assertEqual(sorted(mudou["novos"]), ["a@aurora.example", "b@aurora.example", "fin@clinica.example",
                                                  "g@grupo.example"])
        self.assertEqual(mudou["sumiram"], [])
        self.assertNotIn("antiga@antiga.example", self.cad.contatos)
        a = self.cad["A@aurora.example"]
        self.assertEqual((a.codigos, a.nome, a.origens), (["101"], "PADARIA AURORA LTDA", ["planilha", "grupo:Clientes 1"]))
        self.assertEqual(self.cad["g@grupo.example"].origens, ["grupo:Clientes 1"])
        # some da planilha: sinalizado, não apagado; volta: limpa
        mudou = self.cad.atualizar(EMPRESAS[1:], {}, "Clientes 2026-11-01.xlsx")
        self.assertEqual(sorted(mudou["sumiram"]), ["a@aurora.example", "b@aurora.example"])
        self.assertTrue(self.cad["a@aurora.example"].sumiu_em)
        self.cad.atualizar(EMPRESAS, {}, "Clientes 2026-12-01.xlsx")
        self.assertEqual(self.cad["a@aurora.example"].sumiu_em, "")
        self.assertTrue(self.cad.ultima_atualizacao.startswith("Clientes 2026-12-01.xlsx em "))
        # persistiu em disco
        de_novo = M.Cadastro(self.pasta)
        self.assertEqual(set(de_novo.contatos), set(self.cad.contatos))
        self.assertEqual(de_novo["a@aurora.example"].historico[0]["o_que"], "entrou pela planilha (101)")

    def test_mudancas_do_usuario(self):
        self.cad.atualizar(EMPRESAS, {}, "p.xlsx")
        self.cad.tirar("a@aurora.example", "pediu para sair")
        self.assertEqual(self.cad.fora_da_lista(), {"a@aurora.example": "pediu para sair"})
        self.cad.voltar("a@aurora.example")
        self.assertEqual(self.cad.fora_da_lista(), {})
        self.cad.marcar("novo@x.example", True, padrao=False)
        self.assertEqual(self.cad.forcados(), {"novo@x.example"})
        self.cad.marcar("b@aurora.example", False, padrao=True)
        self.assertIn("b@aurora.example", self.cad.fora_da_lista())
        c = self.cad.acrescentar("Manual@X.example", ["102"], "Fulano")
        self.assertEqual((c.email, c.codigos, c.manual), ("manual@x.example", ["102"], True))
        self.assertEqual(self.cad.manuais(), {"manual@x.example": ["102"]})
        with self.assertRaises(ValueError):
            self.cad.acrescentar("sem-arroba", [])
        novo = self.cad.trocar("fin@clinica.example", "contabil@clinica.example")
        self.assertEqual(novo.codigos, ["102"])
        self.assertTrue(self.cad["fin@clinica.example"].fora)
        self.assertEqual(self.cad["fin@clinica.example"].motivo, "trocado por contabil@clinica.example")
        self.assertTrue(any(h["o_que"] == "substitui fin@clinica.example" for h in novo.historico))

    def test_quem_responde_e_principal(self):
        self.cad.atualizar(EMPRESAS, {}, "p.xlsx")
        agora = datetime.now()
        self.assertEqual(self.cad.anotar_respostas({"a@aurora.example": agora, "x@naoexiste.example": agora}, agora), 1)
        self.assertEqual(self.cad.respondentes(), {"a@aurora.example"})
        self.assertEqual(self.cad.respostas_lidas_ate, agora.strftime("%Y-%m-%d %H:%M"))
        self.cad.anotar_respostas({"b@aurora.example": agora - timedelta(days=400)})
        self.assertNotIn("b@aurora.example", self.cad.respondentes())
        self.assertIn("b@aurora.example", self.cad.respondentes(meses=24))
        self.cad.fixar_principal("fin@clinica.example", True)
        self.assertEqual(self.cad.principais(), {"fin@clinica.example"})
        self.cad.tirar("fin@clinica.example", "x")
        self.assertEqual(self.cad.principais(), set())

    def test_migrar_da_versao_anterior(self):
        self.cad.migrar({"velho@x.example": "desmarcado (01/09/2026)"}, {"forcado@x.example"})
        self.assertTrue(self.cad["velho@x.example"].fora)
        self.assertTrue(self.cad["forcado@x.example"].forcar_envio)
        dados = json.loads((self.pasta / "cadastro.json").read_text(encoding="utf-8"))
        self.assertIn("velho@x.example", dados["contatos"])

    def test_pedido_de_troca(self):
        self.cad.atualizar(EMPRESAS, {}, "p.xlsx")
        texto = "Solicito a troca do e-mail a@aurora.example para financeiro@aurora.example."
        self.assertEqual(M.pedido_de_troca("a@aurora.example", texto, self.cad), ("a@aurora.example", "financeiro@aurora.example"))
        # só o novo no texto: o remetente conhecido é o antigo
        self.assertEqual(M.pedido_de_troca("a@aurora.example", "Favor usar financeiro@aurora.example daqui pra frente.", self.cad),
                         ("a@aurora.example", "financeiro@aurora.example"))
        # e-mail da equipe citado no texto não conta
        self.assertEqual(M.pedido_de_troca("a@aurora.example", "Mandem para ana@contabil.example e novo@aurora.example", self.cad,
                                           "contabil.example"), ("a@aurora.example", "novo@aurora.example"))
        self.assertIsNone(M.pedido_de_troca("desconhecido@x.example", "troquem o e-mail", self.cad))
        self.assertTrue(M.email_valido("a.b+c@dominio.com.br"))
        self.assertFalse(M.email_valido("a@b"))


if __name__ == "__main__":
    unittest.main()
