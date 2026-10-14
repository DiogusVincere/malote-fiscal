import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from malote import planilha as P

CABECALHO = ["Código", "Razão Social", "CNPJ/CPF", "E-mails", "Regime municipal", "Resp. fiscal", "Ex-cliente"]
LINHAS = [
    ["101", "PADARIA AURORA LTDA", "12.345.678/0001-95", "a@aurora.example | *b@aurora.example", "Simples Nacional", "Ana", ""],
    ["102", "CLINICA VITALIS LTDA", "", "fin@vitalis.example; dra@vitalis.example", "Lucro Presumido", "Bruno", ""],
    ["103", "SEM EMAIL LTDA", "", "", "", "Carla", ""],
    ["104", "ISENTA LTDA", "", "ti@isenta.example", "ISENTA MUNICIPAL", "Carla", ""],
    ["105", "INTERNA LTDA", "", "x@interna.example | ana@contabil.example", "", "Ana", ""],
    ["000 EX-CLIENTE", "ANTIGA LTDA", "", "antiga@antiga.example", "", "Carla", ""],
    ["106", "MARCADA LTDA", "", "m@marcada.example", "", "Carla", "Sim"],
    [None, "linha sem código", "", "ignorada@x.example", "", "", ""],
]


def gravar(pasta: Path, nome="Clientes 2026-10-01.xlsx", cabecalho=CABECALHO, linhas=LINHAS) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.append(cabecalho)
    for linha in linhas:
        ws.append(linha)
    caminho = pasta / nome
    wb.save(caminho)
    return caminho


class Ler(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.pasta = Path(self.tmp.name)
        self.empresas = P.ler(gravar(self.pasta))

    def tearDown(self):
        self.tmp.cleanup()

    def test_colunas_e_valores(self):
        e = {x.codigo: x for x in self.empresas}
        self.assertEqual(len(self.empresas), 7)
        self.assertEqual(e["101"].cnpj, "12345678000195")
        self.assertEqual(e["101"].emails, ["a@aurora.example", "b@aurora.example"])
        self.assertEqual(e["102"].emails, ["fin@vitalis.example", "dra@vitalis.example"])
        self.assertEqual(e["101"].rotulo, "101 - PADARIA AURORA LTDA")
        self.assertTrue(e["101"].recebe and e["102"].recebe and e["103"].recebe and e["105"].recebe)
        self.assertFalse(e["104"].recebe)
        self.assertTrue(e["000 EX-CLIENTE"].ex_cliente and not e["000 EX-CLIENTE"].recebe)
        self.assertTrue(e["106"].ex_cliente)

    def test_coluna_faltando(self):
        caminho = gravar(self.pasta, "Clientes sem emails.xlsx", ["Código", "Razão Social", "Regime municipal"],
                         [["1", "X", ""]])
        with self.assertRaises(ValueError):
            P.ler(caminho)

    def test_emails_do_texto(self):
        self.assertEqual(P.emails_do_texto(" <A@X.example>, *b@y.example ; invalido@ ; b@y.example"),
                         ["a@x.example", "b@y.example"])

    def test_achar_a_mais_nova(self):
        gravar(self.pasta, "Clientes 2026-09-01.xlsx")
        gravar(self.pasta, "~$Clientes 2026-12-01.xlsx")
        gravar(self.pasta, "Outra coisa.xlsx")
        self.assertEqual(P.achar([self.pasta, None, self.pasta / "nao-existe"]).name, "Clientes 2026-10-01.xlsx")
        self.assertIsNone(P.achar([self.pasta], prefixo="Fornecedores"))

    def test_por_email(self):
        mapa = P.por_email(self.empresas)
        self.assertEqual([x.codigo for x in mapa["a@aurora.example"]], ["101"])
        self.assertNotIn("antiga@antiga.example", mapa)
        self.assertIn("antiga@antiga.example", P.por_email(self.empresas, so_ativas=False))


class ListaDeEnvio(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.empresas = P.ler(gravar(Path(self.tmp.name)))

    def tearDown(self):
        self.tmp.cleanup()

    def lista(self, **kw):
        kw.setdefault("dominio_interno", "contabil.example")
        return P.lista_de_envio(self.empresas, kw.pop("fora", {}), kw.pop("tamanho", 40), **kw)

    def test_quem_recebe(self):
        lista = self.lista()
        por = {d.email: d for d in lista.destinatarios}
        recebem = {d.email for d in lista.destinatarios if d.incluido}
        self.assertEqual(recebem, {"a@aurora.example", "b@aurora.example", "fin@vitalis.example", "dra@vitalis.example",
                                   "x@interna.example"})
        self.assertEqual(por["ana@contabil.example"].observacao, "e-mail interno")
        self.assertFalse(por["ana@contabil.example"].incluido)
        self.assertNotIn("ti@isenta.example", por)       # isenta: só entra se vier de um grupo
        self.assertNotIn("antiga@antiga.example", por)
        self.assertEqual([e.codigo for e in lista.sem_email], ["103"])
        self.assertEqual(lista.empresas, 4)
        self.assertEqual(lista.lotes, [sorted(recebem)])

    def test_grupos_somam_e_explicam(self):
        grupos = {"ti@isenta.example": ["Clientes 1"], "antiga@antiga.example": ["Clientes 1"],
                  "novo@grupo.example": ["Clientes 2"], "a@aurora.example": ["Clientes 1"]}
        por = {d.email: d for d in self.lista(grupos=grupos).destinatarios}
        self.assertEqual((por["ti@isenta.example"].incluido, por["ti@isenta.example"].observacao),
                         (True, "regime municipal: ISENTA MUNICIPAL"))
        self.assertEqual((por["antiga@antiga.example"].incluido, por["antiga@antiga.example"].observacao),
                         (False, "ex-cliente na planilha"))
        self.assertEqual((por["novo@grupo.example"].origem, por["novo@grupo.example"].observacao),
                         ("só grupo", "não está na planilha"))
        self.assertEqual(por["a@aurora.example"].origem, "planilha + grupo")
        self.assertEqual(por["b@aurora.example"].origem, "só planilha")
        self.assertEqual(por["b@aurora.example"].observacao, "não está nos grupos")
        # marcado à mão mesmo sendo ex-cliente
        por = {d.email: d for d in self.lista(grupos=grupos, incluir={"antiga@antiga.example"}).destinatarios}
        self.assertTrue(por["antiga@antiga.example"].incluido)

    def test_fora_da_lista_e_manuais(self):
        manuais = {"m@manual.example": [self.empresas[0]]}
        lista = self.lista(fora={"a@aurora.example": "pediu para sair (01/10/2026)"}, manuais=manuais)
        por = {d.email: d for d in lista.destinatarios}
        self.assertEqual((por["a@aurora.example"].incluido, por["a@aurora.example"].observacao),
                         (False, "pediu para sair (01/10/2026)"))
        self.assertEqual((por["m@manual.example"].origem, por["m@manual.example"].observacao, por["m@manual.example"].incluido),
                         ("manual", "acrescentado à mão", True))

    def test_quem_responde_e_principal(self):
        # a clínica tem dois e-mails; só o financeiro respondeu: a doutora é poupada
        por = {d.email: d for d in self.lista(respondentes={"fin@vitalis.example"}).destinatarios}
        self.assertFalse(por["dra@vitalis.example"].incluido)
        self.assertEqual(por["dra@vitalis.example"].observacao, "a empresa responde por fin@vitalis.example (102)")
        self.assertTrue(por["fin@vitalis.example"].incluido)
        # principal fixado manda mais que quem respondeu
        por = {d.email: d for d in self.lista(respondentes={"fin@vitalis.example"},
                                              principais={"dra@vitalis.example"}).destinatarios}
        self.assertTrue(por["dra@vitalis.example"].incluido)
        self.assertEqual(por["fin@vitalis.example"].observacao, "principal da empresa: dra@vitalis.example (102)")
        # sem histórico, recebem todos
        por = {d.email: d for d in self.lista().destinatarios}
        self.assertTrue(por["dra@vitalis.example"].incluido and por["fin@vitalis.example"].incluido)

    def test_email_de_varias_empresas_so_sai_se_sobrar_em_todas(self):
        compart = P.Empresa("107", "OUTRA LTDA", "", ["dra@vitalis.example"])
        lista = P.lista_de_envio(self.empresas + [compart], {}, 40, respondentes={"fin@vitalis.example"})
        por = {d.email: d for d in lista.destinatarios}
        self.assertTrue(por["dra@vitalis.example"].incluido)  # é o único e-mail da 107

    def test_lotes(self):
        lista = self.lista(tamanho=2)
        self.assertEqual([len(l) for l in lista.lotes], [2, 2, 1])


if __name__ == "__main__":
    unittest.main()
