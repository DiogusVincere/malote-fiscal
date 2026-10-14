import tempfile
import unittest
from pathlib import Path

from malote import clientes as K
from malote.planilha import Empresa

PASTAS = ["101 - PADARIA AURORA LTDA", "102 - CLINICA VITALIS SERVICOS MEDICOS LTDA", "103 - ACME USINAGEM LTDA",
          "1001 - BYTE CERTO INFORMATICA", "1002 - BYTE CERTO INFORMATICA", "52 - IRMAOS LOG LTDA",
          "052 - RECANTO NORTE LTDA", "Sem codigo"]


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.raiz = Path(self.tmp.name)
        for nome in PASTAS:
            (self.raiz / nome).mkdir()
        (self.raiz / "arquivo solto.txt").write_text("x")
        self.clientes = K.indexar(self.raiz)

    def tearDown(self):
        self.tmp.cleanup()


class Indexar(Base):
    def test_indexar(self):
        self.assertEqual(len(self.clientes), len(PASTAS))
        c = next(x for x in self.clientes if x.pasta.startswith("102"))
        self.assertEqual((c.codigo, c.nome, c.chave), ("102", "CLINICA VITALIS SERVICOS MEDICOS LTDA",
                                                        "CLINICA VITALIS SERVICOS MEDICOS"))
        self.assertEqual(next(x for x in self.clientes if x.pasta == "Sem codigo").codigo, None)
        self.assertEqual(K.indexar(self.raiz / "nao existe"), [])

    def test_normalizar(self):
        self.assertEqual(K.normalizar("J&J Sofá Center Ltda - ME"), "J J SOFA CENTER")
        self.assertEqual(K.normalizar(""), "")


class Identificar(Base):
    def ident(self, aprendidos=None, nomes=None, planilha=None):
        return K.Identificador(self.clientes, aprendidos or {}, nomes or {}, planilha or {})

    def test_planilha_manda(self):
        e101 = Empresa("101", "PADARIA AURORA LTDA", "1")
        e102 = Empresa("102", "CLINICA VITALIS", "2")
        ident = self.ident(planilha={"x@y.example": [e101], "dois@y.example": [e101, e102]})
        p = ident.identificar("X@Y.example")
        self.assertEqual((p.pasta, p.como, p.certo), ("101 - PADARIA AURORA LTDA", "planilha", True))
        p = ident.identificar("dois@y.example")
        self.assertEqual((p.pasta, p.como, p.certo), (None, "várias empresas", False))
        self.assertEqual(p.candidatos, ["101 - PADARIA AURORA LTDA", "102 - CLINICA VITALIS SERVICOS MEDICOS LTDA"])

    def test_pasta_da_empresa(self):
        ident = self.ident()
        self.assertEqual(ident.pasta_da_empresa(Empresa("0101", "qualquer", "")), "101 - PADARIA AURORA LTDA")
        # código repetido na rede: decide pelo nome
        self.assertEqual(ident.pasta_da_empresa(Empresa("52", "RECANTO NORTE LTDA", "")), "052 - RECANTO NORTE LTDA")
        self.assertEqual(ident.pasta_da_empresa(Empresa("52", "IRMAOS LOG", "")), "52 - IRMAOS LOG LTDA")
        # sem pasta com o código: pelo nome; sem nada parecido: None
        self.assertEqual(ident.pasta_da_empresa(Empresa("999", "ACME USINAGEM LTDA", "")), "103 - ACME USINAGEM LTDA")
        self.assertIsNone(ident.pasta_da_empresa(Empresa("999", "TOTALMENTE DIFERENTE", "")))
        self.assertIsNone(ident.pasta_da_empresa(Empresa("APELIDO", "", "")))

    def test_confirmado_e_dominio(self):
        ident = self.ident(aprendidos={"Quem@Acme.example": "103 - ACME USINAGEM LTDA"})
        p = ident.identificar("quem@acme.example")
        self.assertEqual((p.pasta, p.como, p.certo), ("103 - ACME USINAGEM LTDA", "confirmado", True))
        p = ident.identificar("outro@acme.example")
        self.assertEqual((p.pasta, p.como, p.certo), ("103 - ACME USINAGEM LTDA", "mesmo domínio", False))
        # domínio que casa com o nome do cliente
        p = ident.identificar("fin@acmeusinagem.example")
        self.assertEqual((p.pasta, p.como), ("103 - ACME USINAGEM LTDA", "domínio"))
        # provedor de e-mail pessoal não diz nada
        p = self.ident(aprendidos={"a@gmail.com": "103 - ACME USINAGEM LTDA"}).identificar("b@gmail.com")
        self.assertIsNone(p.pasta)

    def test_pelo_nome(self):
        ident = self.ident(nomes={"c@x.example": "Clinica Vitalis Servicos Medicos"})
        p = ident.identificar("c@x.example")
        self.assertEqual((p.pasta, p.como), ("102 - CLINICA VITALIS SERVICOS MEDICOS LTDA", "nome do contato"))
        p = ident.identificar("d@x.example", "Padaria Aurora (d@x.example)")
        self.assertEqual((p.pasta, p.como), ("101 - PADARIA AURORA LTDA", "nome no e-mail"))
        # duas pastas com o mesmo nome: ambíguo, não chuta
        self.assertIsNone(ident.identificar("e@x.example", "Byte Certo Informatica").pasta)
        self.assertIsNone(ident.identificar("f@x.example", "Zé").pasta)
        self.assertEqual(ident.identificar("", ""), K.Palpite(None, "", 0))


class Destino(Base):
    def test_reaproveita_grafias_e_cria_o_que_falta(self):
        cliente = "102 - CLINICA VITALIS SERVICOS MEDICOS LTDA"
        (self.raiz / cliente / "NFe" / "2026" / "09.2026" / "Serviços tomados").mkdir(parents=True)
        d = K.destino(self.raiz, cliente, (2026, 9))
        self.assertEqual(d.pasta, self.raiz / cliente / "NFe" / "2026" / "09.2026" / "Serviços tomados" / "Tomados e-mail")
        self.assertEqual(d.criar, ["Tomados e-mail"])
        self.assertIn("(cria: Tomados e-mail)", d.descricao)
        d = K.destino(self.raiz, cliente, (2026, 8))
        self.assertEqual(d.criar, ["08-2026", "Tomados", "Tomados e-mail"])
        self.assertEqual(d.pasta, self.raiz / cliente / "NFe" / "2026" / "08-2026" / "Tomados" / "Tomados e-mail")
        self.assertFalse(d.pasta.exists())
        self.assertTrue(K.garantir(d).is_dir())
        self.assertEqual(K.destino(self.raiz, cliente, (2026, 8)).criar, [])

    def test_prefere_tomados_exato(self):
        cliente = "101 - PADARIA AURORA LTDA"
        base = self.raiz / cliente / "NF-E" / "2026" / "092026"
        (base / "NFSE Tomados").mkdir(parents=True)
        (base / "Tomados").mkdir()
        self.assertEqual(K.destino(self.raiz, cliente, (2026, 9)).pasta, base / "Tomados" / "Tomados e-mail")

    def test_cliente_sem_pasta(self):
        with self.assertRaises(FileNotFoundError):
            K.destino(self.raiz, "999 - NAO EXISTE", (2026, 9))


if __name__ == "__main__":
    unittest.main()
