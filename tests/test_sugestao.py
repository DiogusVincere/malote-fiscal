import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from malote import sugestao as S


class Categoria(unittest.TestCase):
    def test_categorias(self):
        self.assertEqual(S.categoria("Não somos mais clientes, favor retirar o e-mail.", "Ler"), "Sair da lista")
        self.assertEqual(S.categoria("Solicito a troca do e-mail para x@y.example", "Ler"), "Trocar e-mail")
        self.assertEqual(S.categoria("segue", "Mandou anexo"), "Mandou anexo")
        self.assertEqual(S.categoria("segue", "Mandou link"), "Mandou link")
        self.assertEqual(S.categoria("Pedágio entra como serviço tomado?", "Ler"), "Pergunta")
        self.assertEqual(S.categoria("Posso mandar até sexta", "Ler"), "Pergunta")
        self.assertEqual(S.categoria("Tudo bem? Não contratamos nada.", "Sem serviços"), "Sem serviços")
        self.assertEqual(S.categoria("Ok, obrigado.", "Ler"), "Outro")
        self.assertEqual(S.categoria("Enviado do meu iPhone https://x?y=1", "Ler"), "Outro")


class Saudacao(unittest.TestCase):
    def test_saudacao_e_nome(self):
        self.assertEqual(S.saudacao("Maria Souza", datetime(2026, 1, 1, 9)), "Bom dia, Maria.")
        self.assertEqual(S.saudacao("Financeiro | Empresa", datetime(2026, 1, 1, 15)), "Boa tarde,")
        self.assertEqual(S.saudacao("ACME EMBALAGENS", datetime(2026, 1, 1, 19)), "Boa noite,")
        self.assertEqual(S.primeiro_nome("Joaquina Ferreira"), "Joaquina")
        self.assertEqual(S.primeiro_nome("João da Silva"), "João")
        self.assertEqual(S.primeiro_nome("'Ana'"), "Ana")
        self.assertEqual(S.primeiro_nome("Contato Empresa"), "")
        self.assertEqual(S.primeiro_nome("maria@x.example"), "")
        self.assertEqual(S.primeiro_nome("Jo Silva"), "")


class BaseDeExemplos(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = S.Base(Path(self.tmp.name))

    def tearDown(self):
        self.tmp.cleanup()

    def test_regras_sem_exemplos(self):
        s = self.base.sugerir("segue a nota", "Mandou anexo")
        self.assertEqual((s.categoria, s.resposta), ("Mandou anexo", S.RESPOSTAS_PADRAO["Mandou anexo"]))
        self.assertTrue(s.porque.startswith("Regra:"))
        s = self.base.sugerir("Pedágio entra?", "Ler")
        self.assertEqual((s.categoria, s.resposta), ("Pergunta", ""))
        self.assertIn("ainda não há exemplo", s.porque)
        self.assertEqual(self.base.sugerir("Favor retirar o e-mail", "Ler").acoes, ["Tirar o e-mail da lista de envio"])

    def test_aprende_e_reconhece_parecida(self):
        self.base.ensinar("Despesas com pedágio e gasolina entram como serviço tomado?",
                          "Pedágio e combustível não entram.", "Pergunta", "ensinado")
        self.base.ensinar("Não houve serviço no mês.", "Ok, agradeço o retorno.", "Sem serviços")
        s = self.base.sugerir("pedágio entra como serviço?", "Ler")
        self.assertEqual(s.resposta, "Pedágio e combustível não entram.")
        self.assertGreaterEqual(s.parecenca, S.Base.LIMIAR)
        self.assertIn("Parecida", s.porque)
        # a mesma mensagem ensinada de novo troca a resposta (não duplica)
        self.base.ensinar("Não houve serviço no mês.", "Certo, obrigado!", "Sem serviços")
        self.assertEqual(len(self.base.exemplos), 2)
        self.assertEqual(S.Base(Path(self.tmp.name)).exemplos[1].resposta, "Certo, obrigado!")
        # pouco parecida com uma pergunta: mostra como ponto de partida, avisando
        s = self.base.sugerir("Estacionamento é serviço?", "Ler")
        self.assertEqual(s.categoria, "Pergunta")
        self.assertIn("Confira e ajuste", s.porque) if s.exemplo else self.assertEqual(s.resposta, "")
        self.base.esquecer(self.base.exemplos[0])
        self.assertEqual(len(self.base.exemplos), 1)
        self.assertEqual(self.base.parecidos(""), [])

    def test_montar_e_desmontar(self):
        self.assertEqual(S.montar_texto("Bom dia, Ana.", "Ok."), "Bom dia, Ana.\nEspero que você esteja bem.\nOk.\nAtenciosamente,")
        self.assertEqual(S.montar_texto("", "Ok."), "Espero que você esteja bem.\nOk.\nAtenciosamente,")
        antiga = "Bom dia, Maria.\n\nEspero que você esteja bem.\nOk, recebido.\nObrigado.\nAtenciosamente,\nEquipe"
        self.assertEqual(S.miolo_de_resposta_antiga(antiga), "Ok, recebido.\nObrigado.")


if __name__ == "__main__":
    unittest.main()
