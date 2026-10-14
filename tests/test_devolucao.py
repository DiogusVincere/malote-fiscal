import unittest

from malote import devolucao as D

POSTFIX = """This is the mail system at host mx.contabil.example.

<fiscal@empresa.example>: host mx.empresa.example[10.0.0.1] said: 550 5.1.1
    <fiscal@empresa.example>: Recipient address rejected: User unknown (in reply to RCPT TO command)
"""
RELATORIO = """Delivery report

Final-Recipient: rfc822;contato@empresa.example
Action: failed
Status: 5.2.2 (mailbox full)
"""
OUTLOOK = """Não é possível entregar: Serviços tomados - Setembro de 2026

Sua mensagem não foi entregue a estes destinatários: pessoa@cliente.example

O servidor recusou a mensagem: 550 5.7.26 Unauthenticated email (SPF/DKIM)

Cabeçalhos de mensagem originais
Return-Path: fiscal@contabil.example
To: outra@cliente.example
"""


class Devolucao(unittest.TestCase):
    def test_e_devolucao(self):
        self.assertTrue(D.e_devolucao("", "x", "REPORT.IPM.Note.NDR"))
        self.assertTrue(D.e_devolucao("postmaster@x.example", "x", "IPM.Note"))
        self.assertTrue(D.e_devolucao("Mail Delivery Subsystem", "x", "IPM.Note"))
        self.assertTrue(D.e_devolucao("a@b.example", "Undeliverable: oi", "IPM.Note"))
        self.assertTrue(D.e_devolucao("a@b.example", "Não é possível entregar: oi", "IPM.Note"))
        self.assertFalse(D.e_devolucao("a@b.example", "RES: Serviços tomados", "IPM.Note"))

    def test_motivos(self):
        self.assertEqual(D.motivo("550 5.1.1 User unknown"), ("endereço ou domínio não existe", False))
        self.assertEqual(D.motivo("unable to route: dns lookup failure"), ("endereço ou domínio não existe", False))
        self.assertEqual(D.motivo("552 5.2.2 Over quota"), ("caixa do cliente cheia", True))
        self.assertEqual(D.motivo("550 5.7.26 unauthenticated"), ("recusado pelo servidor do cliente (spam/autenticação)", True))
        self.assertEqual(D.motivo("454 4.7.1 relay access denied"), ("falha temporária", True))
        self.assertEqual(D.motivo("will not continue trying"), ("falha definitiva (o servidor desistiu)", False))
        self.assertEqual(D.motivo("???"), ("não entregue (motivo não reconhecido)", True))

    def test_ler_formatos(self):
        ignorar = {"fiscal@contabil.example"}
        f = D.ler(POSTFIX, ignorar)
        self.assertEqual([(x.email, x.reenviar) for x in f], [("fiscal@empresa.example", False)])
        f = D.ler(RELATORIO, ignorar)
        self.assertEqual([(x.email, x.motivo) for x in f], [("contato@empresa.example", "caixa do cliente cheia")])
        f = D.ler(OUTLOOK, ignorar)
        # só o que vem antes dos cabeçalhos originais; o próprio remetente não conta
        self.assertEqual([x.email for x in f], ["pessoa@cliente.example"])
        self.assertTrue(f[0].reenviar)
        self.assertEqual(D.ler("", ignorar), [])


if __name__ == "__main__":
    unittest.main()
