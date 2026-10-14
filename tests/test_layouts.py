import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from malote import layouts as L


class Campos(unittest.TestCase):
    def test_preencher(self):
        texto = "{saudacao}! Mês {mes} ({mes_nome} de {ano}), {nome}."
        self.assertEqual(L.preencher(texto, (2026, 9), datetime(2026, 10, 2, 8), "Ana"),
                         "Bom dia! Mês 09/2026 (Setembro de 2026), Ana.")
        self.assertEqual(L.periodo(datetime(2026, 1, 1, 14)), "Boa tarde")
        self.assertEqual(L.periodo(datetime(2026, 1, 1, 20)), "Boa noite")

    def test_texto_do_html_e_para_layout(self):
        html = ('<html><body lang=PT-BR><p>Boa tarde!</p><p>Sua empresa <b>tomou</b> serviço em 09/2026 '
                '(Setembro de 2026)?</p><ul><li>sim</li><li>n&atilde;o</li></ul><img src="cid:x"><p>Att,</p></body></html>')
        texto = L.texto_do_html(html)
        self.assertEqual(texto, "Boa tarde!\n\nSua empresa **tomou** serviço em 09/2026 (Setembro de 2026)?\n- sim\n- não\n\nAtt,")
        self.assertEqual(L.para_layout(texto, (2026, 9)),
                         "{saudacao}!\n\nSua empresa **tomou** serviço em {mes} ({mes_nome} de {ano})?\n- sim\n- não\n\nAtt,")

    def test_html(self):
        html = L.html("Olá <você>\nlinha 2\n\n- item **forte**\n- outro\n\nFim", '<img src="cid:ass">')
        self.assertIn("<p class=MsoNormal", html)
        self.assertIn("Olá &lt;você&gt;<br>linha 2", html)
        self.assertIn("<li", html)
        self.assertIn("<b>forte</b>", html)
        self.assertTrue(html.rstrip().endswith('<p class=MsoNormal><img src="cid:ass"></p></body></html>'))
        self.assertEqual(L.texto_do_html(L.html(L.TEXTO_PADRAO)).count("- "), 2)


class Arquivo(unittest.TestCase):
    def test_layouts_persistem(self):
        with tempfile.TemporaryDirectory() as tmp:
            lay = L.Layouts(Path(tmp) / "sub")
            lay.garantir_padrao("", (2026, 9), texto=L.TEXTO_PADRAO)
            self.assertEqual(lay.pergunta[L.PADRAO]["assunto"], L.ASSUNTO_PADRAO)
            self.assertEqual(lay.pergunta[L.PADRAO]["texto"], L.TEXTO_PADRAO)
            lay.salvar_pergunta("Curto", "Notas {mes}", "Mandem as notas de {mes}.")
            lay.salvar_resposta("Obrigado", "Obrigado, {nome}!")
            de_novo = L.Layouts(Path(tmp) / "sub")
            self.assertEqual(de_novo.pergunta_atual, "Curto")
            self.assertEqual(de_novo.resposta, {"Obrigado": "Obrigado, {nome}!"})
            de_novo.excluir_pergunta("Curto")
            self.assertEqual(de_novo.pergunta_atual, L.PADRAO)
            de_novo.excluir_resposta("Obrigado")
            self.assertEqual(de_novo.resposta, {})
            # do último envio: o texto vira layout com campos
            outro = L.Layouts(Path(tmp) / "outro")
            outro.garantir_padrao(L.html("Bom dia!\n\nTomou serviço em 09/2026?"), (2026, 9))
            self.assertEqual(outro.pergunta[L.PADRAO]["texto"], "{saudacao}!\n\nTomou serviço em {mes}?")


if __name__ == "__main__":
    unittest.main()
