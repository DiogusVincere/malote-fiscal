import unittest
from datetime import date

from malote import competencia as C


class Competencia(unittest.TestCase):
    def test_anterior_proxima_ultimas(self):
        self.assertEqual(C.anterior(date(2026, 1, 15)), (2025, 12))
        self.assertEqual(C.anterior(date(2026, 10, 14)), (2026, 9))
        self.assertEqual(C.proxima((2026, 12)), (2027, 1))
        self.assertEqual(C.ultimas(3, date(2026, 10, 14)), [(2026, 9), (2026, 8), (2026, 7)])
        self.assertEqual(C.ultimas(2, date(2026, 1, 3)), [(2025, 12), (2025, 11)])

    def test_texto_e_assunto(self):
        self.assertEqual(C.texto((2026, 8)), "08/2026")
        self.assertEqual(C.do_texto(" 08/2026 "), (2026, 8))
        self.assertEqual(C.assunto((2026, 9)), "Serviços tomados - Setembro de 2026")
        self.assertEqual(C.assunto((2026, 3)), "Serviços tomados - Março de 2026")

    def test_do_assunto_so_no_padrao_exato(self):
        self.assertEqual(C.do_assunto("Serviços tomados - Setembro de 2026"), (2026, 9))
        self.assertEqual(C.do_assunto("servicos tomados - MARCO de 2026"), (2026, 3))
        self.assertIsNone(C.do_assunto("RES: Serviços tomados - Setembro de 2026"))
        self.assertIsNone(C.do_assunto("Serviços tomados - Mes de 2026"))
        self.assertIsNone(C.do_assunto(""))

    def test_mencionada(self):
        self.assertEqual(C.mencionada("RES: Serviços tomados - Agosto de 2026", (2026, 9)), (2026, 8))
        self.assertEqual(C.mencionada(".Setembro - Empresa Exemplo.zip", (2026, 9)), (2026, 9))
        self.assertEqual(C.mencionada("notas 08/2026.pdf", (2026, 9)), (2026, 8))
        self.assertEqual(C.mencionada("NF setembro-26.pdf", (2026, 9)), (2026, 9))
        # mês sem ano fica no ano que deixa a data mais perto da referência
        self.assertEqual(C.mencionada("Dezembro", (2026, 1)), (2025, 12))
        self.assertEqual(C.mencionada("Janeiro", (2026, 12)), (2027, 1))
        # gente não é mês; dois meses diferentes não decidem nada
        self.assertIsNone(C.mencionada("NF 265 MARCO TULIO.pdf", (2026, 3)))
        self.assertIsNone(C.mencionada("08/2026 e 09/2026", (2026, 9)))
        self.assertIsNone(C.mencionada("", (2026, 9)))

    def test_sem_acento(self):
        self.assertEqual(C.sem_acento("Serviços tomados — Março"), "Servicos tomados — Marco")


if __name__ == "__main__":
    unittest.main()
