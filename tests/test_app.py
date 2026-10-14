"""A tela só é importada onde há Tkinter; aqui se confere o que não precisa de janela."""
import importlib.util
import unittest


@unittest.skipUnless(importlib.util.find_spec("tkinter"), "sem tkinter")
class App(unittest.TestCase):
    def test_argumentos_e_pasta_de_dados(self):
        from malote import app
        args = app.argumentos(["--demo", "--dados", "x"])
        self.assertTrue(args.demo)
        self.assertEqual(args.dados, "x")
        self.assertFalse(app.argumentos([]).demo)
        self.assertEqual(app.pasta_de_dados_padrao().name, "Malote Fiscal")
        self.assertEqual(app.GRUPO_PADRAO.pattern, "^clientes")
        self.assertTrue(app.GRUPO_PADRAO.match("Clientes Fiscal 1"))
        self.assertIsNone(app.GRUPO_PADRAO.match("Fornecedores"))


if __name__ == "__main__":
    unittest.main()
