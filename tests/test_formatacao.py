import unittest

from app import converter_decimal, formatar_moeda, formatar_numero


class TestFormatacao(unittest.TestCase):
    def test_converter_decimal(self):
        casos = [
            ("35,50", 35.5),
            ("35.50", 35.5),
            ("1.234,56", 1234.56),
            ("R$ 42,00", 42.0),
        ]
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(converter_decimal(entrada), esperado)

    def test_converter_decimal_vazio(self):
        with self.assertRaises(ValueError):
            converter_decimal("")

    def test_formatar_numero_no_padrao_brasileiro(self):
        self.assertEqual(formatar_numero(1234567.89), "1.234.567,89")

    def test_formatar_moeda_no_padrao_brasileiro(self):
        self.assertEqual(formatar_moeda(1234.5), "R$ 1.234,50")


if __name__ == "__main__":
    unittest.main()
