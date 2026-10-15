# -*- coding: utf-8 -*-
"""A rotina de terça deve transformar os relatórios do Drive exatamente na aba da semana da demonstração."""

import sys
import unittest
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(RAIZ), str(RAIZ / "dados")]

import dados_demo  # noqa: E402
import etl  # noqa: E402


def esperado(l: dict) -> list:
    """A linha que a aba da semana deve ter (17 colunas)."""
    return ["Despesa", l["desc"], l["valor"], etl._como_data(l["venc"]), "10/2026", etl._como_data(l["pgto"]), "Pago",
            l["tipodoc"] or None, int(l["nf"]) if l["nf"] else None, l["cat"], l["conta"], l["forma"] or None,
            l["forn"], None, None, l["centro"], l["placa"] or None]


class TestRotinaDeTerca(unittest.TestCase):
    def test_semana_passada(self):
        self.assertEqual(etl.semana_passada(date(2026, 10, 13)), (date(2026, 10, 5), date(2026, 10, 9)))
        self.assertEqual(etl.semana_passada(date(2026, 10, 19)), (date(2026, 10, 12), date(2026, 10, 16)))

    def test_transforma_os_relatorios_na_aba_da_semana(self):
        relatorios = [dados_demo.linhas_brutas(r) for r in dados_demo.RELATORIOS]
        ini, fim = etl.semana_passada(dados_demo.HOJE_ROTINA)
        saida = etl.transformar(relatorios, ini, fim, dados_demo.CENTROS)
        self.assertEqual(len(saida), len(dados_demo.LINHAS))
        for obtida, l in zip(saida, dados_demo.LINHAS):
            self.assertEqual(obtida, esperado(l))

    def test_movimentos_fora_da_semana_sao_ignorados(self):
        relatorios = [dados_demo.linhas_brutas(r) for r in dados_demo.RELATORIOS]
        total = sum(len(r) for r in relatorios)
        ini, fim = etl.semana_passada(dados_demo.HOJE_ROTINA)
        dentro = len(etl.transformar(relatorios, ini, fim, dados_demo.CENTROS))
        self.assertEqual(total - dentro, len(dados_demo.FORA_DA_SEMANA))

    def test_regras_individuais(self):
        self.assertEqual(etl._primeira_maiuscula("BOLETO"), "Boleto")
        self.assertEqual(etl._numero_nf("NF 1204"), "1204")
        self.assertEqual(etl._numero_nf("NFS-99"), "99")
        self.assertIsNone(etl._numero_nf("recibo"))
        linha = {"ENTRADA": None, "SAIDA": 10.5, "DATA": date(2026, 10, 7)}
        self.assertEqual(etl.aplicar_regra("EXTRAIR_MES", "DATA", linha, {}), "10/2026")
        self.assertEqual(etl.aplicar_regra("VALOR_DUPLO", ("ENTRADA", "SAIDA"), linha, {}), 10.5)
        self.assertEqual(etl.aplicar_regra("CONDICIONAL_DUPLO", ("ENTRADA", "SAIDA", "Receita", "Despesa"), linha, {}), "Despesa")


if __name__ == "__main__":
    unittest.main()
