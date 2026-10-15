# -*- coding: utf-8 -*-
"""Cada cenário da planilha de demonstração deve dar o resultado e a ação esperados.

Não usa Excel nem navegador: reproduz, em memória, o que o robô faz (busca em lotes exata,
busca "contém" para o que o lote não achou) contra as mesmas regras do ERP fictício.

    python -m unittest discover -s testes -v
"""

import sys
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(RAIZ), str(RAIZ / "dados"), str(RAIZ / "erp_ficticio")]

import criterio  # noqa: E402
import dados_demo  # noqa: E402
import modelo  # noqa: E402
import normalizacao  # noqa: E402
from collections import Counter  # noqa: E402


def linhas_da_planilha() -> list:
    """O mesmo que planilha.ler_linhas_nf devolve: só as linhas com Nota Fiscal."""
    saida = []
    for i, l in enumerate(dados_demo.LINHAS, 2):
        if l["tipodoc"] != "Nota Fiscal":
            continue
        saida.append({
            "linha": i, "numero": l["nf"], "valor": l["valor"], "vencimento": l["venc"],
            "fornecedor": l["forn"], "descricao": l["desc"], "data_pagamento": l["pgto"],
            "conta": l["conta"], "forma": l["forma"], "categoria": l["cat"],
            "centro_custo": l["centro"], "placa": l["placa"],
        })
    return saida


def consultar_como_o_robo(linhas: list, estado: list, lote: int = 10) -> dict:
    """Lotes de 10 (busca exata) + busca individual ("contém") para o que o lote não achou."""
    exigir = Counter(l["numero"] for l in linhas)
    numeros = sorted(exigir, key=lambda n: (len(n), n))
    resultados = {}
    for k in range(0, len(numeros), lote):
        grupo = numeros[k:k + lote]
        resposta = modelo.buscar(estado, ", ".join(grupo), pagina=1, por_pagina=100)
        lidas = [normalizacao.resumir(t) for t in resposta["data"]]
        for n in grupo:
            reg = normalizacao.registro_de_lote(n, lidas, 1)
            if reg["quantidade"] > 0 and reg["quantidade"] >= exigir[n]:
                resultados[n] = reg
            else:
                busca = modelo.buscar(estado, n, pagina=1, por_pagina=100)
                resultados[n] = normalizacao.registro_de_leitura(
                    n, [normalizacao.resumir(t) for t in busca["data"]], busca["total_count"], 1)
    return resultados


class TestCenarios(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.linhas = linhas_da_planilha()
        cls.estado = modelo.novo_estado()
        cls.consultas = consultar_como_o_robo(cls.linhas, cls.estado)
        cls.saida = criterio.classificar_linhas(cls.linhas, cls.consultas)
        cls.por_nf = {}
        for l in cls.linhas:
            cls.por_nf.setdefault(l["numero"], []).append(cls.saida[l["linha"]])

    def test_todos_os_cenarios(self):
        for nf, (resultado, acao) in dados_demo.ESPERADO.items():
            with self.subTest(nota=nf):
                principal = self.por_nf[nf][0]
                self.assertEqual(principal["resultado"], resultado)
                self.assertEqual(principal["acao"], acao)

    def test_nenhum_cenario_ficou_sem_teste(self):
        self.assertEqual(set(self.por_nf), set(dados_demo.ESPERADO))

    def test_soma_de_duas_linhas_vira_uma_baixa_so(self):
        a, b = self.por_nf["7001"]
        self.assertEqual((a["acao"], a["valor_a_usar"]), ("DAR BAIXA", 1000.00))
        self.assertEqual(b["acao"], "IGNORAR")

    def test_diferenca_de_valor_entra_como_juros_ou_desconto(self):
        self.assertIn("JUROS", self.por_nf["587"][0]["observacao"])        # planilha R$ 910 x sistema R$ 850
        self.assertEqual(self.por_nf["587"][0]["valor_a_usar"], 910.00)    # o valor a usar é sempre o da planilha
        self.assertEqual(self.por_nf["4410"][0]["valor_a_usar"], 199.60)

    def test_numero_com_ano_e_zeros_so_aparece_na_busca_individual(self):
        self.assertEqual(self.consultas["88"]["via"], "individual")
        self.assertEqual(self.consultas["88"]["transacoes_iguais"], [])
        self.assertEqual(len(self.consultas["88"]["transacoes_parecidas"]), 1)

    def test_lancamento_pago_nunca_recebe_id_para_baixa(self):
        # o lançamento pago aparece só como referência: nada de "ID a usar" nem de DAR BAIXA
        for nf in ("774", "321"):
            r = self.por_nf[nf][0]
            self.assertFalse(r["preencher"])
            self.assertNotEqual(r["acao"], "DAR BAIXA")
            self.assertEqual(r["situacao"], "PAGO")

    def test_totais(self):
        cont = Counter(r["resultado"] for rs in self.por_nf.values() for r in rs)
        self.assertEqual(cont["JA_NO_SISTEMA"], 10)      # inclui as duas linhas da nota 7001
        self.assertEqual(cont["AJUSTAR_VALOR"], 1)
        self.assertEqual(cont["NAO_ENCONTRADA"], 3)
        self.assertEqual(cont["REVISAR"], 4)


class TestErp(unittest.TestCase):
    def setUp(self):
        self.estado = modelo.novo_estado()

    def pagamento(self, code, **extra):
        alvo = next(t for t in self.estado if t["code"] == code)
        base = {"code": code, "amount": alvo["remaining_amount"], "date": "2026-10-09",
                "account": "Conta Operações", "method": "Boleto"}
        base.update(extra)
        return modelo.pagar(self.estado, base)

    def test_pagamento_valido_quita_o_lancamento(self):
        r = self.pagamento(7101)
        self.assertTrue(r["ok"])
        t = next(t for t in self.estado if t["code"] == 7101)
        self.assertEqual((t["status"], t["payment_date"], t["remaining_amount"]), ("PAID", "2026-10-09", 0.0))

    def test_desconto_e_juros_mudam_o_liquido_nao_o_saldo(self):
        r = self.pagamento(7104, interest=60.00)         # saldo 850 + juros 60
        self.assertEqual(r["liquido"], 910.00)

    def test_nao_paga_duas_vezes(self):
        self.assertTrue(self.pagamento(7101)["ok"])
        self.assertFalse(self.pagamento(7101)["ok"])

    def test_nao_paga_lancamento_ja_pago(self):
        self.assertFalse(self.pagamento(7112)["ok"])

    def test_periodo_fechado_bloqueia(self):
        self.assertFalse(self.pagamento(7101, date="2026-08-15")["ok"])

    def test_so_aceita_pagamento_que_quita_o_saldo(self):
        self.assertFalse(self.pagamento(7101, amount=100.00)["ok"])

    def test_busca_exata_x_por_trecho(self):
        exata = modelo.buscar(self.estado, "88, 1204")["data"]
        self.assertEqual([t["document_number"] for t in exata], ["1204"])
        trecho = modelo.buscar(self.estado, "88")["data"]
        self.assertEqual([t["document_number"] for t in trecho], ["202600000088"])


class TestFornecedor(unittest.TestCase):
    def test_niveis(self):
        n = criterio.nivel_fornecedor
        self.assertEqual(n("Logística Rô", "Logística Rô Transportes Ltda"), "bate")
        self.assertEqual(n("Serviços Épsilon", "Épsilon Serviços Ltda"), "bate")
        self.assertEqual(n("Gama Peças", "Gama Autopeças Ltda"), "parecido")
        self.assertEqual(n("Elétrica Teta", "Hidráulica Ni Ltda"), "diferente")


if __name__ == "__main__":
    unittest.main()
