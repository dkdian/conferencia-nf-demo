# -*- coding: utf-8 -*-
"""Simula a pessoa decidindo as linhas em branco da aba Conferência NF (só para a demonstração).

Na vida real, quem faz isso é a pessoa, escolhendo na lista da coluna Ação. Aqui o script faz o mesmo
para as notas de DECISOES_HUMANAS (ver dados_demo.py), para a demonstração poder rodar sem ninguém.

    python dados/aplicar_decisoes.py
"""

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(RAIZ), str(RAIZ / "dados")]

import config as cfg  # noqa: E402
import criterio  # noqa: E402
import dados_demo  # noqa: E402
import planilha  # noqa: E402
from aplicar_na_planilha import colunas  # noqa: E402


def main() -> int:
    def aplicar():
        feitas = []
        with planilha.planilha_aberta(aba=cfg.ABA_RESULTADO, gravar=True) as ws:
            COL = colunas(ws)
            for r in range(cfg.LINHA_TABELA_RESULTADO + 1, planilha.ultima_linha(ws) + 1):
                nf = criterio.normalizar(ws.Cells(r, COL["nf"]).Value)
                acao = dados_demo.DECISOES_HUMANAS.get(nf)
                if acao and not ws.Cells(r, COL["acao"]).Value:
                    ws.Cells(r, COL["acao"]).Value = acao
                    feitas.append((nf, acao))
        return feitas
    feitas = planilha.com_tentativas(aplicar)
    for nf, acao in feitas:
        print("  NF %s -> %s" % (nf, acao))
    print("%d decisão(ões) aplicada(s)." % len(feitas))
    return 0


if __name__ == "__main__":
    sys.exit(planilha.executar(main))
