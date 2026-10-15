# -*- coding: utf-8 -*-
"""
Rotina de terça-feira (na vida real, roda sozinha às 10h pelo Agendador de Tarefas do Windows).

  1. Baixa os dois relatórios de fluxo de caixa que o cliente deixa no Drive e os arquiva numa pasta
     com a data (recebidos/<Mês>/<dd.mm>) e na pasta "Atual".
  2. Aplica as regras de tratamento (etl.py): só a semana passada, segunda a sexta.
  3. Cria na planilha a aba nova "Inserção - dd.mm", sem alterar nenhuma aba existente.
  4. Avisa o gestor no chat para validar as informações. Depois do OK dele, a pessoa segue com os
     botões 1, 2 e 3.

Na demonstração o Drive é a pasta drive_ficticio/ e o chat é uma página do ERP fictício.

    python rotina_terca.py [--hoje dd/mm/aaaa] [--sem-chat]
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import urllib.request
from datetime import datetime

import config as cfg
import etl
import planilha
from erp_ficticio import garantir

sys.path.insert(0, str(cfg.BASE / "dados"))
import dados_demo  # noqa: E402

DRIVE = cfg.BASE / "drive_ficticio"
RECEBIDOS = cfg.BASE / "recebidos"
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
         "Outubro", "Novembro", "Dezembro"]
LARGURAS = [9, 38, 13, 16, 12, 16, 8, 14, 16, 16, 18, 16, 24, 16, 12, 18, 12]


def nome_relatorio(relatorio: str) -> str:
    return "Fluxo de caixa - %s.xlsx" % relatorio


def formatar_aba(ws, linhas: int):
    """Cabeçalho em destaque, larguras e formato de valor (só apresentação)."""
    cab = ws.Range(ws.Cells(1, 1), ws.Cells(1, cfg.ULTIMA_COL_INSERCAO))
    cab.Font.Bold = True
    cab.Font.Color = 0xFFFFFF
    cab.Interior.Color = 0x965430          # azul (BGR)
    cab.WrapText = True
    cab.VerticalAlignment = -4108
    ws.Rows(1).RowHeight = 48
    for i, w in enumerate(LARGURAS, 1):
        ws.Columns(i).ColumnWidth = w
    ws.Range(ws.Cells(2, cfg.COL_VALOR), ws.Cells(linhas + 1, cfg.COL_VALOR)).NumberFormat = planilha.formatos(ws)["numero"]


def avisar_gestor(hoje, ini, fim, total, por_relatorio: dict, aba: str, pasta: str):
    fatos = [["Período", "%s a %s" % (ini.strftime("%d/%m"), fim.strftime("%d/%m/%Y"))],
             ["Lançamentos", str(total)]]
    fatos += [[r, str(n)] for r, n in por_relatorio.items()]
    fatos += [["Pasta do dia", pasta], ["Aba criada na planilha", aba]]
    card = {"autor": "Robô de rotinas", "papel": "robo", "tipo": "card", "hora": "10:00",
            "titulo": "Fluxo de caixa semanal: validar antes de lançar", "fatos": fatos,
            "texto": "A aba foi criada na planilha. Valide as informações e responda OK para seguir com a "
                     "conferência das notas no ERP."}
    req = urllib.request.Request(cfg.URL_ERP + "/api/chat", data=json.dumps(card).encode("utf-8"), method="POST",
                                 headers={"Content-Type": "application/json"})
    urllib.request.urlopen(req, timeout=10).read()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--hoje", help="dia em que a rotina roda, dd/mm/aaaa (padrão da demonstração: 13/10/2026)")
    ap.add_argument("--sem-chat", action="store_true", help="não avisa o gestor")
    args = ap.parse_args()
    hoje = (datetime.strptime(args.hoje or os.environ.get("DEMO_HOJE") or "", "%d/%m/%Y").date()
            if (args.hoje or os.environ.get("DEMO_HOJE")) else dados_demo.HOJE_ROTINA)
    garantir.garantir_servidor()
    ini, fim = etl.semana_passada(hoje)

    print("=" * 70)
    print("  Rotina de terça -- demonstração")
    print("=" * 70)
    print("\n  Hoje: %s, %s" % (["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"][hoje.weekday()],
                              hoje.strftime("%d/%m/%Y")))
    print("  Período analisado: %s a %s (segunda a sexta da semana passada)" % (
        ini.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")))

    # 1) Drive -> pasta do dia e pasta Atual
    print("\n  [1/4] Baixando os relatórios do Drive...")
    pasta_dia = RECEBIDOS / MESES[hoje.month - 1] / hoje.strftime("%d.%m")
    atual = RECEBIDOS / "Atual"
    for pasta in (pasta_dia, atual):
        pasta.mkdir(parents=True, exist_ok=True)
    relatorios = {}
    for rel in dados_demo.RELATORIOS:
        origem = DRIVE / nome_relatorio(rel)
        if not origem.exists():
            print("  ! Não achei %s no Drive. Parei." % origem.name)
            return 1
        for pasta in (pasta_dia, atual):
            shutil.copy2(origem, pasta / origem.name)
        relatorios[rel] = etl.ler_relatorio(pasta_dia / origem.name)
        print("        - %s (%d movimentos)" % (origem.name, len(relatorios[rel])))
    print("        Salvos em %s e em %s" % (pasta_dia.relative_to(cfg.BASE), atual.relative_to(cfg.BASE)))

    # 2) tratamento
    print("\n  [2/4] Aplicando as regras de tratamento...")
    por_relatorio = {rel: len(etl.transformar([linhas], ini, fim, dados_demo.CENTROS)) for rel, linhas in relatorios.items()}
    linhas = etl.transformar(list(relatorios.values()), ini, fim, dados_demo.CENTROS)
    print("        " + " · ".join("%s: %d lançamentos" % (r, n) for r, n in por_relatorio.items()))
    ignorados = sum(len(l) for l in relatorios.values()) - len(linhas)
    print("        %d movimentos de outras semanas ficaram de fora" % ignorados)
    if not linhas:
        print("  Nenhum lançamento na semana. Nada a fazer.")
        return 1

    # 3) aba nova
    print("\n  [3/4] Criando a aba nova na planilha...")

    def criar():
        with planilha.com_workbook(gravar=True) as wb:
            ws = planilha.criar_aba_insercao(wb, etl.CABECALHO, linhas, hoje=datetime(hoje.year, hoje.month, hoje.day),
                                             cadastro=False)
            formatar_aba(ws, len(linhas))
            return ws.Name, planilha.dias_repetidos(wb, ws.Name)
    aba, repetidos = planilha.com_tentativas(criar)
    print("        Aba criada: %s (%d linhas). Nenhuma aba existente foi alterada." % (aba, len(linhas)))
    if repetidos:
        print("        ATENÇÃO, dias que também estão em outra aba:")
        print(planilha.texto_dias_repetidos(repetidos))
    if planilha.ja_esta_aberta():
        print("        A planilha está aberta: a aba foi criada na sua janela, SEM salvar. Salve com Ctrl+S.")

    # 4) aviso ao gestor
    if args.sem_chat:
        print("\n  [4/4] Aviso ao gestor não enviado (--sem-chat).")
    else:
        print("\n  [4/4] Avisando o gestor no chat para validar...")
        avisar_gestor(hoje, ini, fim, len(linhas), por_relatorio, aba, pasta_dia.name)
        print("        Aviso enviado ao canal #validação-semanal.")
    return 0


if __name__ == "__main__":
    sys.exit(planilha.executar(main))
