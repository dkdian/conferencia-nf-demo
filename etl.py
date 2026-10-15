# -*- coding: utf-8 -*-
"""Regras de tratamento dos relatórios do Drive (sem Excel, sem rede): puro e testável.

O cliente deixa dois relatórios por semana. A rotina de terça lê os dois, fica só com a semana passada
(segunda a sexta) e transforma cada movimento numa linha do formato da aba de inserção (17 colunas).
Cada coluna de saída tem uma regra, como na tabela de configuração da planilha original:

  COPIAR              traz o valor como está (datas continuam datas)
  FIXO                o mesmo texto em todas as linhas
  VAZIO               cria a coluna e deixa em branco
  PRIMEIRA_MAIUSCULA  "BOLETO" -> "Boleto"
  EXTRAIR_MES         data -> "mm/aaaa"
  VALOR_DUPLO         usa a coluna que estiver preenchida (SAIDA ou ENTRADA)
  CONDICIONAL_DUPLO   "Despesa" se SAIDA tem valor, "Receita" se ENTRADA tem valor
  DE_PARA             troca o código pelo nome ("OPS-SUL" -> "Operações Sul")
  NUMERO_DOCUMENTO    acha "NF 1204" no campo NF/OS/FATURA: tipo "Nota Fiscal" + número 1204
"""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from pathlib import Path

CABECALHO = ["Tipo (Obrigatório)", "Descrição (Obrigatório)", "Valor (Obrigatório)",
             "Data de vencimento (Obrigatório)", "Competência", "Data de pagamento", "Status",
             "Tipo Documento", "Numero Documento", "Categoria", "Conta bancária",
             "Forma de pagamento", "Colaborador/Cliente", "Documento colaborador/cliente",
             "Observação", "Centro de Custo", "Placa"]

# (coluna de saída, regra, coluna de entrada ou parâmetro)
REGRAS = [
    ("Tipo (Obrigatório)", "CONDICIONAL_DUPLO", ("ENTRADA", "SAIDA", "Receita", "Despesa")),
    ("Descrição (Obrigatório)", "COPIAR", "DESCRICAO"),
    ("Valor (Obrigatório)", "VALOR_DUPLO", ("ENTRADA", "SAIDA")),
    ("Data de vencimento (Obrigatório)", "COPIAR", "VENCIMENTO"),
    ("Competência", "EXTRAIR_MES", "DATA"),
    ("Data de pagamento", "COPIAR", "DATA"),
    ("Status", "FIXO", "Pago"),
    ("Tipo Documento", "NUMERO_DOCUMENTO", "NF/OS/FATURA"),
    ("Numero Documento", "NUMERO_DOCUMENTO", "NF/OS/FATURA"),
    ("Categoria", "COPIAR", "CATEGORIA"),
    ("Conta bancária", "COPIAR", "CONTA"),
    ("Forma de pagamento", "PRIMEIRA_MAIUSCULA", "FORMA PGTO"),
    ("Colaborador/Cliente", "PRIMEIRA_MAIUSCULA", "CLIENTE/FORNECEDOR"),
    ("Documento colaborador/cliente", "VAZIO", None),
    ("Observação", "VAZIO", None),
    ("Centro de Custo", "DE_PARA", "CENTRO DE CUSTO"),
    ("Placa", "COPIAR", "PLACA"),
]

RE_NF = re.compile(r"\bNF[A-Za-z]?\s*[-:.]?\s*(\d+)", re.I)


def semana_passada(hoje: date) -> tuple:
    """(segunda, sexta) da semana anterior à de `hoje`."""
    segunda = hoje - timedelta(days=hoje.weekday() + 7)
    return segunda, segunda + timedelta(days=4)


def _como_data(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _primeira_maiuscula(v):
    return " ".join(p.capitalize() for p in str(v).split()) if v not in (None, "") else None


def _numero_nf(texto):
    m = RE_NF.search(str(texto or ""))
    return m.group(1) if m else None


def aplicar_regra(regra, arg, linha: dict, de_para: dict):
    if regra == "COPIAR":
        v = linha.get(arg)
        return _como_data(v) if arg in ("DATA", "VENCIMENTO") else (v if v not in ("", None) else None)
    if regra == "FIXO":
        return arg
    if regra == "VAZIO":
        return None
    if regra == "PRIMEIRA_MAIUSCULA":
        return _primeira_maiuscula(linha.get(arg))
    if regra == "EXTRAIR_MES":
        d = _como_data(linha.get(arg))
        return d.strftime("%m/%Y") if d else None
    if regra == "VALOR_DUPLO":
        a, b = arg
        return linha.get(a) if linha.get(a) not in (None, "", 0) else linha.get(b)
    if regra == "CONDICIONAL_DUPLO":
        a, b, ta, tb = arg
        return ta if linha.get(a) not in (None, "", 0) else (tb if linha.get(b) not in (None, "", 0) else None)
    if regra == "DE_PARA":
        v = linha.get(arg)
        return de_para.get(v, v)
    raise ValueError("regra desconhecida: %s" % regra)


def transformar_linha(linha: dict, de_para: dict) -> list:
    saida = []
    nf = _numero_nf(linha.get("NF/OS/FATURA"))
    for coluna, regra, arg in REGRAS:
        if regra == "NUMERO_DOCUMENTO":
            if nf is None:
                saida.append(None)
            elif coluna == "Tipo Documento":
                saida.append("Nota Fiscal")
            else:
                saida.append(int(nf))
        else:
            saida.append(aplicar_regra(regra, arg, linha, de_para))
    return saida


def transformar(relatorios: list, inicio: date, fim: date, de_para: dict) -> list:
    """relatorios = [linhas do relatório 1, linhas do relatório 2, ...]. Devolve as linhas da semana, por data."""
    todas = []
    for linhas in relatorios:
        for l in linhas:
            d = _como_data(l.get("DATA"))
            if d and inicio <= d <= fim:
                todas.append(transformar_linha(l, de_para))
    todas.sort(key=lambda r: r[5])      # por data de pagamento (estável: mantém a ordem dos relatórios)
    return todas


def ler_relatorio(caminho: Path) -> list:
    """Lê um relatório .xlsx exportado pelo cliente (aba "Movimentos") como lista de dicionários."""
    import openpyxl
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb["Movimentos"]
    linhas = list(ws.iter_rows(values_only=True))
    wb.close()
    cab = [str(c).strip() for c in linhas[0]]
    return [dict(zip(cab, l)) for l in linhas[1:] if any(v not in (None, "") for v in l)]
