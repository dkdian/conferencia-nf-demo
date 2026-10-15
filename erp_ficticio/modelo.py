# -*- coding: utf-8 -*-
"""Regras do ERP de mentira: busca de lançamentos e registro de pagamento.

Fica separado do servidor para o teste poder usar as mesmas regras sem abrir porta nenhuma.
Comportamento imitado de um ERP financeiro comum:
  * filtro "Número do documento" com um número só: busca por trecho ("contém");
  * vários números separados por vírgula: busca exata de cada um;
  * período e itens por página como filtros da listagem;
  * pagamento só quitando o saldo, em data de período aberto, com conta e forma de pagamento.
"""

from __future__ import annotations

import copy
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "dados"))
import dados_demo  # noqa: E402

PERIODO_FECHADO_ATE = "2026-08-31"      # datas até aqui estão em período contábil fechado


def novo_estado() -> list:
    return copy.deepcopy(dados_demo.transacoes_erp())


def buscar(transacoes: list, numero_documento: str = "", data_inicio=None, data_fim=None,
           pagina: int = 1, por_pagina: int = 10) -> dict:
    texto = (numero_documento or "").strip().lower()
    if "," in texto:
        exatos = {t.strip() for t in texto.split(",") if t.strip()}
        achadas = [t for t in transacoes if t["document_number"].lower() in exatos]
    elif texto:
        achadas = [t for t in transacoes if texto in t["document_number"].lower()]
    else:
        achadas = list(transacoes)
    if data_inicio:
        achadas = [t for t in achadas if t["due_date"] >= data_inicio]
    if data_fim:
        achadas = [t for t in achadas if t["due_date"] <= data_fim]
    achadas.sort(key=lambda t: (t["due_date"], t["code"]), reverse=True)
    total = len(achadas)
    por_pagina = max(1, min(int(por_pagina or 10), 100))
    inicio = (max(int(pagina or 1), 1) - 1) * por_pagina
    return {"data": achadas[inicio:inicio + por_pagina], "total_count": total}


def periodo_fechado(data_iso: str) -> bool:
    return bool(data_iso) and data_iso <= PERIODO_FECHADO_ATE


def pagar(transacoes: list, dados: dict) -> dict:
    """Registra um pagamento que quita o saldo. Devolve {"ok": True, ...} ou {"ok": False, "erro": ...}."""
    alvo = next((t for t in transacoes if t["code"] == dados.get("code")), None)
    if alvo is None:
        return {"ok": False, "erro": "lançamento não encontrado"}
    if alvo["status"] not in ("PENDING", "OVERDUE"):
        return {"ok": False, "erro": "o lançamento já está %s" % alvo["status"]}
    if periodo_fechado(dados.get("date")):
        return {"ok": False, "erro": "período contábil fechado para %s" % dados.get("date")}
    if dados.get("account") not in dados_demo.CONTAS or dados.get("method") not in dados_demo.FORMAS:
        return {"ok": False, "erro": "conta bancária ou forma de pagamento inválida"}

    saldo = round(float(alvo["remaining_amount"]), 2)
    valor = round(float(dados.get("amount") or 0), 2)
    if abs(valor - saldo) > 0.005:
        return {"ok": False, "erro": "só aceito pagamento que quita o saldo (%.2f)" % saldo}
    liquido = round(valor - float(dados.get("discount") or 0) + float(dados.get("interest") or 0)
                    + float(dados.get("fine") or 0) + float(dados.get("fee") or 0)
                    + float(dados.get("other") or 0), 2)
    alvo.update({"status": "PAID", "paid_amount": liquido, "remaining_amount": 0.0,
                 "payment_date": dados["date"], "account_name": dados["account"],
                 "payment_method_name": dados["method"]})
    return {"ok": True, "liquido": liquido, "code": alvo["code"]}
