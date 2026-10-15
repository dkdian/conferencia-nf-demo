# -*- coding: utf-8 -*-
"""Do que o ERP devolve para o formato que o critério entende (sem navegador, sem Excel).

Fica separado para o robô e os testes usarem exatamente o mesmo código.
"""

from __future__ import annotations

import criterio


def resumir(registro: dict) -> dict:
    """Campos úteis de um lançamento devolvido pelo ERP, com os nomes que o critério usa."""
    def pega(nome):
        return registro.get(nome) if registro.get(nome) not in (None, "") else None
    return {
        "id": pega("code"), "descricao": pega("description"), "fornecedor": pega("party_name"),
        "numero_documento": pega("document_number"), "valor": pega("amount"),
        "valor_pago": pega("paid_amount"), "vencimento": pega("due_date"),
        "pagamento": pega("payment_date"), "tipo": "Despesa",
        # usados pela baixa e pela prévia
        "uuid": registro.get("id"), "status": pega("status"), "saldo": pega("remaining_amount"),
        "conta": pega("account_name"), "forma": pega("payment_method_name"),
        "categoria": pega("category_name"), "placa": pega("equipment_name"),
        "tipo_documento": pega("document_type"),
    }


def registro_de_lote(numero: str, lidas_lote: list, paginas: int) -> dict:
    """Registro de um número a partir da resposta de um LOTE (busca exata)."""
    alvo = criterio.normalizar(numero)
    iguais = [t for t in lidas_lote if criterio.normalizar(t.get("numero_documento")) == alvo]
    return {"numero": numero, "status": "LIDO", "quantidade": len(iguais),
            "total_no_sistema": len(iguais), "lidas": len(iguais), "paginas_lidas": paginas,
            "transacoes_iguais": iguais, "transacoes_parecidas": [], "outros_numeros": [],
            "via": "lote"}


def registro_de_leitura(numero: str, lidas: list, total: int, paginas: int) -> dict:
    """Registro de um número buscado SOZINHO, do jeito "contém" (enxerga ano+zeros e zero à esquerda)."""
    cls = criterio.separar(numero, lidas, total)
    iguais_ids = {id(t) for t in cls["transacoes_iguais"]}
    outros = []
    for t in lidas:
        doc = t.get("numero_documento")
        if id(t) not in iguais_ids and doc and doc not in outros:
            outros.append(doc)
    return {
        "numero": numero, "status": "LIDO", "quantidade": len(cls["transacoes_iguais"]),
        "total_no_sistema": total, "lidas": len(lidas), "paginas_lidas": paginas,
        "transacoes_iguais": cls["transacoes_iguais"], "transacoes_parecidas": cls["transacoes_parecidas"],
        "outros_numeros": outros[:8], "via": "individual",
    }
