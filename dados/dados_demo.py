# -*- coding: utf-8 -*-
"""Dados 100% fictícios da demonstração: o ERP de mentira e a planilha da semana saem daqui.

Cada linha da planilha foi desenhada para exercitar uma situação real da conferência
(ver testes/test_cenarios.py, que confere o resultado esperado de cada uma).
Nomes de empresas, valores e números de nota são inventados.
"""

from __future__ import annotations

from datetime import date

HOJE = date(2026, 10, 14)          # "hoje" do ERP de mentira (fixo, para o vídeo e os testes serem repetíveis)
ABA_SEMANA = "Inserção - 13.10"

CONTAS = ["Conta Operações", "Conta Matriz", "Conta Reembolsos"]
FORMAS = ["Boleto", "PIX", "Débito", "Transferência", "Depósito"]

# ----------------------------------------------------------------- ERP ---
# (código, documento, fornecedor, descrição, valor, vencimento, status, pago_em, placa)
# status: "ABERTO" (vira PENDING/OVERDUE conforme o vencimento) ou "PAGO"
_ERP = [
    (7101, "1204", "Mecânica Alfa Ltda", "NF Manutenção preventiva de frota", 1250.00, "2026-10-02", "ABERTO", None, ""),
    (7102, "3381", "Posto Beta Combustíveis Ltda", "NF Abastecimento de caminhões", 480.00, "2026-10-03", "ABERTO", None, ""),
    (7103, "2250", "Auto Pneus Sigma Ltda", "NF Troca de pneus", 1780.00, "2026-10-04", "ABERTO", None, ""),
    (7104, "587", "Peças Gama Comércio Ltda", "NF Peças de reposição", 850.00, "2026-10-05", "ABERTO", None, ""),
    (7105, "587", "Omicron Embalagens Ltda", "NF Caixas e fitas", 70.00, "2026-10-12", "ABERTO", None, ""),
    (7106, "774", "Ferragens Ômega Ltda", "NF Parcela 1/3 ferragens", 640.00, "2026-09-07", "PAGO", "2026-09-07", ""),
    (7107, "165", "Hidráulica Ni Ltda", "NF Conexões hidráulicas", 1100.00, "2026-10-10", "ABERTO", None, ""),
    (7108, "5560", "Lubrificantes Kappa Ltda", "NF Óleo lubrificante", 725.40, "2026-10-08", "ABERTO", None, ""),
    (7109, "4410", "Oficina Lambda ME", "NF Serviço de alinhamento", 200.00, "2026-10-08", "ABERTO", None, ""),
    (7110, "202600000088", "Sementes Zeta Agro Ltda", "NF Sementes de milho", 3500.00, "2026-10-08", "ABERTO", None, ""),
    (7111, "7001", "Épsilon Serviços Ltda", "NF Consultoria técnica", 1000.00, "2026-10-09", "ABERTO", None, ""),
    (7112, "321", "Papelaria Mu Ltda", "NF Material de escritório", 85.90, "2026-10-02", "PAGO", "2026-10-09", ""),
    (7113, "6100", "Logística Rô Transportes Ltda", "NF Transporte de carga", 4800.00, "2026-10-09", "ABERTO", None, ""),
    (7115, "3050", "Hidráulica Ni Ltda", "NF Mangueiras hidráulicas", 1420.00, "2026-10-09", "ABERTO", None, ""),
    (7116, "3390", "Posto Beta Combustíveis Ltda", "NF Abastecimento de máquinas", 560.00, "2026-10-09", "ABERTO", None, ""),
    (7117, "1777", "Gama Autopeças Ltda", "NF Filtros", 295.00, "2026-10-06", "ABERTO", None, ""),
    # ruído: números parecidos com os da planilha e outros lançamentos da empresa
    (7120, "12040", "Mecânica Alfa Ltda", "NF Pastilhas de freio", 300.00, "2026-10-20", "ABERTO", None, ""),
    (7121, "9020", "Borracharia Chi Ltda", "NF Remendos e válvulas", 90.00, "2026-10-15", "ABERTO", None, ""),
    (7122, "10", "Gráfica Psi Ltda", "NF Impressos", 410.00, "2026-10-25", "ABERTO", None, ""),
    (7123, "2001", "Eletrônica Fi ME", "NF Sensores", 735.50, "2026-10-18", "ABERTO", None, ""),
    (7124, "5000", "Cimento Tau Ltda", "NF Cimento e areia", 5120.00, "2026-10-30", "ABERTO", None, ""),
    (7125, "611", "Vidraçaria Upsilon ME", "NF Vidros", 320.00, "2026-09-28", "PAGO", "2026-09-30", ""),
    (7126, "845", "Autoelétrica Delta Sul Ltda", "NF Alternador", 1290.00, "2026-10-16", "ABERTO", None, ""),
    (7127, "3382", "Posto Beta Combustíveis Ltda", "NF Abastecimento de caminhões", 515.00, "2026-10-17", "ABERTO", None, ""),
    (7128, "77", "Limpeza Sigma ME", "NF Produtos de limpeza", 188.00, "2026-10-12", "ABERTO", None, ""),
    (7129, "1205", "Mecânica Alfa Ltda", "NF Troca de óleo", 380.00, "2026-10-22", "ABERTO", None, ""),
    (7130, "4411", "Oficina Lambda ME", "NF Balanceamento", 120.00, "2026-10-19", "ABERTO", None, ""),
    (7131, "9013", "Transportes Delta Ltda", "NF Frete (outro cliente)", 2300.00, "2026-11-05", "ABERTO", None, ""),
]


def transacoes_erp() -> list:
    """Lançamentos do ERP de mentira, já com status PENDING / OVERDUE / PAID."""
    saida = []
    for codigo, doc, fornecedor, desc, valor, venc, situacao, pago_em, placa in _ERP:
        if situacao == "PAGO":
            status, pago, saldo = "PAID", valor, 0.0
        else:
            status = "OVERDUE" if venc < HOJE.isoformat() else "PENDING"
            pago, saldo = 0.0, valor
        saida.append({
            "id": "tx-%d" % codigo, "code": codigo, "document_number": doc,
            "party_name": fornecedor, "description": desc, "amount": valor,
            "paid_amount": pago, "remaining_amount": saldo, "due_date": venc,
            "payment_date": pago_em, "status": status,
            "account_name": "Conta Operações" if situacao == "PAGO" else None,
            "payment_method_name": "Boleto" if situacao == "PAGO" else None,
            "equipment_name": placa or None, "category_name": None,
            "document_type": "Nota Fiscal",
        })
    return saida


# ------------------------------------------------------------ planilha ---
# Colunas da aba (A..Q): Tipo, Descrição, Valor, Vencimento, Competência, Data de pagamento,
# Status, Tipo Documento, Numero Documento, Categoria, Conta bancária, Forma de pagamento,
# Colaborador/Cliente, Documento colaborador/cliente, Observação, Centro de Custo, Placa
CABECALHO = ["Tipo (Obrigatório)", "Descrição (Obrigatório)", "Valor (Obrigatório)",
             "Data de vencimento (Obrigatório)", "Competência", "Data de pagamento", "Status",
             "Tipo Documento", "Numero Documento", "Categoria", "Conta bancária",
             "Forma de pagamento", "Colaborador/Cliente", "Documento colaborador/cliente",
             "Observação", "Centro de Custo", "Placa"]


def _nf(desc, valor, venc, pgto, nf, forn, forma, cat="MANUTENÇÃO", conta="Conta Operações",
        centro="Operações Sul", placa=""):
    return dict(desc=desc, valor=valor, venc=venc, pgto=pgto, tipodoc="Nota Fiscal", nf=nf,
                forn=forn, forma=forma, cat=cat, conta=conta, centro=centro, placa=placa)


def _sem(desc, valor, venc, pgto, forn, forma, cat, conta="Conta Operações", centro="Administrativo"):
    return dict(desc=desc, valor=valor, venc=venc, pgto=pgto, tipodoc="", nf="",
                forn=forn, forma=forma, cat=cat, conta=conta, centro=centro, placa="")


LINHAS = [
    _sem("Aluguel do escritório 10/12", 4200.00, "2026-10-05", "2026-10-05", "Imobiliária Pi", "Boleto", "ALUGUEL"),
    _nf("Manutenção preventiva de frota", 1250.00, "2026-10-02", "2026-10-05", "1204", "Mecânica Alfa", "Boleto"),
    _nf("Abastecimento de caminhões", 480.00, "2026-10-03", "2026-10-05", "3381", "Posto Beta", "Boleto", "COMBUSTÍVEL"),
    _sem("Seguro de vida dos colaboradores", 1180.30, "2026-10-05", "2026-10-05", "Seguradora Rô", "Boleto", "SEGUROS"),
    _nf("Troca de pneus", 1780.00, "2026-10-04", "2026-10-06", "2250", "Auto Pneus Sigma", "", "PNEUS"),
    _nf("Peças de reposição", 910.00, "2026-10-05", "2026-10-06", "587", "Peças Gama", "Boleto", "PEÇAS"),
    _nf("Reparo de pneus", 150.00, "2026-10-06", "2026-10-06", "902", "Borracharia Pi", "Pix", "PNEUS"),
    _sem("Tarifas bancárias", 96.40, "2026-10-06", "2026-10-06", "Banco Exemplo", "Débito", "TARIFAS"),
    _nf("Filtros", 295.00, "2026-10-06", "2026-10-07", "1777", "Gama Peças", "Boleto", "PEÇAS"),
    _nf("Parcela 2/3 ferragens", 640.00, "2026-10-07", "2026-10-07", "774", "Ferragens Ômega", "Boleto", "FERRAGENS"),
    _nf("Frete de insumos", 2300.00, "2026-10-06", "2026-10-07", "9012", "Transportes Delta", "Pix", "FRETES"),
    _sem("Adiantamento salarial", 1000.00, "2026-10-07", "2026-10-07", "Colaborador A", "Pix", "FOLHA"),
    _nf("Material elétrico", 330.00, "2026-10-07", "2026-10-08", "165", "Elétrica Teta", "Pix", "ELÉTRICA"),
    _nf("Óleo lubrificante", 725.40, "2026-10-08", "2026-10-08", "5560", "Lubrificantes Kappa", "Boleto", "LUBRIFICANTES"),
    _nf("Serviço de alinhamento", 199.60, "2026-10-08", "2026-10-08", "4410", "Oficina Lambda", "Pix", "SERVIÇOS"),
    _sem("Consultoria contábil 3/6", 2650.00, "2026-10-08", "2026-10-08", "Contabilidade Chi", "Boleto", "CONSULTORIAS"),
    _nf("Sementes de milho", 3500.00, "2026-10-08", "2026-10-09", "88", "Sementes Zeta", "Boleto", "INSUMOS", centro="Agrícola"),
    _nf("Consultoria técnica (parte 1)", 600.00, "2026-10-09", "2026-10-09", "7001", "Serviços Épsilon", "Boleto", "CONSULTORIAS"),
    _nf("Consultoria técnica (parte 2)", 400.00, "2026-10-09", "2026-10-09", "7001", "Serviços Épsilon", "Boleto", "CONSULTORIAS"),
    _nf("Material de escritório", 85.90, "2026-10-02", "2026-10-09", "321", "Papelaria Mu", "Pix", "ESCRITÓRIO"),
    _nf("Transporte de carga", 4800.00, "2026-10-09", "2026-10-09", "6100", "Logística Rô", "Boleto", "FRETES"),
    _nf("Mangueiras hidráulicas", 1420.00, "2026-10-09", "2026-10-09", "3050", "Hidráulica Ni", "Boleto", "HIDRÁULICA"),
    _nf("Abastecimento de máquinas", 560.00, "2026-10-09", "2026-10-09", "3390", "Posto Beta", "Pix", "COMBUSTÍVEL"),
    _sem("Adiantamento salarial", 800.00, "2026-10-09", "2026-10-09", "Colaborador B", "Pix", "FOLHA"),
    _sem("Licença de software mensal", 349.90, "2026-10-09", "2026-10-09", "Software Rho", "Débito", "TECNOLOGIA"),
]


# ---------------------------------------------------- rotina de terça (Drive) ---
HOJE_ROTINA = date(2026, 10, 13)                 # terça-feira em que a rotina "roda" na demonstração
SEMANA = (date(2026, 10, 5), date(2026, 10, 9))  # segunda a sexta da semana passada
RELATORIOS = ["Operações", "Administrativo"]     # os dois relatórios que o cliente deixa no Drive
CENTROS = {"OPS-SUL": "Operações Sul", "ADM": "Administrativo", "AGR": "Agrícola"}   # tabela de-para
COLUNAS_BRUTAS = ["DATA", "VENCIMENTO", "DESCRICAO", "ENTRADA", "SAIDA", "NF/OS/FATURA", "CATEGORIA",
                  "CONTA", "FORMA PGTO", "CLIENTE/FORNECEDOR", "CENTRO DE CUSTO", "PLACA"]


def relatorio_de(l: dict) -> str:
    return "Administrativo" if l["centro"] == "Administrativo" else "Operações"


LINHAS.sort(key=lambda l: (l["pgto"], RELATORIOS.index(relatorio_de(l))))   # estável

# Movimentos que estão nos relatórios mas fora da semana (a rotina precisa ignorá-los)
FORA_DA_SEMANA = [
    _nf("Manutenção de colhedora", 3100.00, "2026-09-25", "2026-09-29", "1198", "Mecânica Alfa", "Boleto"),
    _nf("Combustível de frota", 940.00, "2026-09-28", "2026-10-01", "3375", "Posto Beta", "Boleto", "COMBUSTÍVEL"),
    _sem("Aluguel do escritório 09/12", 4200.00, "2026-09-05", "2026-09-30", "Imobiliária Pi", "Boleto", "ALUGUEL"),
    _sem("Adiantamento salarial", 950.00, "2026-10-02", "2026-10-02", "Colaborador C", "Pix", "FOLHA"),
    _nf("Pneus de reposição", 2210.00, "2026-10-12", "2026-10-12", "2262", "Auto Pneus Sigma", "Boleto", "PNEUS"),
    _sem("Licença de software mensal", 349.90, "2026-10-12", "2026-10-12", "Software Rho", "Débito", "TECNOLOGIA"),
]


def linhas_brutas(relatorio: str) -> list:
    """Linhas de um relatório como o cliente as exporta (nomes de coluna e valores "sujos")."""
    codigo = {v: k for k, v in CENTROS.items()}
    saida = []
    for l in LINHAS + FORA_DA_SEMANA:
        if relatorio_de(l) != relatorio:
            continue
        saida.append({
            "DATA": l["pgto"], "VENCIMENTO": l["venc"], "DESCRICAO": l["desc"], "ENTRADA": None,
            "SAIDA": l["valor"], "NF/OS/FATURA": ("NF %s" % l["nf"]) if l["nf"] else None,
            "CATEGORIA": l["cat"], "CONTA": l["conta"], "FORMA PGTO": l["forma"].upper() if l["forma"] else None,
            "CLIENTE/FORNECEDOR": l["forn"].upper(), "CENTRO DE CUSTO": codigo[l["centro"]], "PLACA": l["placa"] or None,
        })
    return saida


# Resultado esperado de cada nota (usado pelo teste): nf -> (resultado, ação sugerida)
ESPERADO = {
    "1204": ("JA_NO_SISTEMA", "DAR BAIXA"),
    "3381": ("JA_NO_SISTEMA", "DAR BAIXA"),
    "2250": ("JA_NO_SISTEMA", "DAR BAIXA"),
    "587": ("REVISAR", ""),                 # valor R$ 910 na planilha x R$ 850 no sistema: juros de R$ 60
    "902": ("NAO_ENCONTRADA", "IMPORTAR"),
    "1777": ("REVISAR", ""),                # fornecedor só parecido
    "774": ("REVISAR", "IMPORTAR"),         # parcela anterior já paga em outra data
    "9012": ("NAO_ENCONTRADA", "IMPORTAR"),
    "165": ("NAO_ENCONTRADA", "IMPORTAR"),  # o número existe, mas de outro fornecedor e outro valor
    "5560": ("JA_NO_SISTEMA", "DAR BAIXA"),
    "4410": ("AJUSTAR_VALOR", "DAR BAIXA"),  # R$ 0,40 de diferença: entra como desconto
    "88": ("REVISAR", ""),                  # no sistema o número é 202600000088 (ano + zeros)
    "7001": ("JA_NO_SISTEMA", "DAR BAIXA"),  # as duas linhas somam o lançamento (a 2ª vira IGNORAR)
    "321": ("JA_NO_SISTEMA", "IGNORAR"),    # já paga no sistema, na mesma data
    "6100": ("JA_NO_SISTEMA", "DAR BAIXA"),
    "3050": ("JA_NO_SISTEMA", "DAR BAIXA"),
    "3390": ("JA_NO_SISTEMA", "DAR BAIXA"),
}

# Linhas que a pessoa decide à mão depois da conferência (o que acontece na vida real):
# nf -> ação escolhida
DECISOES_HUMANAS = {"587": "DAR BAIXA", "88": "DAR BAIXA", "1777": "DAR BAIXA"}
