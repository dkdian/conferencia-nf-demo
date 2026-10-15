# -*- coding: utf-8 -*-
"""
Escreve o resultado da conferência na aba "Conferência NF" da planilha.

A aba Inserção NÃO é tocada: é ela que vira o arquivo de upload.

Layout v2 da tabela (cabeçalho na linha LINHA_TABELA_RESULTADO), (layout atual):

  PLANILHA   A Linha | B Fornecedor | C Descrição | D Nº NF | E Valor | F Vencimento |
  (azul)     G Data pagamento | H Conta | I Forma | J Placa             -> o robô copia
  SISTEMA    K ID | L Fornecedor | M Descrição | N Valor | O Vencimento | P Situação |
  (laranja)  Q Data pagamento | R Placa                               -> o robô (melhor candidata)
  DECISÃO    S Resultado | T Ação | U ID a usar | V Valor a usar | W Data pagamento a usar |
  (amarelo)  X Conta a usar | Y Forma a usar | Z Descrição a usar | AA Observações
                                                            -> o robô sugere, você edita
  EXECUÇÃO   AB Status da baixa | AC Data/hora | AD Mensagem     -> o robô (baixa)

A baixa e o arquivo de inserção usam SEMPRE as colunas "a usar" e "Ação". "Valor a usar" é o
da planilha; a diferença para o saldo do sistema entra na baixa como juros ou desconto.
"Descrição a usar" só vem preenchida nas linhas a IMPORTAR que saem de um lançamento dividido
por placa (o arquivo de inserção usa esse texto no lugar da descrição da Inserção).

`colunas(ws)` reconhece o layout pela linha de cabeçalho (há também um layout antigo, de 24 colunas).

Uso avulso (reclassifica o último resultado sem consultar o sistema de novo):
    python aplicar_na_planilha.py
"""

from __future__ import annotations

import json
import sys
from datetime import datetime

import config as cfg
import criterio
import planilha

# --------------------------------------------------------------- layout ---
_NOMES_V2 = ["linha", "fornecedor", "descricao", "nf", "valor", "vencimento", "data_pgto",
             "conta", "forma", "placa",
             "s_id", "s_fornecedor", "s_descricao", "s_valor", "s_venc", "s_situacao",
             "s_pagamento", "s_placa",
             "resultado", "acao", "id_usar", "valor_usar", "data_usar", "conta_usar",
             "forma_usar", "descricao_usar", "obs",
             "exec_status", "exec_quando", "exec_msg"]
COL = {n: i for i, n in enumerate(_NOMES_V2, 1)}
ULTIMA_COL = COL["exec_msg"]
COL_OPCOES = 40   # AN (oculta): lista da coluna Ação e, em AN5, a aba que originou a conferência

_NOMES_V1 = ["linha", "fornecedor", "nf", "valor", "vencimento", "data_pgto", "conta",
             "forma", "s_id", "s_fornecedor", "s_valor", "s_venc", "s_situacao",
             "resultado", "acao", "id_usar", "valor_usar", "data_usar", "conta_usar",
             "forma_usar", "obs", "exec_status", "exec_quando", "exec_msg"]
COL_V1 = {n: i for i, n in enumerate(_NOMES_V1, 1)}
COL_V1["_ultima"] = len(_NOMES_V1)
COL["_ultima"] = ULTIMA_COL

CABECALHOS = ["Linha", "Fornecedor", "Descrição", "Nº NF", "Valor", "Vencimento",
              "Data pagamento", "Conta", "Forma", "Placa",
              "ID", "Fornecedor", "Descrição", "Valor", "Vencimento", "Situação",
              "Data pagamento", "Placa",
              "Resultado", "Ação", "ID a usar", "Valor a usar",
              "Data pagamento a usar", "Conta a usar", "Forma a usar", "Descrição a usar",
              "Observações",
              "Status da baixa", "Data/hora", "Mensagem"]
CABECALHOS_V1 = ["Linha", "Fornecedor", "Nº NF", "Valor", "Vencimento",
                 "Data pagamento", "Conta", "Forma",
                 "ID", "Fornecedor", "Valor", "Vencimento", "Situação",
                 "Resultado", "Ação", "ID a usar", "Valor a usar",
                 "Data pagamento a usar", "Conta a usar", "Forma a usar",
                 "Observações",
                 "Status da baixa", "Data/hora", "Mensagem"]
LARGURAS = [7, 30, 38, 11, 13, 12, 13, 20, 12, 14,
            9, 30, 38, 13, 12, 11, 13, 14,
            18, 12, 10, 14, 14, 20, 12, 38, 70,
            16, 16, 40]

# cores (BGR, que é como o Excel espera): planilha = azul, sistema = laranja, decisão = amarelo
AZUL_FUNDO, AZUL_FONTE, AZUL_CAB = 0xF7EBDD, 0x794E1F, 0xE6C29B
LARANJA_FUNDO, LARANJA_FONTE, LARANJA_CAB = 0xD6E4FC, 0x0C3C84, 0x84B0F4
AMARELO_FUNDO = 0xE6FAFF
BLOCOS = [("PLANILHA (azul — o robô copia)", COL["linha"], COL["placa"], AZUL_CAB),
          ("SISTEMA (laranja — melhor candidata)", COL["s_id"], COL["s_placa"], LARANJA_CAB),
          ("DECISÃO (amarelo — você edita Ação e \"a usar\")", COL["resultado"],
           COL["obs"], 0xCCF2FF),
          ("EXECUÇÃO (baixa)", COL["exec_status"], COL["exec_msg"], 0xD9D9D9)]

ACOES = [criterio.DAR_BAIXA, criterio.IMPORTAR, criterio.IGNORAR]

ROTULO = {
    criterio.JA_NO_SISTEMA: "JÁ NO SISTEMA",
    criterio.AJUSTAR_VALOR: "AJUSTAR VALOR",
    criterio.REVISAR: "REVISAR",
    criterio.NAO_ENCONTRADA: "NÃO ENCONTRADA",
    criterio.ERRO: "ERRO NA CONSULTA",
}
COR = {  # BGR, que é como o Excel espera
    criterio.JA_NO_SISTEMA: 0xC4E7C4,      # verde claro
    criterio.AJUSTAR_VALOR: 0xF5E4BD,      # azul claro
    criterio.REVISAR: 0x99F2FF,            # amarelo claro
    criterio.NAO_ENCONTRADA: 0xC5C5FF,     # vermelho claro
    criterio.ERRO: 0x9CDCFF,               # laranja claro
}
VERMELHO_PAGO, VERDE_ABERTO = 0x0000C0, 0x007000


def _data(iso):
    if not iso:
        return None
    try:
        return planilha.serial_excel(datetime.strptime(str(iso)[:10], "%Y-%m-%d"))
    except ValueError:
        return str(iso)


def _cab_igual(lidos, esperados) -> bool:
    def n(v):
        return " ".join(str(v or "").split()).lower()
    return len(lidos) >= len(esperados) and all(
        n(a) == n(b) for a, b in zip(lidos, esperados))


def colunas(ws) -> dict:
    """Mapa nome -> nº da coluna do layout que a aba Conferência NF tem AGORA (v2 ou v1).

    Chave "_ultima" = última coluna da tabela. Aba ainda sem cabeçalho = layout v2.
    """
    topo = cfg.LINHA_TABELA_RESULTADO
    if ws.Cells(topo, 1).Value in (None, ""):
        return COL
    lidos = ws.Range(ws.Cells(topo, 1), ws.Cells(topo, ULTIMA_COL)).Value
    lidos = list(lidos[0]) if lidos else []
    if _cab_igual(lidos, CABECALHOS):
        return COL
    if _cab_igual(lidos, CABECALHOS_V1):
        return COL_V1
    raise SystemExit("Layout da aba '%s' não reconhecido (cabeçalho da linha %d mudou?).\n"
                     "Rode o botão 1 (Conferir notas) para refazer a aba." % (ws.Name, topo))


def carregar_resultado() -> dict:
    if not cfg.ARQUIVO_RESULTADO.exists():
        raise SystemExit(
            "Não achei %s.\nRode antes: python conferir_nf.py" % cfg.ARQUIVO_RESULTADO)
    return json.loads(cfg.ARQUIVO_RESULTADO.read_text(encoding="utf-8"))


def aba_conferida(ws_conf) -> str:
    """Nome da aba de inserção que originou a conferência ('' se não registrada)."""
    for celula in (cfg.CEL_ABA_CONFERIDA, cfg.CEL_ABA_CONFERIDA_V1):
        v = ws_conf.Cells(*celula).Value
        if v:
            return str(v).strip()
    return ""


def resumo(classificacao: dict) -> dict:
    cont = {k: 0 for k in ROTULO}
    for r in classificacao.values():
        cont[r["resultado"]] = cont.get(r["resultado"], 0) + 1
    return cont


def montar_linha(item: dict, r: dict) -> list:
    """Uma linha da tabela (30 células, layout v2)."""
    cand = r.get("candidata") or {}
    acao = r["acao"]
    id_usar = valor_usar = data_usar = conta_usar = forma_usar = None

    # "a usar" só para lançamento EM ABERTO que esta linha pode baixar
    if r.get("preencher"):
        id_usar = r.get("id")
        valor_usar = r.get("valor_a_usar")
        data_usar = _data(item.get("data_pagamento"))
        conta_usar = item.get("conta") or None
        forma_usar = item.get("forma") or None

    linha = [None] * ULTIMA_COL

    def poe(nome, valor):
        linha[COL[nome] - 1] = valor

    poe("linha", item["linha"]); poe("fornecedor", item.get("fornecedor") or "")
    poe("descricao", item.get("descricao") or None)
    poe("nf", item["numero"]); poe("valor", item.get("valor"))
    poe("vencimento", _data(item.get("vencimento")))
    poe("data_pgto", _data(item.get("data_pagamento")))
    poe("conta", item.get("conta") or None); poe("forma", item.get("forma") or None)
    poe("placa", item.get("placa") or None)
    poe("s_id", cand.get("id")); poe("s_fornecedor", cand.get("fornecedor"))
    poe("s_descricao", cand.get("descricao"))
    poe("s_valor", cand.get("valor")); poe("s_venc", _data(cand.get("vencimento")))
    poe("s_situacao", r.get("situacao") or None)
    poe("s_pagamento", _data(cand.get("pagamento")))
    poe("s_placa", cand.get("placa"))
    poe("resultado", ROTULO.get(r["resultado"], r["resultado"]))
    poe("acao", acao or None)
    poe("id_usar", id_usar); poe("valor_usar", valor_usar)
    poe("data_usar", data_usar); poe("conta_usar", conta_usar); poe("forma_usar", forma_usar)
    poe("descricao_usar", r.get("descricao_a_usar"))
    poe("obs", r.get("observacao") or None)
    return linha


# colunas da Decisão e da Execução que são preservadas quando a baixa já foi tentada
_PRESERVAR = ["resultado", "acao", "id_usar", "valor_usar", "data_usar", "conta_usar",
              "forma_usar", "descricao_usar", "obs", "exec_status", "exec_quando", "exec_msg"]


def _ler_decisoes_executadas(ws, topo, ultima):
    """Linhas cuja baixa já foi tentada: a decisão e a execução delas são preservadas
    quando a conferência roda de novo (nunca se perde o registro de uma baixa).

    Lê pelo NOME da coluna do layout que a aba tem agora (v1 ou v2) e devolve
    {(linha, nf): {nome: valor}}, para regravar no layout novo.
    """
    if ultima <= topo:
        return {}
    c = colunas(ws)
    bloco = ws.Range(ws.Cells(topo + 1, 1), ws.Cells(ultima, c["_ultima"])).Value
    if bloco and not isinstance(bloco[0], (tuple, list)):
        bloco = (bloco,)
    guardadas = {}
    for valores in bloco or ():
        if valores[c["exec_status"] - 1] not in (None, ""):
            chave = (valores[c["linha"] - 1], criterio.normalizar(valores[c["nf"] - 1]))
            guardadas[chave] = {nome: valores[c[nome] - 1] for nome in _PRESERVAR if nome in c}
    return guardadas


def _pegar_ou_criar_aba(ws_insercao):
    wb = ws_insercao.Parent
    for aba in wb.Worksheets:
        if aba.Name == cfg.ABA_RESULTADO:
            return aba
    nova = wb.Worksheets.Add(None, ws_insercao)  # (Before, After): logo depois da Inserção
    nova.Name = cfg.ABA_RESULTADO
    return nova


def gravar_aba_resultado(ws_insercao, dados: dict):
    linhas, consultas = dados["linhas"], dados["consultas"]
    quando = datetime.fromisoformat(dados["gerado_em"])
    classificacao = criterio.classificar_linhas(linhas, consultas)
    cont = resumo(classificacao)
    contagem_acao = {a: 0 for a in ACOES + [""]}
    for r in classificacao.values():
        contagem_acao[r["acao"]] = contagem_acao.get(r["acao"], 0) + 1

    ws = _pegar_ou_criar_aba(ws_insercao)
    topo = cfg.LINHA_TABELA_RESULTADO
    ultima_antiga = planilha.ultima_linha(ws)

    # Antes de apagar: guarda cópia e separa o que já foi baixado
    preservar = {}
    if ultima_antiga > topo:
        planilha.backup_aba(ws, "Conferência NF - antes de reconferir")
        preservar = _ler_decisoes_executadas(ws, topo, ultima_antiga)

    # Só limpa células: os botões (formas) continuam onde estão.
    ws.Range(ws.Cells(1, 1), ws.Cells(max(ultima_antiga, 3000), COL_OPCOES)).Clear()
    try:
        ws.Range(ws.Cells(topo + 1, 1), ws.Cells(3000, ULTIMA_COL)).Validation.Delete()
    except Exception:
        pass
    # no layout v1 a coluna Z (26) era a lista oculta; agora é a Descrição a usar
    for i in range(1, ULTIMA_COL + 1):
        ws.Columns(i).Hidden = False

    ws.Range("A1").Value = "Conferência das Notas Fiscais no sistema"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Última conferência: %s   |   Aba conferida: %s" % (
        quando.strftime("%d/%m/%Y %H:%M"), ws_insercao.Name)
    # de qual aba veio (o Gerar arquivo e o Dar baixa seguem esta aba, sem adivinhar)
    ws.Cells(*cfg.CEL_ABA_CONFERIDA).Value = ws_insercao.Name
    ws.Range("A3").Value = (
        "Já no sistema: %d   |   Ajustar valor: %d   |   Não encontradas: %d   |   "
        "Revisar: %d   |   Falhas: %d") % (
        cont[criterio.JA_NO_SISTEMA], cont[criterio.AJUSTAR_VALOR],
        cont[criterio.NAO_ENCONTRADA], cont[criterio.REVISAR], cont[criterio.ERRO])
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = (
        "Ações sugeridas: Dar baixa %d   |   Importar %d   |   Ignorar %d   |   "
        "Em branco (você decide) %d   ||   Azul = planilha · Laranja = sistema · "
        "Amarelo = você edita" % (
            contagem_acao[criterio.DAR_BAIXA], contagem_acao[criterio.IMPORTAR],
            contagem_acao[criterio.IGNORAR], contagem_acao[""]))
    ws.Range("A4").Font.Italic = True

    # Blocos e cabeçalho
    for titulo, c1, c2, cor in BLOCOS:
        ws.Cells(cfg.LINHA_GRUPOS_RESULTADO, c1).Value = titulo
        faixa = ws.Range(ws.Cells(cfg.LINHA_GRUPOS_RESULTADO, c1),
                         ws.Cells(cfg.LINHA_GRUPOS_RESULTADO, c2))
        faixa.Interior.Color = cor
        faixa.Font.Bold = True
    for i, titulo in enumerate(CABECALHOS, 1):
        c = ws.Cells(topo, i)
        c.Value = titulo
        c.Font.Bold = True
        c.WrapText = True
        c.VerticalAlignment = -4108   # centralizado
        c.Interior.Color = 0xE0E0E0
    for c1, c2, cab, fonte in ((COL["linha"], COL["placa"], AZUL_CAB, AZUL_FONTE),
                               (COL["s_id"], COL["s_placa"], LARANJA_CAB, LARANJA_FONTE)):
        faixa = ws.Range(ws.Cells(topo, c1), ws.Cells(topo, c2))
        faixa.Interior.Color = cab
        faixa.Font.Color = fonte
    for i, largura in enumerate(LARGURAS, 1):
        ws.Columns(i).ColumnWidth = largura

    itens = sorted(linhas, key=lambda i: i["linha"])
    ultima = topo + len(itens)
    if itens:
        ws.Range(ws.Cells(topo + 1, COL["nf"]),
                 ws.Cells(ultima, COL["nf"])).NumberFormat = "@"   # NF como texto

        matriz = [montar_linha(item, classificacao[item["linha"]]) for item in itens]
        for i, item in enumerate(itens):
            chave = (item["linha"], criterio.normalizar(item["numero"]))
            for nome, valor in (preservar.get(chave) or {}).items():
                matriz[i][COL[nome] - 1] = valor
        ws.Range(ws.Cells(topo + 1, 1), ws.Cells(ultima, ULTIMA_COL)).Value = \
            tuple(tuple(l) for l in matriz)

        fmt = planilha.formatos(ws)

        def formato(nome, codigo):
            ws.Range(ws.Cells(topo + 1, COL[nome]), ws.Cells(ultima, COL[nome])).NumberFormat = codigo
        for nome in ("valor", "s_valor", "valor_usar"):
            formato(nome, fmt["numero"])
        for nome in ("vencimento", "data_pgto", "s_venc", "s_pagamento", "data_usar"):
            formato(nome, fmt["data"])
        formato("exec_quando", fmt["data_hora"])

        # Cores por origem: planilha = azul, sistema = laranja (fundo e letra)
        for c1, c2, fundo, fonte in ((COL["linha"], COL["placa"], AZUL_FUNDO, AZUL_FONTE),
                                     (COL["s_id"], COL["s_placa"], LARANJA_FUNDO, LARANJA_FONTE)):
            faixa = ws.Range(ws.Cells(topo + 1, c1), ws.Cells(ultima, c2))
            faixa.Interior.Color = fundo
            faixa.Font.Color = fonte
        for nome in ("descricao", "s_descricao", "obs", "descricao_usar"):
            ws.Range(ws.Cells(topo + 1, COL[nome]), ws.Cells(ultima, COL[nome])).WrapText = True

        # Colunas que você edita ficam com fundo amarelo bem claro
        ws.Range(ws.Cells(topo + 1, COL["acao"]),
                 ws.Cells(ultima, COL["descricao_usar"])).Interior.Color = AMARELO_FUNDO

        for i, item in enumerate(itens):
            r = classificacao[item["linha"]]
            linha_x = topo + 1 + i
            ws.Cells(linha_x, COL["resultado"]).Interior.Color = COR[r["resultado"]]
            if r.get("situacao"):
                cel = ws.Cells(linha_x, COL["s_situacao"])
                cel.Font.Bold = True
                cel.Font.Color = VERMELHO_PAGO if r["situacao"] == criterio.PAGO else VERDE_ABERTO
            cand = r.get("candidata") or {}
            if item.get("valor") is not None and cand.get("valor") is not None \
                    and abs(float(item["valor"]) - float(cand["valor"])) > cfg.TOLERANCIA_IGUAL:
                ws.Cells(linha_x, COL["valor"]).Font.Bold = True
                ws.Cells(linha_x, COL["s_valor"]).Font.Bold = True

        # Lista suspensa da coluna Ação (lista em coluna oculta: independe do
        # separador de listas do Excel do usuário)
        for i, a in enumerate(ACOES, 1):
            ws.Cells(i, COL_OPCOES).Value = a
        ws.Columns(COL_OPCOES).Hidden = True
        val = ws.Range(ws.Cells(topo + 1, COL["acao"]), ws.Cells(ultima, COL["acao"])).Validation
        val.Delete()
        letra = _letra(COL_OPCOES)
        val.Add(3, 1, 1, "=$%s$1:$%s$%d" % (letra, letra, len(ACOES)))   # xlValidateList, alerta de parar
        val.IgnoreBlank = True
        val.InCellDropdown = True
        val.ErrorTitle = "Ação"
        val.ErrorMessage = "Escolha: " + ", ".join(ACOES) + " (ou deixe em branco)."

    ws.Activate()
    ws.Range("A1").Select()
    return cont, classificacao


def _letra(n: int) -> str:
    texto = ""
    while n:
        n, resto = divmod(n - 1, 26)
        texto = chr(65 + resto) + texto
    return texto


def main() -> int:
    dados = carregar_resultado()
    planilha.escolher_aba(dados.get("aba", ""))   # a aba que originou este resultado
    print("Resultado gerado em %s" % dados.get("gerado_em", "?"))
    print("Aplicando %d linhas na aba '%s'..." % (len(dados["linhas"]), cfg.ABA_RESULTADO))

    ja_aberta = planilha.ja_esta_aberta()
    with planilha.planilha_aberta(gravar=True) as ws:
        cont, _ = gravar_aba_resultado(ws, dados)

    print("")
    for chave, texto in (("JA_NO_SISTEMA", "Já no sistema  "), ("AJUSTAR_VALOR", "Ajustar valor  "),
                         ("NAO_ENCONTRADA", "Não encontradas"), ("REVISAR", "Para revisar   ")):
        print("  %s : %d" % (texto, cont[getattr(criterio, chave)]))
    if cont[criterio.ERRO]:
        print("  Com falha       : %d" % cont[criterio.ERRO])
    print("")
    if ja_aberta:
        print("Resultado escrito na aba '%s' da planilha aberta. "
              "Ainda NÃO foi salvo: salve quando quiser." % cfg.ABA_RESULTADO)
    else:
        print("Resultado gravado e salvo na aba '%s'." % cfg.ABA_RESULTADO)
    return 0


if __name__ == "__main__":
    sys.exit(main())
