# -*- coding: utf-8 -*-
"""
Separa a aba de inserção em "completa" (tudo, com marcas) e "de trabalho" (só o que falta subir).

Usada pelo botão "Dar baixa" (no fim do modo real) e pelo "Gerar arquivo de inserção".

O que faz, para a aba de trabalho `Inserção - dd.mm`:
  1. Acha as linhas que NÃO devem subir para o sistema:
       * BAIXADA  -> a baixa foi feita por nós (Status da baixa = BAIXADA na Conferência NF);
       * IGNORAR  -> já estava pago no sistema. Em teoria não deveria acontecer sem ser por
                     nós, então vai marcada "JÁ PAGA NO SISTEMA — revisar".
  2. Guarda uma cópia da aba em backup/.
  3. Cria (na 1ª vez) a aba `Inserção - dd.mm completa`: cópia de TUDO, antes de retirar qualquer
     linha. Nas vezes seguintes só marca as linhas retiradas (coluna "Situação", com cor).
  4. RETIRA essas linhas da aba de trabalho (que passa a ter só o que precisa subir).
  5. Renumera a coluna Linha da Conferência NF, para os botões seguintes continuarem batendo.

Antes de retirar, confere que a linha da aba ainda tem a mesma NF e o mesmo valor da
conferência; se você editou, NÃO retira e avisa. Nada é apagado sem a cópia em backup/ e a
aba completa.
"""

from __future__ import annotations

from datetime import datetime

import config as cfg
import criterio
import planilha
from aplicar_na_planilha import colunas

MOTIVO_BAIXADA = "BAIXADA no sistema"
MOTIVO_JA_PAGA = "JÁ PAGA NO SISTEMA — revisar (não foi baixada por nós)"
MOTIVO_SOMA = "BAIXADA no sistema junto com outra linha (mesma nota, valores somados)"
COR_BAIXADA = 0xC4E7C4     # verde claro (BGR)
COR_JA_PAGA = 0x9CDCFF     # laranja claro
COL_SITUACAO = cfg.ULTIMA_COL_INSERCAO + 1    # R


def nome_completa(aba: str) -> str:
    n = aba + cfg.SUFIXO_COMPLETA
    if len(n) > 31:                       # limite do Excel para nome de aba
        n = aba[:31 - len(cfg.SUFIXO_COMPLETA)] + cfg.SUFIXO_COMPLETA
    return n


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _chave(valores) -> tuple:
    """Identidade de uma linha (colunas A..Q) para achá-la na aba completa."""
    def norm(v):
        if v in (None, ""):
            return ""
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        n = _num(v)
        if n is not None and not isinstance(v, str):
            return "%.2f" % n
        return " ".join(str(v).split()).upper()
    return tuple(norm(v) for v in list(valores)[:cfg.ULTIMA_COL_INSERCAO])


def _numero_doc(v) -> str:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return criterio.normalizar(v)


def ler_decisoes(ws_conf) -> list:
    """Linhas da tabela da Conferência NF: [{row, linha, nf, valor, acao, status, msg}]."""
    topo = cfg.LINHA_TABELA_RESULTADO
    fim = planilha.ultima_linha(ws_conf)
    if fim <= topo:
        return []
    COL = colunas(ws_conf)
    bloco = ws_conf.Range(ws_conf.Cells(topo + 1, 1), ws_conf.Cells(fim, COL["_ultima"])).Value
    if bloco and not isinstance(bloco[0], (tuple, list)):
        bloco = (bloco,)
    saida = []
    for i, v in enumerate(bloco or ()):
        bruto = v[COL["linha"] - 1]
        try:
            linha = int(bruto) if bruto not in (None, "") and not isinstance(bruto, str) else None
        except (TypeError, ValueError):
            linha = None
        saida.append({
            "row": topo + 1 + i, "linha": linha,
            "nf": _numero_doc(v[COL["nf"] - 1]), "valor": _num(v[COL["valor"] - 1]),
            "acao": str(v[COL["acao"] - 1] or "").strip().upper(),
            "status": str(v[COL["exec_status"] - 1] or "").strip().upper(),
            "msg": str(v[COL["exec_msg"] - 1] or "").strip(),
            "fornecedor": str(v[COL["fornecedor"] - 1] or "").strip(),
            "obs": str(v[COL["obs"] - 1] or "").strip(),
        })
    return saida


def planejar(ws_conf, ws_aba):
    """Quais linhas seriam retiradas. Devolve (cabecalho, todas as linhas, retirar, avisos)."""
    decisoes = ler_decisoes(ws_conf)
    cabecalho, linhas = planilha.ler_insercao(ws_aba)
    por_linha = {l["linha"]: l for l in linhas}
    retirar, avisos = [], []
    for d in decisoes:
        if d["linha"] is None:
            continue
        if d["status"].startswith("BAIXADA"):
            motivo = MOTIVO_BAIXADA
        elif d["acao"] == criterio.IGNORAR:
            # a 2ª linha de uma soma é IGNORAR porque a baixa é feita na 1ª: não é "já paga por outros"
            motivo = MOTIVO_SOMA if "a baixa é feita na linha" in d["obs"] else MOTIVO_JA_PAGA
        else:
            continue
        item = por_linha.get(d["linha"])
        if item is None:
            avisos.append("linha %d (NF %s): não existe mais na aba — não retirada" % (d["linha"], d["nf"]))
            continue
        vals = item["valores"]
        valor_atual = _num(vals[cfg.COL_VALOR - 1])
        if _numero_doc(vals[cfg.COL_NUMERO_DOCUMENTO - 1]) != d["nf"] or (
                d["valor"] is not None and valor_atual is not None
                and abs(valor_atual - d["valor"]) > cfg.TOLERANCIA_IGUAL):
            avisos.append("linha %d: a NF/valor na aba não bate mais com a conferência (você editou?) "
                          "— não retirada" % d["linha"])
            continue
        retirar.append({"linha": d["linha"], "row_conf": d["row"], "motivo": motivo,
                        "valores": vals, "nf": d["nf"], "fornecedor": d["fornecedor"],
                        "msg": d["msg"]})
    return cabecalho, linhas, retirar, avisos


def _marcar(wc, retirar):
    """Marca na aba completa cada linha retirada (texto na coluna Situação + cor)."""
    cab, linhas_c = planilha.ler_insercao(wc)
    fim = planilha.ultima_linha(wc)
    situacoes = wc.Range(wc.Cells(2, COL_SITUACAO), wc.Cells(max(fim, 2), COL_SITUACAO)).Value
    if situacoes is not None and not isinstance(situacoes, (tuple, list)):
        situacoes = ((situacoes,),)
    livres = {}
    for l in linhas_c:
        marca = situacoes[l["linha"] - 2][0] if situacoes and l["linha"] - 2 < len(situacoes) else None
        if marca in (None, ""):
            livres.setdefault(_chave(l["valores"]), []).append(l["linha"])
    quando = datetime.now().strftime("%d/%m/%Y %H:%M")
    for r in retirar:
        fila = livres.get(_chave(r["valores"]))
        if fila:
            linha_c = fila.pop(0)
        else:                                   # não estava na cópia (linha nova): acrescenta
            linha_c = planilha.acrescentar_insercao(wc, [r["valores"]])
        texto = "%s — %s%s" % (r["motivo"], quando, (" — " + r["msg"][:120]) if r["msg"] else "")
        wc.Cells(linha_c, COL_SITUACAO).Value = texto
        cor = COR_BAIXADA if r["motivo"] in (MOTIVO_BAIXADA, MOTIVO_SOMA) else COR_JA_PAGA
        wc.Range(wc.Cells(linha_c, 1), wc.Cells(linha_c, COL_SITUACAO)).Interior.Color = cor


def _renumerar_conferencia(ws_conf, retiradas_linhas: set, aba: str):
    """Depois de apagar linhas da aba, acerta a coluna Linha da Conferência NF."""
    topo = cfg.LINHA_TABELA_RESULTADO
    COL = colunas(ws_conf)
    for d in ler_decisoes(ws_conf):
        if d["linha"] is None:
            continue
        if d["linha"] in retiradas_linhas:
            ws_conf.Cells(d["row"], COL["linha"]).Value = "(retirada)"
            if not d["status"]:
                ws_conf.Cells(d["row"], COL["exec_status"]).Value = "RETIRADA DA ABA"
                ws_conf.Cells(d["row"], COL["exec_msg"]).Value = (
                    "retirada de '%s' (linha IGNORAR); ver '%s'" % (aba, nome_completa(aba)))
        else:
            antes = sum(1 for x in retiradas_linhas if x < d["linha"])
            if antes:
                ws_conf.Cells(d["row"], COL["linha"]).Value = d["linha"] - antes


def separar(wb, aba: str, ws_conf, confirmar=None) -> dict:
    """Faz a separação descrita no topo. `confirmar(retirar)` (opcional) pode vetar.

    Devolve {retiradas, avisos, completa, cancelado}. Não salva: quem chama decide
    (com a planilha aberta no seu Excel, fica sem salvar).
    """
    ws = wb.Worksheets(aba)
    cabecalho, todas, retirar, avisos = planejar(ws_conf, ws)
    resultado = {"retiradas": retirar, "avisos": avisos, "completa": None, "cancelado": False}
    if not retirar:
        return resultado
    if confirmar is not None and not confirmar(retirar):
        resultado["cancelado"] = True
        resultado["retiradas"] = []
        return resultado

    planilha.backup_aba(ws, "%s - antes de separar" % aba.replace("/", "-"))

    nome = nome_completa(aba)
    if nome in [w.Name for w in wb.Worksheets]:
        wc = wb.Worksheets(nome)
    else:
        wc = planilha.criar_aba_insercao(
            wb, cabecalho, [l["valores"] for l in todas], nome=nome, depois_de=ws,
            cadastro=False, extra_cabecalho=["Situação"])
    _marcar(wc, retirar)
    resultado["completa"] = nome

    linhas_apagar = sorted({r["linha"] for r in retirar}, reverse=True)
    antes = len(todas)
    for n in linhas_apagar:
        ws.Rows(n).Delete()
    depois = len(planilha.ler_insercao(ws)[1])
    if antes - depois != len(linhas_apagar):
        raise RuntimeError(
            "Conferi a aba '%s' depois de retirar as linhas e a contagem não bateu (%d → %d, esperado -%d). "
            "A cópia está em backup/ e na aba '%s'. Não mexa e me avise." % (
                aba, antes, depois, len(linhas_apagar), nome))
    _renumerar_conferencia(ws_conf, set(linhas_apagar), aba)
    return resultado


def texto_retiradas(resultado: dict) -> str:
    """Resumo para a janela preta."""
    linhas = []
    r = resultado["retiradas"]
    if not r:
        return ""
    baixadas = [x for x in r if x["motivo"] in (MOTIVO_BAIXADA, MOTIVO_SOMA)]
    ja = [x for x in r if x["motivo"] == MOTIVO_JA_PAGA]
    linhas.append("  Retiradas da aba de trabalho: %d (%d baixadas, %d já pagas no sistema)." % (
        len(r), len(baixadas), len(ja)))
    linhas.append("  A cópia completa, com as marcas, está na aba '%s'." % resultado["completa"])
    if ja:
        linhas.append("")
        linhas.append("  >> REVISAR: estas linhas já estavam PAGAS no sistema (não foi a nossa baixa):")
        for x in ja[:30]:
            v = x["valores"]
            linhas.append("     NF %-8s %-30s %s" % (x["nf"], (x["fornecedor"] or "")[:30],
                                                     criterio.formatar_valor(_num(v[cfg.COL_VALOR - 1]))))
    for a in resultado["avisos"]:
        linhas.append("  ! " + a)
    return "\n".join(linhas)
