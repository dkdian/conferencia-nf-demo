# -*- coding: utf-8 -*-
"""
Gera o Inserção.xlsx (o arquivo do botão Importar do sistema) a partir da aba
Inserção, respeitando as decisões da aba Conferência NF. A aba Inserção NÃO é
alterada.

Entra no arquivo:
  * linhas SEM nota fiscal (não passam pela conferência)
  * linhas com Ação = IMPORTAR (se a Conferência NF tem "Descrição a usar" preenchida, o arquivo
    leva essa descrição no lugar da da Inserção: linhas divididas por placa)

Fica de fora:
  * linhas cuja baixa já foi feita com sucesso (Status da baixa = BAIXADA)
  * linhas com Ação = IGNORAR

Fica de fora e é AVISADO (para não gerar duplicata nem perder linha sem ninguém ver):
  * Ação = DAR BAIXA ainda não executada
  * Ação em branco (REVISAR que você ainda não decidiu)

Nota Fiscal que não consta na Conferência NF (ex.: linha nova depois da conferência)
entra no arquivo, mas com aviso: rode a conferência de novo para ter certeza.

Antes de gerar, retira da aba as linhas baixadas/IGNORAR (guardando tudo na aba "... completa",
ver completa.py), se você confirmar.

Antes de gerar, confere se a Inserção ainda é a mesma da conferência (o número da
nota e o valor de cada linha conferida). Se mudou, NÃO gera.

Uso:  python gerar_insercao.py [--saida caminho.xlsx]
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

import completa
import config as cfg
import criterio
import planilha
from aplicar_na_planilha import colunas, aba_conferida

STATUS_BAIXADA = "BAIXADA"


def _texto(v) -> str:
    return "" if v is None else str(v).strip()


def _num(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _mostrar_numero(v) -> str:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return _texto(v)


def _numero_doc(v) -> str:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return criterio.normalizar(v)


def ler_decisoes(ws_conf) -> dict:
    """{linha da Inserção: decisão} lida da tabela da aba Conferência NF."""
    topo = cfg.LINHA_TABELA_RESULTADO
    fim = planilha.ultima_linha(ws_conf)
    if fim <= topo:
        return {}
    COL = colunas(ws_conf)                    # layout v2 ou o antigo (v1, sem Descrição a usar)
    bloco = ws_conf.Range(ws_conf.Cells(topo + 1, 1), ws_conf.Cells(fim, COL["_ultima"])).Value
    if bloco and not isinstance(bloco[0], (tuple, list)):
        bloco = (bloco,)
    decisoes = {}
    for v in bloco or ():
        linha = v[COL["linha"] - 1]
        if linha in (None, "") or isinstance(linha, str):      # "(retirada)" = já saiu da aba
            continue
        decisoes[int(linha)] = {
            "nf": _numero_doc(v[COL["nf"] - 1]),
            "valor": _num(v[COL["valor"] - 1]),
            "acao": _texto(v[COL["acao"] - 1]).upper(),
            "exec_status": _texto(v[COL["exec_status"] - 1]).upper(),
            "descricao_usar": _texto(v[COL["descricao_usar"] - 1]) if "descricao_usar" in COL else "",
        }
    return decisoes


def decidir(cabecalho, linhas, decisoes):
    """Separa as linhas da Inserção. Devolve (incluir, excluidas, retidas, avisos, erros)."""
    i_tipo, i_num, i_valor = (cfg.COL_TIPO_DOCUMENTO - 1, cfg.COL_NUMERO_DOCUMENTO - 1,
                              cfg.COL_VALOR - 1)
    incluir, excluidas, retidas, sem_conferencia, erros = [], [], [], [], []

    for item in linhas:
        n, vals = item["linha"], item["valores"]
        d = decisoes.get(n)
        eh_nf = (cfg.TIPO_DOCUMENTO_ALVO in _texto(vals[i_tipo]).lower()
                 and _texto(vals[i_num]) != "")

        if d is None:
            if eh_nf:
                sem_conferencia.append(item)
            incluir.append(item)
            continue

        # a linha conferida ainda é a mesma?
        valor_atual = _num(vals[i_valor])
        if _numero_doc(vals[i_num]) != d["nf"] or (
                d["valor"] is not None and valor_atual is not None
                and abs(valor_atual - d["valor"]) > cfg.TOLERANCIA_IGUAL):
            erros.append("linha %d: na conferência era NF %s, R$ %s; na Inserção agora é NF %s, R$ %s"
                         % (n, d["nf"], d["valor"], _mostrar_numero(vals[i_num]) or "(vazio)",
                            valor_atual))
            continue

        if d["exec_status"].startswith(STATUS_BAIXADA):
            excluidas.append((item, "baixada no sistema"))
        elif d["acao"] == criterio.IGNORAR:
            excluidas.append((item, "IGNORAR"))
        elif d["acao"] == criterio.IMPORTAR:
            if d.get("descricao_usar"):      # lançamento dividido por placa: mesma descrição do editado
                item = dict(item, descricao_usar=d["descricao_usar"])
            incluir.append(item)
        elif d["acao"] == criterio.DAR_BAIXA:
            retidas.append((item, "DAR BAIXA ainda não executada"))
        else:
            retidas.append((item, "Ação em branco (sem decisão)"))

    return incluir, excluidas, retidas, sem_conferencia, erros


def gravar_xlsx(cabecalho, incluir, destino: Path):
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cfg.ABA_UPLOAD    # o sistema espera a aba chamada Inserção
    ws.append([c for c in cabecalho])
    for item in incluir:
        linha = []
        for j, v in enumerate(item["valores"]):
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            if j == cfg.COL_DESCRICAO - 1 and item.get("descricao_usar"):
                v = item["descricao_usar"]
            linha.append(v)
        ws.append(linha)
    for c in ws[1]:
        c.font = Font(bold=True)
    for j, c in enumerate(cabecalho, 1):
        letra = openpyxl.utils.get_column_letter(j)
        larg = max([len(str(c or ""))] +
                   [len(str(i["valores"][j - 1] or "")) for i in incluir[:200]])
        ws.column_dimensions[letra].width = min(max(larg + 2, 10), 45)
        for cel in ws[letra][1:]:
            if isinstance(cel.value, datetime):
                cel.number_format = "dd/mm/yyyy"
    ws.freeze_panes = "A2"
    wb.save(destino)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--saida", help="caminho alternativo do .xlsx gerado")
    ap.add_argument("--aba", default="", help="aba de inserção (padrão: a que foi conferida)")
    ap.add_argument("--sem-perguntas", action="store_true", help="só para teste: não pergunta S/N")
    args = ap.parse_args()
    destino = Path(args.saida) if args.saida else cfg.ARQ_INSERCAO

    print("=" * 70)
    print("  Gerar arquivo de inserção -- %s" % destino.name)
    print("=" * 70)

    # A aba usada é a que foi CONFERIDA (registrada na Conferência NF). Os números de
    # linha da conferência só valem para ela.
    with planilha.com_workbook() as wb:
        try:
            conferida = aba_conferida(wb.Worksheets(cfg.ABA_RESULTADO))
        except Exception:
            conferida = ""
    if args.aba:
        planilha.escolher_aba(args.aba)
        if conferida and conferida != args.aba:
            print("\n  A conferência atual é da aba '%s', não de '%s'. NADA foi gerado." % (conferida, args.aba))
            print("  Rode a conferência (botão verde) nessa aba, ou use a aba conferida.")
            return 1
    elif conferida:
        planilha.escolher_aba(conferida)          # sem pergunta: segue a conferência
        print("\n  Usando a aba '%s' (a que foi conferida)." % conferida)
    else:
        print("\n  A Conferência NF não registra de qual aba veio; usando a aba padrão.")
        planilha.escolher_aba()

    if not planilha.avisar_dias_repetidos(cfg.ABA, args.sem_perguntas, "gerar o arquivo de importação"):
        print("\n  Cancelado. Nada foi gerado.")
        return 1

    # Retira da aba as linhas que não devem subir (baixadas / já pagas no sistema), guardando
    # tudo na aba "... completa". Se você não quiser, o arquivo sai certo do mesmo jeito.
    def confirma_retirada(retirar):
        b = sum(1 for x in retirar if x["motivo"] in (completa.MOTIVO_BAIXADA, completa.MOTIVO_SOMA))
        print("\n  Linhas a retirar da aba '%s': %d (%d baixadas, %d já pagas no sistema)." % (
            cfg.ABA, len(retirar), b, len(retirar) - b))
        print("  Nada se perde: cópia em backup/ e na aba '%s'." % completa.nome_completa(cfg.ABA))
        return planilha.confirmar_sn("Retirar essas linhas da aba agora?", args.sem_perguntas)

    def separar():
        with planilha.com_workbook(gravar=True) as wb:
            try:
                ws_c = wb.Worksheets(cfg.ABA_RESULTADO)
            except Exception:
                return None
            return completa.separar(wb, cfg.ABA, ws_c, confirmar=confirma_retirada)

    resultado_sep = planilha.com_tentativas(separar)
    if resultado_sep and resultado_sep["retiradas"]:
        print("\n" + completa.texto_retiradas(resultado_sep))
        if planilha.ja_esta_aberta():
            print("\n  A planilha está aberta: a retirada foi feita na janela, SEM salvar. Salve com Ctrl+S.")
    elif resultado_sep and resultado_sep["cancelado"]:
        print("\n  Retirada não feita; o arquivo sai sem as baixadas/IGNORAR do mesmo jeito.")

    with planilha.planilha_aberta(aba=cfg.ABA) as ws_ins:
        cabecalho, linhas = planilha.ler_insercao(ws_ins)
        try:
            ws_conf = ws_ins.Parent.Worksheets(cfg.ABA_RESULTADO)
            decisoes = ler_decisoes(ws_conf)
        except Exception:
            decisoes = {}

    if not decisoes:
        print("\n  ! A aba '%s' está vazia ou não existe: nenhuma nota foi conferida." %
              cfg.ABA_RESULTADO)

    incluir, excluidas, retidas, sem_conf, erros = decidir(cabecalho, linhas, decisoes)

    if erros:
        print("\n  A aba Inserção mudou desde a conferência. NADA foi gerado.")
        for e in erros[:15]:
            print("   - " + e)
        print("\n  Rode a conferência de novo (botão verde) e tente outra vez.")
        return 1

    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists():
        cfg.BACKUP.mkdir(parents=True, exist_ok=True)
        shutil.copy2(destino, cfg.BACKUP / ("Inserção anterior %s.xlsx" %
                                            datetime.now().strftime("%Y%m%d_%H%M%S")))
    try:
        gravar_xlsx(cabecalho, incluir, destino)
    except PermissionError:
        print("\n  ! Não consegui gravar %s: o arquivo está aberto. Feche-o e tente de novo."
              % destino)
        return 1

    cfg.SAIDA.mkdir(parents=True, exist_ok=True)
    (cfg.SAIDA / "ultima_insercao.json").write_text(json.dumps({
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "arquivo": str(destino),
        "incluidas": [i["linha"] for i in incluir],
        "excluidas": [{"linha": i["linha"], "motivo": m} for i, m in excluidas],
        "retidas": [{"linha": i["linha"], "motivo": m} for i, m in retidas],
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    print("")
    print("  Linhas na aba Inserção : %d" % len(linhas))
    print("  Entraram no arquivo    : %d" % len(incluir))
    print("  Ficaram de fora        : %d  (%d baixadas/IGNORAR)" % (
        len(excluidas) + len(retidas), len(excluidas)))
    print("")
    print("  Arquivo: %s" % destino)

    trocadas = [i for i in incluir if i.get("descricao_usar")]
    if trocadas:
        print("")
        print("  Descrição trocada em %d linha(s) (lançamento dividido por placa; a aba Inserção "
              "não foi alterada):" % len(trocadas))
        for item in trocadas[:30]:
            print("     linha %-4d -> %s" % (item["linha"], item["descricao_usar"][:80]))

    if retidas:
        print("")
        print("  >> ATENÇÃO: %d linha(s) NÃO entraram porque falta decisão ou baixa:" % len(retidas))
        for item, motivo in retidas[:30]:
            v = item["valores"]
            print("     linha %-4d %-28s NF %-8s -> %s" % (
                item["linha"], _texto(v[cfg.COL_FORNECEDOR - 1])[:28],
                _mostrar_numero(v[cfg.COL_NUMERO_DOCUMENTO - 1]), motivo))
        print("     Resolva na aba Conferência NF e gere de novo.")
    if sem_conf:
        print("")
        print("  >> ATENÇÃO: %d linha(s) com Nota Fiscal NÃO estão na Conferência NF "
              "(entraram no arquivo sem conferir):" % len(sem_conf))
        for item in sem_conf[:15]:
            v = item["valores"]
            print("     linha %-4d NF %s" % (item["linha"], _mostrar_numero(v[cfg.COL_NUMERO_DOCUMENTO - 1])))
        print("     Rode a conferência para ter certeza de que não duplicam.")
    return 0


if __name__ == "__main__":
    sys.exit(planilha.executar(main))
