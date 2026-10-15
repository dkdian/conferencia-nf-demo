# -*- coding: utf-8 -*-
"""
Botão 2 — dá baixa no ERP das linhas da aba Conferência NF com Ação = DAR BAIXA.

Fluxo (nada é gravado antes do passo 3):
  1. LÊ a aba Conferência NF: linhas com Ação = DAR BAIXA e sem baixa registrada.
  2. PRÉVIA: consulta cada ID no ERP e mostra, por linha, o que seria feito (valor, desconto/juros,
     data, conta, forma) e os problemas encontrados. Linha com problema fica "NÃO EXECUTADO":
     nunca é adivinhada.
  3. Você escolhe:
        SIMULAR       preenche a janela "Confirmar Pagamento" de cada linha na tela e clica CANCELAR.
                      A rede fica TRAVADA para qualquer requisição que possa gravar: a simulação
                      não tem como dar baixa.
        BAIXAR [N]    dá baixa DE VERDADE (só nas N primeiras, se N for dado).
        (Enter)       cancela.
     Só o modo real grava. Ele exige digitar a palavra BAIXAR (não há opção que pule isso).
  4. Real: para no PRIMEIRO erro. Depois de cada Confirmar, reconsulta o ID: precisa estar PAGO,
     com a data de pagamento da planilha. Só então registra BAIXADA.

O registro vai nas colunas Execução da aba (Status da baixa / Data-hora / Mensagem) e no log
saida/baixa_<data-hora>.jsonl.

Regras de cada linha (qualquer falha -> NÃO EXECUTADO, com o motivo):
  * o ID existe, está em aberto (PENDING/OVERDUE) e não é pagamento parcial;
  * fornecedor da planilha não é "diferente" do do ERP;
  * a diferença entre o saldo do ERP e o Valor a usar, de qualquer tamanho, vira desconto
    (planilha menor) ou juros (planilha maior) -- a prévia destaca as de R$ 1,00 ou mais;
  * há data de pagamento, conta e forma na planilha, e a conta/forma existem na tela;
  * o mesmo ID não aparece em duas linhas;
  * o período contábil da data não está fechado;
  * depois de preencher, o valor líquido lido na tela = Valor a usar e o saldo após = 0,00.

Uso:  python dar_baixa.py [--sim] [--limite N]
      --sim   pula a pergunta e faz SÓ a simulação (serve para testar)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import unicodedata
from datetime import datetime, timedelta

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Error as PWError

import completa
import config as cfg
import conferir_nf as tela
import criterio
import planilha
from aplicar_na_planilha import colunas, aba_conferida
from erp_ficticio import garantir

sys.path.insert(0, str(cfg.BASE / "dados"))
import dados_demo  # noqa: E402  (só para a lista de contas e formas do cadastro)

PENDENTES = ("PENDING", "OVERDUE")
CONTAS_CONHECIDAS = tuple(dados_demo.CONTAS)
FORMAS_CONHECIDAS = tuple(dados_demo.FORMAS)
RE_VALOR = r"R\$\s*(-?[\d\.]+,\d{2})"

# O que a tela faz por POST só para LER (a simulação deixa passar apenas estas rotas)
LEITURA = re.compile(r"/api/(transacoes/listar|periodo-fechado|config|saude|estado)")


# ----------------------------------------------------------------- utilidades ---

def _norm(t) -> str:
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().upper()


def _num_pt(texto: str) -> float:
    return float(texto.replace(".", "").replace(",", "."))


def _fmt_input(v: float) -> str:
    return ("%.2f" % v).replace(".", ",")


def _para_data(v):
    """Data da célula -> 'aaaa-mm-dd'. Aceita data do Excel, número de série e 'dd/mm/aaaa'."""
    if v in (None, ""):
        return None
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime("%Y-%m-%d")
    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", str(v))
    if m:
        return "%s-%02d-%02d" % (m.group(3), int(m.group(2)), int(m.group(1)))
    return None


def _agora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def pausa(segundos: float = 1.0):
    if cfg.RITMO_S:
        time.sleep(segundos * cfg.RITMO_S)


# ---------------------------------------------------------- leitura da aba ---

def _ler_linhas_ws(ws) -> list:
    """Linhas DAR BAIXA da Conferência NF que ainda podem ser baixadas."""
    topo, fim = cfg.LINHA_TABELA_RESULTADO, planilha.ultima_linha(ws)
    if fim <= topo:
        return []
    COL = colunas(ws)
    bloco = ws.Range(ws.Cells(topo + 1, 1), ws.Cells(fim, COL["_ultima"])).Value
    if bloco and not isinstance(bloco[0], (tuple, list)):
        bloco = (bloco,)
    linhas = []
    for v in bloco or ():
        def g(nome, _v=v):
            return _v[COL[nome] - 1]
        if isinstance(g("linha"), str) or g("linha") in (None, ""):
            continue                     # "(retirada)": já saiu da aba
        if _norm(g("acao")) != criterio.DAR_BAIXA:
            continue
        status = _norm(g("exec_status"))
        if status.startswith("BAIXADA") or status.startswith("ERRO") or status.startswith("RETIRADA"):
            continue          # já baixada (nunca de novo) ou com erro a resolver na mão
        linhas.append({
            "linha": int(g("linha")), "nf": str(g("nf") or "").strip(),
            "fornecedor": str(g("fornecedor") or "").strip(),
            "id": int(g("id_usar")) if g("id_usar") not in (None, "") else None,
            "valor": float(g("valor_usar")) if g("valor_usar") not in (None, "") else None,
            "data": _para_data(g("data_usar")),
            "conta": str(g("conta_usar") or "").strip(),
            "forma": str(g("forma_usar") or "").strip(),
        })
    return linhas


def registrar(linha: int, status: str, mensagem: str, log=None):
    """Escreve o resultado na aba (e no log), na hora, para nunca perder o registro."""
    if log is not None:
        log.write(json.dumps({"quando": _agora(), "linha": linha, "status": status,
                              "mensagem": mensagem}, ensure_ascii=False) + "\n")
        log.flush()

    def gravar():
        with planilha.planilha_aberta(aba=cfg.ABA_RESULTADO, gravar=True) as ws:
            topo, fim = cfg.LINHA_TABELA_RESULTADO, planilha.ultima_linha(ws)
            COL = colunas(ws)
            for r in range(topo + 1, fim + 1):
                if ws.Cells(r, COL["linha"]).Value == linha:
                    ws.Cells(r, COL["exec_status"]).Value = status
                    agora = datetime.now()
                    ws.Cells(r, COL["exec_quando"]).Value = planilha.serial_excel(agora) + \
                        (agora.hour * 3600 + agora.minute * 60) / 86400
                    ws.Cells(r, COL["exec_msg"]).Value = mensagem[:250]
                    ws.Cells(r, COL["exec_quando"]).NumberFormat = planilha.formatos(ws)["data_hora"]
                    return
    planilha.com_tentativas(gravar)


# -------------------------------------------------------------- ERP / tela ---

def _do_numero(texto):
    def f(r):
        return (cfg.ROTA_LISTAGEM in r.url and str(tela._corpo(r).get("document_number")) == texto)
    return f


def _achar_id(dados, id_):
    for r in dados:
        if r.get("code") == id_:
            return r
    return None


def _ler_busca(pag, st: dict, texto: str, paginar: bool) -> list:
    """Digita `texto` no filtro e devolve as linhas da resposta (e das próximas páginas, se `paginar`)."""
    with pag.expect_response(_do_numero(texto), timeout=cfg.TIMEOUT_CONSULTA_S * 1000) as e:
        st["campo"].fill("")
        st["campo"].fill(texto)
    st["filtro"] = texto
    conteudo = e.value.json()
    dados = tela._linhas_da_resposta(conteudo)
    total = int(conteudo.get("total_count") or 0)
    pagina = 1
    while paginar and len(dados) < total and pagina < cfg.MAX_PAGINAS:
        proxima = pagina + 1

        def resposta_da_pagina(r, _n=proxima):
            corpo = tela._corpo(r)
            return (cfg.ROTA_LISTAGEM in r.url and str(corpo.get("document_number")) == texto
                    and int(corpo.get("page") or 1) == _n)
        try:
            with pag.expect_response(resposta_da_pagina, timeout=8000) as e2:
                if not tela._proxima_pagina(pag):
                    raise PWTimeout("sem botão de próxima página")
            mais = tela._linhas_da_resposta(e2.value.json())
        except PWTimeout:
            break
        if not mais:
            break
        dados = dados + mais
        pagina = proxima
    return dados


def consultar_id(pag, st: dict, linha: dict):
    """Acha o ID na consulta da NF e devolve (registro do ID, todas as linhas lidas).

    1) Busca EXATA ("488,": com a vírgula o ERP compara o número inteiro): traz só os lançamentos
       daquela NF.
    2) Se o ID não estiver lá (número em formato ano+zeros), busca "contém", lendo todas as
       páginas. Isso é feito uma vez por NF por execução.
    """
    nf = linha["nf"]
    st.setdefault("texto", {})
    st.setdefault("contem", {})
    if nf not in st["cache"]:
        st["cache"][nf] = _ler_busca(pag, st, nf + ",", paginar=False)
        st["texto"][nf] = nf + ","
    reg = _achar_id(st["cache"][nf], linha["id"])
    if reg is None and not st["contem"].get(nf):
        st["contem"][nf] = True
        mais = _ler_busca(pag, st, nf, paginar=True)
        ja = {r.get("code") for r in st["cache"][nf]}
        st["cache"][nf] = st["cache"][nf] + [r for r in mais if r.get("code") not in ja]
        st["texto"][nf] = nf
        reg = _achar_id(st["cache"][nf], linha["id"])
    return reg, st["cache"][nf]


def _linha_do_id(pag, id_):
    return pag.locator("tbody tr").filter(has_text=str(id_))


def _id_visivel(pag, id_, espera_ms=2500) -> bool:
    try:
        _linha_do_id(pag, id_).first.wait_for(state="visible", timeout=espera_ms)
        return True
    except PWTimeout:
        return False


def garantir_filtro(pag, st: dict, linha: dict):
    """Deixa a tabela filtrada pela NF da linha e na página em que o ID aparece."""
    texto = st.get("texto", {}).get(linha["nf"], linha["nf"])
    if st.get("filtro") != texto:
        st["campo"].fill("")
        st["campo"].fill(texto)
        st["filtro"] = texto
    if _id_visivel(pag, linha["id"], 4000):
        return
    for recomecar in (False, True):
        if recomecar:
            st["campo"].fill("")
            st["campo"].fill(texto)
            pag.wait_for_timeout(1500)
            if _id_visivel(pag, linha["id"]):
                return
        for _ in range(cfg.MAX_PAGINAS):
            if not tela._proxima_pagina(pag):
                break
            pag.wait_for_timeout(1000)
            if _id_visivel(pag, linha["id"]):
                return
    raise PWTimeout("não achei a linha do ID %s em nenhuma página da NF %s" % (linha["id"], linha["nf"]))


def ler_apos_confirmar(pag, st: dict, linha: dict, clicar):
    """Clica em Confirmar e devolve o registro do ID como o ERP mostra DEPOIS.

    Usa a atualização da listagem que a própria tela faz depois de gravar. Nunca recarrega a página.
    """
    lidos = []

    def guarda(r):
        if cfg.ROTA_LISTAGEM in r.url:
            try:
                lidos.append(tela._linhas_da_resposta(r.json()))
            except Exception:
                pass
    pag.on("response", guarda)
    try:
        clicar()
        pag.wait_for_timeout(2500)
    finally:
        pag.remove_listener("response", guarda)
    for dados in reversed(lidos):
        reg = _achar_id(dados, linha["id"])
        if reg is not None:
            st["cache"][linha["nf"]] = dados
            return reg
    return None


def validar(linha: dict, reg) -> list:
    """Problemas que impedem a baixa (lista vazia = ok) + avisos em 'linha[avisos]'."""
    problemas, avisos = [], []
    if linha["id"] is None:
        return ["sem ID a usar"]
    if reg is None:
        return ["o ID %s não apareceu na consulta da NF %s" % (linha["id"], linha["nf"])]
    doc = reg.get("document_number")
    if str(doc).strip() != linha["nf"] and \
            criterio.normalizar(doc) != criterio.normalizar(linha["nf"]) and \
            not criterio.formato_ano_zeros(doc, linha["nf"]):
        problemas.append("o ID %s é do documento %s, não da NF %s" % (linha["id"], doc, linha["nf"]))
    if reg.get("status") not in PENDENTES:
        problemas.append("o lançamento está %s (só baixo PENDING/OVERDUE)" % reg.get("status"))
    saldo = reg.get("remaining_amount")
    if saldo is None or float(saldo) <= 0:
        problemas.append("sem saldo pendente")
    forn = criterio.nivel_fornecedor(linha["fornecedor"], reg.get("party_name"))
    if forn == "diferente":
        problemas.append("fornecedor diferente (planilha: %s / ERP: %s)" % (linha["fornecedor"], reg.get("party_name")))
    elif forn == "parecido":
        avisos.append("fornecedor só parecido (%s)" % reg.get("party_name"))
    if linha["valor"] is None:
        problemas.append("sem Valor a usar")
    elif saldo is not None and float(saldo) > 0:
        # A diferença de QUALQUER tamanho entra como juros (planilha maior) ou desconto (menor);
        # a conferência na tela (líquido = Valor a usar) é que protege.
        dif = round(linha["valor"] - float(saldo), 2)
        if abs(dif) >= cfg.LIMITE_AJUSTE:
            avisos.append("%s de %s (%.1f%% do saldo %s): confira o Valor a usar"
                          % ("JUROS" if dif > 0 else "DESCONTO", criterio.formatar_valor(abs(dif)),
                             abs(dif) / float(saldo) * 100, criterio.formatar_valor(saldo)))
    for campo_, nome in (("data", "Data pagamento a usar"), ("conta", "Conta a usar"), ("forma", "Forma a usar")):
        if not linha[campo_]:
            problemas.append("sem %s" % nome)
    if linha["conta"] and _norm(linha["conta"]) not in [_norm(c) for c in CONTAS_CONHECIDAS]:
        problemas.append("conta '%s' não existe na lista do ERP" % linha["conta"])
    if linha["forma"] and _norm(linha["forma"]) not in [_norm(f) for f in FORMAS_CONHECIDAS]:
        problemas.append("forma '%s' não existe na lista do ERP" % linha["forma"])
    linha["avisos"] = avisos
    return problemas


def calcular(linha: dict, reg) -> dict:
    saldo = round(float(reg["remaining_amount"]), 2)
    dif = round(saldo - linha["valor"], 2)
    return {"saldo": saldo, "desconto": dif if dif > 0 else 0.0, "juros": -dif if dif < 0 else 0.0}


def _janela(pag):
    return pag.locator("[role=dialog]").last


def _campo(jan, rotulo):
    return jan.locator("xpath=.//label[normalize-space()='%s']/following::input[1]" % rotulo).first


def _combo(jan, rotulo):
    return jan.locator(
        "xpath=.//label[normalize-space()='%s']/following::button[@role='combobox'][1]" % rotulo).first


def _escolher_opcao(pag, jan, rotulo, desejado: str):
    """Abre a lista, acha a opção igual (sem acento/maiúscula) e confere o que ficou."""
    combo = _combo(jan, rotulo)
    combo.click()
    pag.wait_for_timeout(400)
    opcoes = pag.get_by_role("option").all()
    achadas = [o for o in opcoes if _norm(o.inner_text()) == _norm(desejado)]
    if len(achadas) != 1:
        pag.keyboard.press("Escape")
        nomes = [o.inner_text() for o in opcoes]
        raise ValueError("%s: não há opção única igual a '%s' (opções: %s)" % (rotulo, desejado, nomes))
    achadas[0].click()
    pag.wait_for_timeout(300)
    if _norm(combo.inner_text()) != _norm(desejado):
        raise ValueError("%s: depois de escolher, a tela mostra '%s'" % (rotulo, combo.inner_text()))


def _texto_janela(jan) -> str:
    return re.sub(r"\s+", " ", jan.inner_text())


def preencher_e_conferir(pag, linha: dict, reg, calc: dict) -> dict:
    """Abre 'Confirmar Pagamento', preenche, relê tudo. Devolve o que a tela mostra.

    Levanta ValueError com o motivo se qualquer conferência falhar. NÃO confirma nem cancela.
    """
    # a linha é achada pelo identificador (e não pela posição): a tabela pode ser redesenhada logo depois do filtro
    tr = pag.locator("tbody tr[data-code='%s']" % linha["id"])
    try:
        tr.first.wait_for(state="visible", timeout=5000)
    except PWTimeout:
        raise ValueError("não achei a linha do ID %s na tabela" % linha["id"])
    tr.scroll_into_view_if_needed()
    botao = tr.locator("button[aria-label='Dar baixa']")     # nunca o botão de excluir
    try:
        botao.first.wait_for(state="visible", timeout=3000)
    except PWTimeout:
        pass
    if botao.count() != 1:
        raise ValueError("botão de pagar não identificado com certeza (%d candidatos)" % botao.count())
    pausa(0.8)
    botao.click()
    pag.wait_for_timeout(700)
    jan = _janela(pag)
    cab = _texto_janela(jan)
    if "Confirmar Pagamento" not in cab:
        raise ValueError("a janela aberta não é 'Confirmar Pagamento'")

    m = re.search(r"Saldo pendente:\s*" + RE_VALOR, cab)
    if not m or abs(_num_pt(m.group(1)) - calc["saldo"]) > 0.005:
        raise ValueError("saldo pendente na tela (%s) difere do consultado (%s)"
                         % (m.group(1) if m else "?", calc["saldo"]))

    _campo(jan, "Valor do Pagamento").fill(_fmt_input(calc["saldo"]))
    pausa(0.4)
    _campo(jan, "Desconto").fill(_fmt_input(calc["desconto"]))
    _campo(jan, "Juros").fill(_fmt_input(calc["juros"]))
    pausa(0.4)
    for rot in ("Multa", "Tarifa", "Outras taxas"):
        _campo(jan, rot).fill("0,00")

    # data: confere o período fechado (a tela consulta ao mudar a data)
    data = jan.locator("input[type=date]").first
    fechado = None
    if data.input_value() != linha["data"]:
        with pag.expect_response(lambda r: cfg.ROTA_PERIODO in r.url, timeout=8000) as e:
            data.fill(linha["data"])
        try:
            fechado = e.value.json()
        except Exception:
            fechado = None
        if not (isinstance(fechado, dict) and fechado.get("closed") is False):
            raise ValueError("período contábil fechado para %s (resposta: %s)" % (linha["data"], fechado))
    if data.input_value() != linha["data"]:
        raise ValueError("a data ficou %s em vez de %s" % (data.input_value(), linha["data"]))
    pausa(0.4)

    _escolher_opcao(pag, jan, "Conta Bancária", linha["conta"])
    _escolher_opcao(pag, jan, "Forma de Pagamento", linha["forma"])
    pag.wait_for_timeout(300)

    t = _texto_janela(jan)
    liq = re.search(r"Este será o valor considerado no pagamento:\s*" + RE_VALOR, t)
    apos = re.search(r"Saldo após este registro:\s*" + RE_VALOR, t)
    lido = {
        "liquido": _num_pt(liq.group(1)) if liq else None,
        "saldo_apos": _num_pt(apos.group(1)) if apos else None,
        "data": data.input_value(),
        "conta": _combo(jan, "Conta Bancária").inner_text().strip(),
        "forma": _combo(jan, "Forma de Pagamento").inner_text().strip(),
        "desconto": _campo(jan, "Desconto").input_value(),
        "juros": _campo(jan, "Juros").input_value(),
        "periodo_fechado": fechado,
    }
    if lido["liquido"] is None or abs(lido["liquido"] - linha["valor"]) > 0.005:
        raise ValueError("valor líquido na tela (%s) ≠ Valor a usar (%s)" % (lido["liquido"], linha["valor"]))
    if lido["saldo_apos"] is None or abs(lido["saldo_apos"]) > 0.005:
        raise ValueError("saldo após o registro seria %s (esperado 0,00): pagamento parcial" % lido["saldo_apos"])
    return lido


def fechar_janela(pag, confirmar: bool):
    jan = _janela(pag)
    nome = "Confirmar" if confirmar else "Cancelar"
    jan.get_by_role("button", name=re.compile(r"^\s*%s\s*$" % nome, re.I)).click()
    pag.wait_for_timeout(1200 if confirmar else 600)


# -------------------------------------------------------------------- prévia ---

def mostrar_previa(linhas: list):
    print("")
    print("=" * 78)
    print("  PRÉVIA — nada foi gravado")
    print("=" * 78)
    for l in linhas:
        c = l.get("calc")
        marca = "OK " if not l["problemas"] else "NÃO"
        print("\n [%s] linha %d | ID %s | NF %s | %s" % (marca, l["linha"], l["id"], l["nf"], l["fornecedor"]))
        if c:
            ajuste = ""
            if c["desconto"]:
                ajuste = " (desconto %s)" % criterio.formatar_valor(c["desconto"])
            if c["juros"]:
                ajuste = " (juros %s)" % criterio.formatar_valor(c["juros"])
            print("       pagar %s | saldo no sistema %s%s"
                  % (criterio.formatar_valor(l["valor"]), criterio.formatar_valor(c["saldo"]), ajuste))
            print("       data %s | conta %s | forma %s"
                  % (criterio.formatar_data(l["data"]), l["conta"], l["forma"]))
        for a in l.get("avisos", []):
            print("       aviso: " + a)
        for p in l["problemas"]:
            print("       PROBLEMA: " + p)
    ok = sum(1 for l in linhas if not l["problemas"])
    print("\n  %d linha(s) prontas, %d NÃO EXECUTADAS (com problema)." % (ok, len(linhas) - ok))
    grandes = [l for l in linhas if l.get("calc") and max(l["calc"]["desconto"], l["calc"]["juros"]) >= cfg.LIMITE_AJUSTE]
    if grandes:
        print("\n  ATENÇÃO — %d linha(s) com diferença de valor que entrará como JUROS/DESCONTO:" % len(grandes))
        for l in grandes:
            c = l["calc"]
            print("     linha %-4d ID %-8s saldo %s  valor a usar %s  -> %s %s" % (
                l["linha"], l["id"], criterio.formatar_valor(c["saldo"]), criterio.formatar_valor(l["valor"]),
                "JUROS" if c["juros"] else "DESCONTO", criterio.formatar_valor(c["juros"] or c["desconto"])))
        print("     (a tela só deixa confirmar se o valor líquido fechar com o Valor a usar)")


def perguntar(qtd_ok: int) -> tuple:
    """(modo, limite): modo em {'sim','real',None}."""
    print("\n  Digite:")
    print("    SIMULAR          preenche cada janela e CANCELA (não grava nada)")
    print("    BAIXAR           dá baixa DE VERDADE nas %d linhas prontas" % qtd_ok)
    print("    BAIXAR 1         dá baixa de verdade só na primeira linha pronta")
    print("    (Enter)          cancela")
    r = input("\n  > ").strip().upper()
    if r.startswith("SIMULAR"):
        return "sim", None
    m = re.fullmatch(r"BAIXAR(?:\s+(\d+))?", r)
    if m:
        return "real", int(m.group(1)) if m.group(1) else None
    return None, None


# ---------------------------------------------------------------------- main ---

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sim", action="store_true", help="só simulação, sem perguntar")
    ap.add_argument("--limite", type=int, help="máximo de linhas a processar")
    ap.add_argument("--sem-perguntas", action="store_true", help="só para teste: não pergunta S/N")
    args = ap.parse_args()
    garantir.garantir_servidor()

    print("=" * 78)
    print("  Dar baixa no ERP — demonstração")
    print("=" * 78)

    def ler_contexto():
        with planilha.com_workbook() as wb:
            try:
                ws_c = wb.Worksheets(cfg.ABA_RESULTADO)
            except Exception:
                return "", True, {}, []
            conf = aba_conferida(ws_c)
            existe = (not conf) or conf in planilha.abas_insercao(wb)
            rep = planilha.dias_repetidos(wb, conf) if conf and existe else {}
            return conf, existe, rep, _ler_linhas_ws(ws_c)

    conferida, existe, repetidos, linhas = planilha.com_tentativas(ler_contexto)
    if conferida:
        print("\n  As linhas da Conferência NF vêm da aba '%s'." % conferida)
        if not existe:
            print("  ! Essa aba não existe mais na planilha. Rode a conferência de novo. Nada foi feito.")
            return 1
        if repetidos and not (args.sem_perguntas or args.sim):
            print("\n  >> ATENÇÃO: a aba '%s' tem lançamentos de dias que também estão em outra aba:" % conferida)
            print(planilha.texto_dias_repetidos(repetidos))
            if not planilha.confirmar_sn("Você vai dar baixa em lançamentos dessas datas. Tem certeza que quer continuar?"):
                print("\n  Cancelado. Nada foi feito.")
                return 1
    if not linhas:
        print("\n  Nenhuma linha com Ação = DAR BAIXA e sem baixa registrada na aba '%s'." % cfg.ABA_RESULTADO)
        return 0
    print("\n  %d linha(s) marcadas DAR BAIXA." % len(linhas))

    contagem = {}
    for l in linhas:
        contagem[l["id"]] = contagem.get(l["id"], 0) + 1

    cfg.SAIDA.mkdir(parents=True, exist_ok=True)
    log = open(cfg.SAIDA / ("baixa_%s.jsonl" % datetime.now().strftime("%Y%m%d_%H%M%S")), "a", encoding="utf-8")
    resultados = []           # (linha, status, mensagem)
    modo_real = False

    with sync_playwright() as p:
        navegador, pag = tela.abrir_navegador(p)
        estado = {"trava": True}

        def rota(route):
            """Na simulação, bloqueia NA REDE tudo que possa gravar no ERP."""
            r = route.request
            if (not estado["trava"] or r.method in ("GET", "HEAD", "OPTIONS")
                    or not r.url.startswith(cfg.URL_ERP)
                    or (r.method == "POST" and LEITURA.search(r.url))):
                return route.continue_()
            print("   !!! BLOQUEADO na simulação: %s %s" % (r.method, r.url.split("?")[0]))
            return route.abort()
        pag.route("**/*", rota)

        try:
            pag.goto(cfg.URL_TRANSACOES, wait_until="domcontentloaded")
            pag.wait_for_timeout(1500)
            st = {"campo": tela.preparar_tela(pag), "cache": {}, "filtro": None}

            # ---- prévia (só leitura) ----
            for l in linhas:
                l["problemas"] = []
                if contagem[l["id"]] > 1:
                    l["problemas"].append("o ID %s está em %d linhas" % (l["id"], contagem[l["id"]]))
                try:
                    reg, _ = consultar_id(pag, st, l)
                except PWTimeout:
                    l["problemas"].append("o ERP não respondeu à consulta")
                    continue
                l["reg"] = reg
                l["problemas"] += validar(l, reg)
                if reg is not None and not l["problemas"]:
                    l["calc"] = calcular(l, reg)
            mostrar_previa(linhas)
            prontas = [l for l in linhas if not l["problemas"]]

            for l in linhas:
                if l["problemas"]:
                    msg = "NÃO EXECUTADO — " + "; ".join(l["problemas"])
                    registrar(l["linha"], "NÃO EXECUTADO", msg, log)
                    resultados.append((l["linha"], "NÃO EXECUTADO", msg))
            if not prontas:
                print("\n  Nenhuma linha pronta. Corrija a aba e rode de novo.")
                return 0

            if args.sim:
                modo, limite = "sim", args.limite
            else:
                modo, limite = perguntar(len(prontas))
                limite = limite or args.limite
            if modo is None:
                print("\n  Cancelado. Nada foi feito.")
                return 0
            modo_real = modo == "real"
            estado["trava"] = not modo_real
            fila = prontas[:limite] if limite else prontas
            print("\n  MODO %s — %d linha(s)." % ("REAL (grava no ERP)" if modo_real
                                                  else "SIMULAÇÃO (não grava)", len(fila)))

            # ---- execução ----
            for l in fila:
                print("\n  linha %d | ID %s ..." % (l["linha"], l["id"]))
                reg, _ = consultar_id(pag, st, l)          # reconsulta: pode ter mudado
                de_novo = validar(l, reg)
                if de_novo:
                    msg = "NÃO EXECUTADO — " + "; ".join(de_novo)
                    registrar(l["linha"], "NÃO EXECUTADO", msg, log)
                    resultados.append((l["linha"], "NÃO EXECUTADO", msg))
                    print("   " + msg)
                    continue
                calc = calcular(l, reg)
                try:
                    garantir_filtro(pag, st, l)
                    lido = preencher_e_conferir(pag, l, reg, calc)
                except (ValueError, PWTimeout, PWError) as e:
                    try:
                        fechar_janela(pag, confirmar=False)
                    except Exception:
                        pass
                    msg = "NÃO EXECUTADO — " + str(e)
                    registrar(l["linha"], "NÃO EXECUTADO", msg, log)
                    resultados.append((l["linha"], "NÃO EXECUTADO", msg))
                    print("   " + msg)
                    if modo_real:
                        print("\n  PAREI no primeiro problema (modo real).")
                        break
                    continue

                pag.screenshot(path=str(cfg.SAIDA / ("baixa_%s_%s.png" % (l["id"], "sim" if not modo_real else "real"))))
                pausa(1.2)
                if not modo_real:
                    fechar_janela(pag, confirmar=False)
                    msg = "SIMULAÇÃO OK — líquido %s, data %s, conta %s, forma %s (nada gravado)" % (
                        criterio.formatar_valor(lido["liquido"]), criterio.formatar_data(lido["data"]),
                        lido["conta"], lido["forma"])
                    registrar(l["linha"], "SIMULADO", msg, log)
                    resultados.append((l["linha"], "SIMULADO", msg))
                    print("   " + msg)
                    continue

                # ---- REAL: confirmar e verificar ----
                try:
                    depois = ler_apos_confirmar(pag, st, l, lambda: fechar_janela(pag, confirmar=True))
                except Exception as e:                       # qualquer falha depois do clique
                    depois = None
                    print("   ! falha ao verificar depois de Confirmar: %s" % e)
                ok = (depois is not None and depois.get("status") == "PAID"
                      and str(depois.get("payment_date"))[:10] == l["data"])
                if not ok:
                    msg = ("ERRO — cliquei em Confirmar e não consegui confirmar que colou (o ERP mostra status %s, "
                           "pagamento %s). CONFERIR NO ERP antes de qualquer coisa."
                           % (depois and depois.get("status"), depois and depois.get("payment_date")))
                    registrar(l["linha"], "ERRO", msg, log)
                    resultados.append((l["linha"], "ERRO", msg))
                    print("   " + msg)
                    print("\n  PAREI: erro na verificação depois da baixa.")
                    break
                msg = "BAIXADA — %s em %s (%s / %s)" % (
                    criterio.formatar_valor(lido["liquido"]), criterio.formatar_data(l["data"]),
                    lido["conta"], lido["forma"])
                registrar(l["linha"], "BAIXADA", msg, log)
                resultados.append((l["linha"], "BAIXADA", msg))
                print("   " + msg)
            pausa(1.5)
        finally:
            navegador.close()
            log.close()

    if modo_real and conferida:
        print("\n" + "-" * 78)
        print("  Separando a aba: retirando as linhas baixadas / já pagas (cópia em backup/ e na aba completa)...")
        try:
            def separar():
                with planilha.com_workbook(gravar=True) as wb:
                    return completa.separar(wb, conferida, wb.Worksheets(cfg.ABA_RESULTADO))
            res = planilha.com_tentativas(separar)
            print(completa.texto_retiradas(res) or "  Nada a retirar.")
            if planilha.ja_esta_aberta() and res["retiradas"]:
                print("\n  A planilha está aberta: a retirada foi feita na janela, SEM salvar. Salve com Ctrl+S.")
        except Exception as e:
            print("  ! Não consegui retirar as linhas da aba (%s)." % e)
            print("    As baixas já feitas estão registradas. Use o botão 3 (Gerar arquivo) para tentar de novo.")

    cont = {}
    for _, s, _m in resultados:
        cont[s] = cont.get(s, 0) + 1
    print("\n" + "-" * 78)
    print("  Resumo:", ", ".join("%s: %d" % kv for kv in sorted(cont.items())) or "nada")
    print("-" * 78)
    return 1 if cont.get("ERRO") else 0


if __name__ == "__main__":
    sys.exit(planilha.executar(main))
