# -*- coding: utf-8 -*-
"""
Botão 1 — confere, no ERP, quais notas fiscais da aba da semana já existem.

Como funciona
-------------
1. Lê a planilha (pelo próprio Excel, ver planilha.py) e separa as linhas cujo "Tipo Documento"
   é "Nota Fiscal".
2. Abre o navegador no ERP e prepara a tela: período em "Todo o período" e 100 itens por página.
   O robô NÃO lê a tela: escuta as respostas que a própria página recebe do servidor.
3. Busca os números em lotes de 10 (o ERP aceita vários separados por vírgula e busca EXATO).
   Todo número que não voltou nada no lote é buscado de novo sozinho, do jeito "contém", que ainda
   enxerga números gravados com ano e zeros na frente (202600000088 para a nota 88).
4. Para cada LINHA da planilha, compara número + fornecedor + valor + vencimento com o que o ERP
   devolveu (regras em criterio.py) e sugere uma Ação.
5. Grava saida/resultado_nf.json e a aba "Conferência NF" da planilha.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

import aplicar_na_planilha
import config as cfg
import criterio
import normalizacao
import planilha
from erp_ficticio import garantir


def ler_planilha() -> list:
    """Linhas com Nota Fiscal, lidas pelo próprio Excel (funciona com o arquivo aberto)."""
    with planilha.planilha_aberta() as ws:
        return planilha.ler_linhas_nf(ws)


def pausa(segundos: float = 1.0):
    if cfg.RITMO_S:
        time.sleep(segundos * cfg.RITMO_S)


# ------------------------------------------------------------ navegador ---

def argumentos_janela() -> list:
    """--window-position/--window-size a partir de DEMO_JANELA="x,y,largura,altura"."""
    if not cfg.JANELA_NAVEGADOR:
        return []
    x, y, w, h = [int(v) for v in cfg.JANELA_NAVEGADOR.split(",")]
    return ["--window-position=%d,%d" % (x, y), "--window-size=%d,%d" % (w, h)]


def abrir_navegador(p):
    """Abre o navegador do robô.

    Variáveis de ambiente (usadas só para gravar vídeos de demonstração):
      DEMO_HEADLESS=1       navegador sem janela
      DEMO_GRAVAR=<pasta>   grava em vídeo o que o robô faz na página (arquivo .webm)
      DEMO_T0=<arquivo>     escreve ali a hora em que a página foi criada (para sincronizar o vídeo)
      DEMO_VIEWPORT=LxA     tamanho da página gravada (padrão 720x540)
    """
    headless = os.environ.get("DEMO_HEADLESS") == "1"
    navegador = p.chromium.launch(headless=headless, args=argumentos_janela())
    pasta_video = os.environ.get("DEMO_GRAVAR")
    if pasta_video:
        tamanho = os.environ.get("DEMO_VIEWPORT", "720x540").split("x")
        tela = {"width": int(tamanho[0]), "height": int(tamanho[1])}
        contexto = navegador.new_context(viewport=tela, record_video_dir=pasta_video, record_video_size=tela)
    else:
        contexto = navegador.new_context(no_viewport=True)
    pagina = contexto.new_page()
    if os.environ.get("DEMO_T0"):
        Path(os.environ["DEMO_T0"]).write_text(str(time.time()), encoding="utf-8")
    return navegador, pagina


def salvar_diagnostico(pagina, nome: str):
    """Guarda print + HTML da tela para entender o que mudou."""
    cfg.SAIDA.mkdir(parents=True, exist_ok=True)
    imagem = cfg.SAIDA / ("diagnostico_%s.png" % nome)
    html = cfg.SAIDA / ("diagnostico_%s.html" % nome)
    try:
        pagina.screenshot(path=str(imagem), full_page=True)
        html.write_text(pagina.content(), encoding="utf-8")
        print("\n  Salvei o diagnóstico desta tela em:\n    %s\n    %s\n" % (imagem, html))
    except Exception as e:
        print("  (não consegui salvar o diagnóstico: %s)" % e)


# ------------------------------------------------------ preparação da tela ---

def preparar_tela(pagina):
    """Deixa a tela pronta para as consultas. Devolve o campo de documento."""
    try:
        pagina.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    campo = pagina.get_by_placeholder("Número do documento").first
    campo.wait_for(state="visible", timeout=10000)
    print("  Campo 'Número do documento' localizado.")

    # Sem "Todo o período" as notas fora do mês apareceriam como "não encontradas".
    periodo = pagina.locator("#filtro-periodo")
    if periodo.input_value() != "todo":
        print("  Ajustando o período para 'Todo o período'...")
        periodo.select_option(label="Todo o período")
        pagina.wait_for_timeout(700)
    else:
        print("  Período já está em 'Todo o período'.")
    if periodo.input_value() != "todo":
        salvar_diagnostico(pagina, "periodo")
        raise RuntimeError("Não consegui pôr o período em 'Todo o período'. Parei aqui, "
                           "porque o resultado seria enganoso.")
    pausa(1)

    pagina.locator("#itens-pagina").select_option(str(cfg.ITENS_POR_PAGINA))
    pagina.wait_for_timeout(500)
    print("  Itens por página ajustado para %d." % cfg.ITENS_POR_PAGINA)
    return campo


# ------------------------------------------------------------- consultas ---

def _linhas_da_resposta(dados):
    if isinstance(dados, list):
        return dados
    if isinstance(dados, dict) and isinstance(dados.get("data"), list):
        return dados["data"]
    return []


def _corpo(resposta) -> dict:
    try:
        return json.loads(resposta.request.post_data or "{}")
    except Exception:
        return {}


def _conferir_periodo(corpo_enviado: str):
    """Aborta se o filtro de período não estiver em 'Todo o período'."""
    try:
        enviado = json.loads(corpo_enviado)
    except Exception:
        return
    if enviado.get("start_date") or enviado.get("end_date"):
        raise RuntimeError(
            "O filtro de período não está em 'Todo o período' (o ERP está consultando de %s até %s). "
            "Notas fora desse intervalo apareceriam como 'não encontradas', então parei antes de "
            "gerar resultado errado." % (enviado.get("start_date"), enviado.get("end_date")))


_resumir = normalizacao.resumir


def _proxima_pagina(pagina) -> bool:
    """Clica em 'Próxima página', se existir e estiver habilitada."""
    botao = pagina.get_by_role("button", name="Próxima página")
    try:
        limite = time.time() + 6
        while time.time() < limite:
            if not botao.is_disabled():
                botao.click()
                return True
            pagina.wait_for_timeout(200)
    except Exception:
        pass
    return False


def _consultar_um(pagina, campo, numero, checar_periodo):
    """Digita o número e lê todas as páginas de resultado que conseguir.

    Devolve (lançamentos lidos, total que o ERP diz ter, páginas lidas).
    """
    def resposta_do_numero(resposta):
        return (cfg.ROTA_LISTAGEM in resposta.url
                and str(_corpo(resposta).get("document_number")) == numero)

    with pagina.expect_response(resposta_do_numero, timeout=cfg.TIMEOUT_CONSULTA_S * 1000) as espera:
        campo.fill("")
        campo.fill(numero)
    resposta = espera.value
    if checar_periodo:
        _conferir_periodo(resposta.request.post_data or "")

    conteudo = resposta.json()
    lidas = _linhas_da_resposta(conteudo)
    total = max(int(conteudo.get("total_count") or 0), len(lidas))
    paginas = 1

    while len(lidas) < total and paginas < cfg.MAX_PAGINAS:
        proxima = paginas + 1

        def resposta_da_pagina(r, _n=proxima):
            corpo = _corpo(r)
            return (cfg.ROTA_LISTAGEM in r.url and str(corpo.get("document_number")) == numero
                    and int(corpo.get("page") or 1) == _n)
        try:
            with pagina.expect_response(resposta_da_pagina, timeout=8000) as espera:
                if not _proxima_pagina(pagina):
                    raise PWTimeout("sem botão de próxima página")
            mais = _linhas_da_resposta(espera.value.json())
        except PWTimeout:
            break
        if not mais:
            break
        lidas.extend(mais)
        paginas = proxima
    return [_resumir(r) for r in lidas if isinstance(r, dict)], total, paginas


_registro_de_lote = normalizacao.registro_de_lote


def _registro_individual(pagina, campo, numero, checar_periodo):
    """Busca UM número do jeito "contém": enxerga também ano+zeros e zero à esquerda."""
    try:
        lidas, total, paginas = _consultar_um(pagina, campo, numero, checar_periodo)
        return normalizacao.registro_de_leitura(numero, lidas, total, paginas)
    except PWTimeout:
        return {"numero": numero, "status": criterio.ERRO, "quantidade": 0,
                "erro": "o ERP não respondeu dentro do tempo limite"}


def _resumo(registro) -> str:
    if registro["status"] == criterio.ERRO:
        return "ERRO"
    return "%d igual(is) de %d no sistema%s%s" % (
        registro["quantidade"], registro["total_no_sistema"],
        "" if registro["lidas"] >= registro["total_no_sistema"] else " (li só %d)" % registro["lidas"],
        "  [lote]" if registro.get("via") == "lote" else "")


def consultar_numeros(numeros, exigir=None, lote=None):
    """Roda o navegador e devolve {numero: {...}}.

      1) busca em lotes (rápido);
      2) todo número que não voltou nada no lote -- ou voltou menos notas do que a planilha tem
         linhas (`exigir`) -- é buscado de novo SOZINHO, do jeito "contém", que ainda enxerga
         ano+zeros e zero à esquerda. Sem isso viraria "não encontrada" por engano e a nota
         seria lançada em dobro;
      3) qualquer erro no lote volta ao modo individual para aquele grupo.
    """
    resultados = {}
    exigir = exigir or {}
    tamanho = lote or cfg.LOTE_NUMEROS

    with sync_playwright() as p:
        navegador, pagina = abrir_navegador(p)
        try:
            pagina.goto(cfg.URL_TRANSACOES, wait_until="domcontentloaded")
            pagina.wait_for_timeout(1500)
            campo = preparar_tela(pagina)

            print("\n  Consultando %d números distintos%s...\n" % (
                len(numeros), " (em lotes de %d)" % tamanho if tamanho > 1 else ""))
            periodo_conferido = False
            individuais = []

            if tamanho > 1:
                total_lotes = (len(numeros) + tamanho - 1) // tamanho
                for k in range(0, len(numeros), tamanho):
                    grupo = numeros[k:k + tamanho]
                    resolvidos = 0
                    try:
                        lidas, total, paginas = _consultar_um(
                            pagina, campo, ", ".join(grupo), not periodo_conferido)
                        periodo_conferido = True
                        completo = total <= len(lidas)
                    except Exception as e:                       # lote falhou: modo individual
                        print("   (lote com problema: %s; busco um por um)" % type(e).__name__)
                        completo, lidas, paginas = False, [], 0
                    for n in grupo:
                        reg = _registro_de_lote(n, lidas, paginas) if completo else None
                        if reg and reg["quantidade"] > 0 and reg["quantidade"] >= exigir.get(n, 1):
                            resultados[n] = reg
                            resolvidos += 1
                        else:
                            individuais.append(n)
                    print("  [lote %d/%d] %d números: %d achados de uma vez, %d para busca individual" % (
                        k // tamanho + 1, total_lotes, len(grupo), resolvidos, len(grupo) - resolvidos))
                    pausa(1.2)
                    time.sleep(cfg.PAUSA_ENTRE_CONSULTAS_S)
            else:
                individuais = list(numeros)

            for i, numero in enumerate(individuais, 1):
                registro = _registro_individual(pagina, campo, numero, not periodo_conferido)
                if registro["status"] != criterio.ERRO:
                    periodo_conferido = True
                resultados[numero] = registro
                print("  [%3d/%d] doc %-14s -> %s" % (i, len(individuais), numero, _resumo(registro)))
                pausa(0.8)
                time.sleep(cfg.PAUSA_ENTRE_CONSULTAS_S)
            pausa(1.5)
        finally:
            navegador.close()
    return resultados


# ------------------------------------------------------------------ main ---

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--aba", default="", help="aba de inserção a conferir (senão: a única, ou pergunta)")
    args = ap.parse_args()
    planilha.escolher_aba(args.aba)
    garantir.garantir_servidor()

    print("=" * 70)
    print("  Conferência de Notas Fiscais -- planilha x ERP")
    print("=" * 70)
    print("\nLendo %s (aba '%s')..." % (cfg.PLANILHA.name, cfg.ABA))
    linhas = planilha.com_tentativas(ler_planilha)
    if not linhas:
        print("Nenhuma linha com Tipo Documento = 'Nota Fiscal'. Nada a fazer.")
        return 0

    numeros = sorted({l["numero"] for l in linhas}, key=lambda n: (len(n), n))
    print("  %d linhas com Nota Fiscal, %d números distintos a consultar." % (len(linhas), len(numeros)))

    por_numero = Counter(l["numero"] for l in linhas)
    resultados = consultar_numeros(numeros, exigir=por_numero)

    dados = {"gerado_em": datetime.now().isoformat(timespec="seconds"), "planilha": cfg.PLANILHA.name,
             "aba": cfg.ABA, "linhas": linhas, "consultas": resultados}
    cfg.SAIDA.mkdir(parents=True, exist_ok=True)
    cfg.ARQUIVO_RESULTADO.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\nGravando na aba '%s'..." % cfg.ABA_RESULTADO)
    ja_aberta = planilha.ja_esta_aberta()

    def gravar():
        with planilha.planilha_aberta(gravar=True) as ws:
            return aplicar_na_planilha.gravar_aba_resultado(ws, dados)
    cont, _ = planilha.com_tentativas(gravar)

    print("\n" + "-" * 70)
    print("  Já no sistema   : %d linhas" % cont[criterio.JA_NO_SISTEMA])
    print("  Ajustar valor   : %d linhas" % cont[criterio.AJUSTAR_VALOR])
    print("  Não encontradas : %d linhas" % cont[criterio.NAO_ENCONTRADA])
    print("  Para revisar    : %d linhas" % cont[criterio.REVISAR])
    if cont[criterio.ERRO]:
        print("  Com falha       : %d linhas" % cont[criterio.ERRO])
    print()
    if ja_aberta:
        print("  Resultado na aba '%s' da planilha aberta." % cfg.ABA_RESULTADO)
        print("  Ainda NÃO foi salvo: salve quando quiser.")
    else:
        print("  Resultado salvo na aba '%s'." % cfg.ABA_RESULTADO)
    print("-" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(planilha.executar(main))
