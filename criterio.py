# -*- coding: utf-8 -*-
"""Critério da conferência: cada LINHA da planilha contra o que o sistema devolveu.

O filtro "Número do documento" do sistema é do tipo *contém* (buscar "10" traz
108152), então o número é conferido aqui, um a um, e junto vêm fornecedor, valor
e vencimento (regras de negócio validadas com a equipe).

Comparações
  Número      igual, sem zeros à esquerda. Formato ano+zeros (20260000000123 p/ NF 123)
              conta só como "parecido".
  Valor       igual (até R$ 0,01) | a ajustar (< R$ 1) | próximo (até 10%) | diferente
  Fornecedor  bate (nomes iguais sem acento/pontuação/palavras genéricas) |
              parecido (uma palavra significativa em comum) | diferente
  Soma        linhas da planilha com a mesma NF também são comparadas pela soma
  Parcelas    entre as que servem, vale a de vencimento mais próximo (em aberto
              antes da paga, em caso de empate)

Resultados
  JA_NO_SISTEMA   número + fornecedor bate + valor igual
  AJUSTAR_VALOR   número + fornecedor bate + diferença < R$ 1
  REVISAR         qualquer coisa "mais ou menos", diferença grande, vencimento muito
                  distante numa nota já paga, ou resultados que não deu para ler
  NAO_ENCONTRADA  nenhum número igual, ou número igual com fornecedor E valor diferentes
  ERRO            o sistema não respondeu

A cada linha é sugerida uma Ação (DAR BAIXA / IMPORTAR / IGNORAR / em branco). É
só sugestão: quem decide é você, na aba Conferência NF.

Regras para pagamentos e divisões
  * Lançamento já PAGO nunca recebe "ID a usar": paga em outra data = REVISAR + IMPORTAR
    (é a parcela seguinte); o pago só aparece como referência.
  * "Valor a usar" é SEMPRE o da planilha (soma, se for o caso). Diferença de valor entra na
    baixa como juros (planilha maior) ou desconto (menor); a linha fica REVISAR avisando.
  * Várias linhas -> 1 lançamento em aberto com placas diferentes (todas identificadas):
    REVISAR em todas; a principal fica com o ID e as outras vão para IMPORTAR com a
    descrição do lançamento ("Descrição a usar"). Sem placa em alguma: baixa única pela soma.
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

import config as cfg

JA_NO_SISTEMA = "JA_NO_SISTEMA"
AJUSTAR_VALOR = "AJUSTAR_VALOR"
REVISAR = "REVISAR"
NAO_ENCONTRADA = "NAO_ENCONTRADA"
ERRO = "ERRO"

DAR_BAIXA = "DAR BAIXA"
IMPORTAR = "IMPORTAR"
IGNORAR = "IGNORAR"

EM_ABERTO = "EM ABERTO"
PAGO = "PAGO"
PARCIAL = "PARCIAL"

_NIVEL = {JA_NO_SISTEMA: 0, AJUSTAR_VALOR: 1, REVISAR: 2, NAO_ENCONTRADA: 3}


# ------------------------------------------------------------------ formato ---

def formatar_valor(valor) -> str:
    if valor is None:
        return "?"
    try:
        texto = "{:,.2f}".format(float(valor))
    except (TypeError, ValueError):
        return str(valor)
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_data(iso) -> str:
    if not iso:
        return "?"
    partes = str(iso)[:10].split("-")
    if len(partes) == 3:
        return "%s/%s/%s" % (partes[2], partes[1], partes[0])
    return str(iso)[:10]


# ------------------------------------------------------------------- número ---

def normalizar(numero) -> str:
    """Texto comparável: maiúsculo, sem espaços; sem zeros à esquerda se for só dígitos."""
    texto = re.sub(r"\s+", "", str(numero or "")).upper()
    if texto.isdigit():
        return texto.lstrip("0") or "0"
    return texto


def formato_ano_zeros(numero_sistema, numero) -> bool:
    """20260000000123 para a NF 123: ano (20xx) + zeros + número."""
    alvo = normalizar(numero)
    if not alvo.isdigit():
        return False
    return re.fullmatch(r"20\d{2}0*%s" % re.escape(alvo),
                        re.sub(r"\s+", "", str(numero_sistema or ""))) is not None


def separar(numero: str, transacoes: list, total_no_sistema: int) -> dict:
    """Do que foi lido, separa as de número IGUAL e as de formato ano+zeros."""
    alvo = normalizar(numero)
    iguais = [t for t in transacoes if normalizar(t.get("numero_documento")) == alvo]
    parecidas = [t for t in transacoes if t not in iguais
                 and formato_ano_zeros(t.get("numero_documento"), numero)]
    return {
        "transacoes_iguais": iguais,
        "transacoes_parecidas": parecidas,
        "incompleto": total_no_sistema > len(transacoes),
    }


# --------------------------------------------------------------- fornecedor ---

GENERICAS = {
    "LTDA", "ME", "EPP", "EIRELI", "MEI", "SA", "S", "A", "E", "DE", "DA", "DO",
    "DAS", "DOS", "CIA", "COMERCIO", "COMERCIAL", "INDUSTRIA", "INDUSTRIAL",
    "SERVICO", "SERVICOS", "EMPRESA", "GRUPO", "FILIAL", "MATRIZ",
}


def _palavras(nome) -> list:
    texto = unicodedata.normalize("NFKD", str(nome or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c)).upper()
    palavras = re.sub(r"[^A-Z0-9]+", " ", texto).split()
    return [p for p in palavras if p not in GENERICAS]


def _tokens(nome) -> frozenset:
    return frozenset(_palavras(nome))


def nivel_fornecedor(a, b) -> str:
    """'bate', 'parecido' ou 'diferente'."""
    pa, pb = _palavras(a), _palavras(b)
    if not pa or not pb:
        # só palavras genéricas: compara o texto inteiro
        return "bate" if re.sub(r"\W+", "", str(a or "").upper()) ==             re.sub(r"\W+", "", str(b or "").upper()) and str(a or "").strip() else "diferente"
    # "ignorando espaço": MEGAFERRO = MEGA FERRO, AB SERVICOS = A B SERVICOS
    if frozenset(pa) == frozenset(pb) or "".join(pa) == "".join(pb):
        return "bate"
    # nome abreviado na planilha: o menor está CONTIDO no maior,
    # por palavras (ALFA em ALFA COMERCIO DE PECAS) ou por prefixo sem espaço
    # (MEGAFERRO em MEGA FERRO DISTRIBUIDORA). Mínimo de 4 letras p/ não casar "JB", "MV".
    menor, maior = (pa, pb) if len("".join(pa)) <= len("".join(pb)) else (pb, pa)
    jm, jM = "".join(menor), "".join(maior)
    if len(jm) >= 4 and (frozenset(menor) <= frozenset(maior) or jM.startswith(jm)):
        return "bate"
    if frozenset(pa) & frozenset(pb):
        return "parecido"
    # MEGAFERRO LUBRIFICANTES x MEGA FERRO DISTRIBUIDORA: começam igual (sem espaço), >= 6 letras
    comum = 0
    for x, y in zip(jm, jM):
        if x != y:
            break
        comum += 1
    if comum >= 6:
        return "parecido"
    return "diferente"


# -------------------------------------------------------------------- valor ---

def nivel_valor(planilha, sistema) -> str:
    """'igual', 'ajustar', 'proximo', 'diferente' ou 'sem valor'."""
    if planilha is None or sistema is None:
        return "sem valor"
    diferenca = round(abs(float(planilha) - float(sistema)), 2)
    if diferenca <= cfg.TOLERANCIA_IGUAL:
        return "igual"
    if diferenca < cfg.LIMITE_AJUSTE:
        return "ajustar"
    if diferenca <= cfg.LIMITE_PROXIMO * max(abs(float(sistema)), 0.01):
        return "proximo"
    return "diferente"


# ---------------------------------------------------------------- situação ---

def situacao(t: dict) -> str:
    status = str(t.get("status") or "").upper()
    if status == "PAID":
        return PAGO
    if status == "PARTIAL":
        return PARCIAL
    if status in ("PENDING", "OVERDUE"):
        return EM_ABERTO
    valor = float(t.get("valor") or 0)
    pago = float(t.get("valor_pago") or 0)
    if t.get("pagamento") or (valor and pago >= valor - cfg.TOLERANCIA_IGUAL):
        return PAGO
    if pago > 0:
        return PARCIAL
    return EM_ABERTO


def _dias(iso_a, iso_b):
    def como_data(iso):
        if isinstance(iso, datetime):
            return iso.date()
        if isinstance(iso, date):
            return iso
        try:
            return datetime.strptime(str(iso)[:10], "%Y-%m-%d").date()
        except (TypeError, ValueError):
            return None
    a, b = como_data(iso_a), como_data(iso_b)
    if a is None or b is None:
        return None
    return abs((a - b).days)


# ---------------------------------------------------------------- avaliação ---

def _resultado_do_par(forn: str, valor: str) -> str:
    if forn == "bate":
        return {"igual": JA_NO_SISTEMA, "ajustar": AJUSTAR_VALOR}.get(valor, REVISAR)
    if forn == "parecido":
        return REVISAR
    # fornecedor diferente: só "não encontrada" se o valor também não bate
    return REVISAR if valor in ("igual", "ajustar", "proximo") else NAO_ENCONTRADA


DIAS_MESMO_PAGAMENTO = 3


def _parcela(sit: str, pgto_linha, t: dict) -> int:
    """Quão bem a parcela do sistema corresponde a ESTE pagamento da planilha.

    O vencimento da planilha não identifica a parcela ("03/03" vence dia 17 na
    planilha e dia 27 no sistema), a data de pagamento sim:
      0 = já paga na mesma data (± 3 dias) que a planilha diz  -> é esta
      1 = em aberto                                            -> é a que falta baixar
      2 = paga em OUTRA data                                   -> outra parcela
    """
    if sit == PAGO:
        dias = _dias(pgto_linha, t.get("pagamento"))
        if dias is None:          # sem uma das datas: não dá para dizer
            return 1
        return 0 if dias <= DIAS_MESMO_PAGAMENTO else 2
    return 1


def norm_placa(placa) -> str:
    """'AB12 C ' -> 'AB12C' (maiúsculas, sem espaço nem hífen)."""
    return re.sub(r"[\s\-]+", "", str(placa or "")).upper()


_RE_PARCELA = re.compile(r"(?<!\d)(\d{1,2})\s*/\s*(\d{1,2})(?!\d)")


def parcela_no(texto):
    """(n, total) da primeira 'n/total' plausível na descrição (1 <= n <= total <= 60), ou None."""
    for m in _RE_PARCELA.finditer(str(texto or "")):
        n, total = int(m.group(1)), int(m.group(2))
        if 1 <= n <= total <= 60:
            return n, total
    return None


def _placa_difere(l, t) -> int:
    a, b = norm_placa((l or {}).get("placa")), norm_placa(t.get("placa"))
    return 1 if a and b and a != b else 0


def _parcela_difere(l, t) -> int:
    a, b = parcela_no((l or {}).get("descricao")), parcela_no(t.get("descricao"))
    return 1 if a and b and a[0] != b[0] else 0


def _avaliar(valor_linha, fornecedor_linha, venc_linha, t: dict, parecida=False,
             pgto_linha=None, l=None) -> dict:
    forn = nivel_fornecedor(fornecedor_linha, t.get("fornecedor"))
    val = nivel_valor(valor_linha, t.get("valor"))
    resultado = _resultado_do_par(forn, val)
    motivo = ""
    if parecida and resultado in (JA_NO_SISTEMA, AJUSTAR_VALOR):
        resultado = REVISAR
        motivo = "número em formato ano+zeros (%s)" % t.get("numero_documento")
    dias = _dias(venc_linha, t.get("vencimento"))
    sit = situacao(t)
    parcela = _parcela(sit, pgto_linha, t)
    return {
        "t": t, "resultado": resultado, "forn": forn, "valor": val,
        "situacao": sit, "dias": dias, "motivo": motivo, "parecida": parecida,
        "parcela": parcela,
        "placa_dif": _placa_difere(l, t), "parcela_dif": _parcela_difere(l, t),
    }


def _chave(av: dict):
    """Ordena candidatas: melhor resultado, vencimento mais próximo, em aberto antes."""
    return (_NIVEL[av["resultado"]],
            {"bate": 0, "parecido": 1, "diferente": 2}[av["forn"]],
            {"igual": 0, "ajustar": 1, "proximo": 2, "diferente": 3,
             "sem valor": 4}[av["valor"]],
            av.get("parcela", 1),
            av.get("placa_dif", 0),
            av["dias"] if av["dias"] is not None else 10 ** 6)


def _chave_revisao(av: dict):
    """Ordem das candidatas do passo 3: em aberto antes de paga em outra data, valor antes de nome."""
    return (_NIVEL[av["resultado"]],
            av["forn"] == "diferente",
            av.get("parcela", 1) == 2,
            {"igual": 0, "ajustar": 1, "proximo": 2, "diferente": 3, "sem valor": 4}[av["valor"]],
            {"bate": 0, "parecido": 1, "diferente": 2}[av["forn"]],
            av.get("parcela_dif", 0),
            av.get("placa_dif", 0),
            av["dias"] if av["dias"] is not None else 10 ** 6)


def _rebaixar_se_distante(av: dict, incompleto: bool, soma: bool) -> dict:
    """Nota já paga, soma, ou leitura incompleta com vencimento muito distante: REVISAR.

    Uma linha marcada "já no sistema" sai do arquivo de importação; se o vencimento
    está a meses de distância, pode ser outra parcela, e errar aqui perderia uma
    inserção sem ninguém ver. Baixa de parcela em aberto (atrasada) segue normal.
    """
    if av["resultado"] not in (JA_NO_SISTEMA, AJUSTAR_VALOR):
        return av
    if av.get("parcela") == 2:
        av = dict(av)
        av["resultado"] = REVISAR
        av["motivo"] = ("a parcela que corresponde já foi paga em %s, e a planilha diz "
                        "pagamento em outra data: pode ser outra parcela"
                        % formatar_data(av["t"].get("pagamento")))
        return av
    dias = av["dias"]
    if dias is None or dias <= cfg.DIAS_VENCIMENTO_DISTANTE:
        return av
    if av["situacao"] == EM_ABERTO and not incompleto and not soma:
        return av
    av = dict(av)
    av["resultado"] = REVISAR
    av["motivo"] = ("vencimento do sistema (%s) está %d dias longe do da planilha%s"
                    % (formatar_data(av["t"].get("vencimento")), dias,
                       ", e nem todas as parcelas foram lidas" if incompleto else ""))
    return av


def _acao_sugerida(resultado: str, sit: str) -> str:
    if resultado == NAO_ENCONTRADA:
        return IMPORTAR
    if resultado in (JA_NO_SISTEMA, AJUSTAR_VALOR):
        if sit == EM_ABERTO:
            return DAR_BAIXA
        if sit == PAGO:   # já paga (AJUSTAR: diferença de centavos, nada a fazer)
            return IGNORAR
    return ""


# ----------------------------------------------------- classificação em lote ---

def classificar_linhas(linhas: list, consultas: dict) -> dict:
    """{nº da linha da planilha: resultado} para todas as linhas com Nota Fiscal.

    Cada resultado traz: resultado, acao, situacao, id, valor_sistema,
    valor_a_usar, motivo, candidata (dict da transação), avaliacao (fornecedor,
    valor, dias), soma (bool), grupo (linhas envolvidas), observacao.
    """
    saida = {}
    por_numero = {}
    for l in linhas:
        por_numero.setdefault(normalizar(l["numero"]), []).append(l)

    for chave, grupo in por_numero.items():
        consulta = consultas.get(grupo[0]["numero"], {})
        saida.update(_classificar_grupo(grupo, consulta))
    return saida


def _sem_consulta(l, consulta) -> dict:
    return {"resultado": ERRO, "acao": "", "situacao": "", "id": None,
            "valor_sistema": None, "valor_a_usar": None, "soma": False,
            "grupo": [l["linha"]], "candidata": None, "avaliacao": None,
            "pagavel": False, "preencher": False, "dividir": False,
            "descricao_a_usar": None,
            "motivo": consulta.get("erro") or "sem resposta do sistema",
            "observacao": "Erro na consulta: %s" % (
                consulta.get("erro") or "sem resposta do sistema")}


def _classificar_grupo(grupo: list, consulta: dict) -> dict:
    if not consulta or consulta.get("status") == ERRO or "erro" in consulta:
        return {l["linha"]: _sem_consulta(l, consulta) for l in grupo}

    iguais = list(consulta.get("transacoes_iguais") or [])
    parecidas = list(consulta.get("transacoes_parecidas") or [])
    total = consulta.get("total_no_sistema") or 0
    lidas = consulta.get("lidas")
    if lidas is None:      # resultado antigo: deduz do motivo gravado
        incompleto = "só foi possível ler" in (consulta.get("motivo") or "")
    else:
        incompleto = total > lidas

    resultados = {}
    usadas = set()          # ids de transações já ocupadas por outra linha

    def livre(t):
        return id(t) not in usadas

    # 1) linha a linha, atribuição gulosa das melhores combinações JÁ/AJUSTAR
    pares = []
    for l in grupo:
        for t in iguais:
            av = _rebaixar_se_distante(
                _avaliar(l["valor"], l["fornecedor"], l["vencimento"], t,
                         pgto_linha=l.get("data_pagamento"), l=l),
                incompleto, soma=False)
            if av["resultado"] in (JA_NO_SISTEMA, AJUSTAR_VALOR):
                pares.append((_chave(av), l["linha"], av))
    pares.sort(key=lambda p: (p[0], p[1]))
    atribuidas = {}
    for _, linha, av in pares:
        if linha in atribuidas or not livre(av["t"]):
            continue
        atribuidas[linha] = av
        usadas.add(id(av["t"]))

    # 2) soma das linhas que sobraram (mesma NF, mesmo fornecedor)
    sobra = [l for l in grupo if l["linha"] not in atribuidas]
    soma_ok = None
    if len(sobra) > 1 and len({nivel_fornecedor(sobra[0]["fornecedor"], l["fornecedor"])
                                for l in sobra}) == 1 and all(
            nivel_fornecedor(sobra[0]["fornecedor"], l["fornecedor"]) == "bate"
            for l in sobra):
        total_valor = round(sum(float(l["valor"] or 0) for l in sobra), 2)
        venc_ref = min((l["vencimento"] for l in sobra if l["vencimento"]), default=None)
        pgto_ref = min((l.get("data_pagamento") for l in sobra if l.get("data_pagamento")),
                       default=None)
        candidatas = []
        for t in iguais:
            if not livre(t):
                continue
            bruto = _avaliar(total_valor, sobra[0]["fornecedor"], venc_ref, t,
                             pgto_linha=pgto_ref)
            if bruto["resultado"] not in (JA_NO_SISTEMA, AJUSTAR_VALOR):
                continue
            # a soma bate; se o vencimento estiver longe, vira REVISAR (mas segue
            # sendo dito que é a soma, para você não achar que são linhas soltas)
            candidatas.append(_rebaixar_se_distante(bruto, incompleto, soma=True))
        if candidatas:
            candidatas.sort(key=_chave)
            melhor = candidatas[0]
            modo, nota = None, ""
            if melhor["resultado"] != REVISAR:
                usadas.add(id(melhor["t"]))
                placas = [norm_placa(m.get("placa")) for m in sobra]
                if melhor["situacao"] == EM_ABERTO and all(placas) and len(set(placas)) > 1:
                    modo = "placa"
                elif any(placas) and not all(placas):
                    sem = [str(m["linha"]) for m in sobra if not norm_placa(m.get("placa"))]
                    nota = "linha%s %s sem placa: baixa única pela soma" % (
                        "s" if len(sem) > 1 else "", ", ".join(sem))
            soma_ok = (melhor, total_valor, sobra, modo, nota)

    # 3) o restante: melhor candidata livre, ou NÃO ENCONTRADA / REVISAR
    for l in grupo:
        linha = l["linha"]
        if linha in atribuidas:
            resultados[linha] = _montar(l, atribuidas[linha], incompleto, [linha])
            continue
        if soma_ok and l in soma_ok[2]:
            av, total_valor, membros, modo, nota = soma_ok
            if modo == "placa":
                resultados[linha] = _montar_dividido(l, av, membros)
                continue
            primeira = min(m["linha"] for m in membros)
            resultados[linha] = _montar(
                l, av, incompleto, [m["linha"] for m in membros], soma=total_valor,
                principal=(linha == primeira), nota=nota)
            continue

        pg = l.get("data_pagamento")
        candidatas = [_avaliar(l["valor"], l["fornecedor"], l["vencimento"], t, pgto_linha=pg, l=l)
                      for t in iguais if livre(t)]
        candidatas += [_avaliar(l["valor"], l["fornecedor"], l["vencimento"], t,
                                parecida=True, pgto_linha=pg, l=l) for t in parecidas if livre(t)]
        candidatas = [_rebaixar_se_distante(av, incompleto, soma=False)
                      for av in candidatas]
        # candidatas de ano+zeros só valem quando apontam algo (nunca NÃO ENCONTRADA)
        candidatas = [av for av in candidatas
                      if not (av["parecida"] and av["resultado"] == NAO_ENCONTRADA)]
        candidatas.sort(key=_chave_revisao)

        if candidatas:
            resultados[linha] = _montar(l, candidatas[0], incompleto, [linha])
        else:
            sem = _resultado_vazio(l, iguais, incompleto, total, lidas)
            resultados[linha] = sem
    return resultados


def _resultado_vazio(l, iguais, incompleto, total, lidas) -> dict:
    if incompleto:
        resultado = REVISAR
        motivo = ("o sistema tem %d resultados para esse trecho e só foi possível "
                  "ler %s" % (total, lidas if lidas is not None else "parte deles"))
    elif iguais:
        resultado = NAO_ENCONTRADA
        motivo = ("todas as transações com esse número já correspondem a outras "
                  "linhas da planilha")
    else:
        resultado, motivo = NAO_ENCONTRADA, ""
    return {"resultado": resultado, "acao": IMPORTAR if resultado == NAO_ENCONTRADA else "",
            "situacao": "", "id": None, "valor_sistema": None, "valor_a_usar": None,
            "soma": False, "grupo": [l["linha"]], "candidata": None,
            "pagavel": False, "preencher": False, "dividir": False,
            "descricao_a_usar": None,
            "avaliacao": None, "motivo": motivo, "observacao": motivo}


def _lista_placas(membros) -> str:
    vistas, saida = set(), []
    for m in membros:
        p = str(m.get("placa") or "").strip()
        if norm_placa(p) not in vistas:
            vistas.add(norm_placa(p))
            saida.append(p)
    return " / ".join(saida)


def _montar_dividido(l, av, membros) -> dict:
    """Várias linhas, placas diferentes, 1 lançamento em aberto: REVISAR em todas.

    Principal = a linha com a placa do lançamento (senão a de menor número): fica com o ID e o
    valor DELA, Ação em branco (você marca DAR BAIXA depois de editar o lançamento no sistema).
    As demais vão para IMPORTAR com a descrição do lançamento ("Descrição a usar").
    """
    t = av["t"]
    placa_sis = norm_placa(t.get("placa"))
    principal = next((m for m in sorted(membros, key=lambda m: m["linha"])
                      if placa_sis and norm_placa(m.get("placa")) == placa_sis),
                     min(membros, key=lambda m: m["linha"]))
    outras = sorted(m["linha"] for m in membros if m is not principal)
    base = {"resultado": REVISAR, "situacao": EM_ABERTO, "valor_sistema": t.get("valor"),
            "soma": False, "grupo": [m["linha"] for m in membros], "candidata": t,
            "pagavel": True, "dividir": True,
            "avaliacao": {"fornecedor": av["forn"], "valor": av["valor"], "dias": av["dias"]}}
    if l is principal:
        obs = ("%d linhas, placas %s → 1 lançamento (ID %s, %s, %s). 1) Edite o lançamento no "
               "sistema: placa %s e valor %s. 2) Marque DAR BAIXA nesta linha. 3) As linhas %s vão "
               "por importação, com a mesma descrição do lançamento." % (
                   len(membros), _lista_placas(membros), t.get("id"), formatar_valor(t.get("valor")),
                   ("placa " + str(t["placa"])) if t.get("placa") else "sem placa",
                   str(l.get("placa") or "").strip(), formatar_valor(l["valor"]),
                   ", ".join(str(n) for n in outras)))
        return dict(base, acao="", id=t.get("id"), valor_a_usar=l["valor"], preencher=True,
                    descricao_a_usar=None, motivo=obs, observacao=obs)
    obs = ("importar com a placa %s; o lançamento ID %s fica com a linha %d (editar no sistema)" % (
        str(l.get("placa") or "").strip(), t.get("id"), principal["linha"]))
    return dict(base, acao=IMPORTAR, id=None, valor_a_usar=None, preencher=False,
                descricao_a_usar=t.get("descricao") or None, motivo=obs, observacao=obs)


def _montar(l, av, incompleto, grupo, soma=None, principal=True, nota="") -> dict:
    t, resultado, sit = av["t"], av["resultado"], av["situacao"]
    motivo = av["motivo"]

    if incompleto and resultado in (REVISAR, NAO_ENCONTRADA) and not motivo:
        resultado = REVISAR
        motivo = "o sistema tem mais resultados do que foi possível ler"
    if av["parecida"] and not motivo and resultado == REVISAR:
        motivo = "número em formato ano+zeros (%s)" % t.get("numero_documento")

    avaliacao = {"fornecedor": av["forn"], "valor": av["valor"], "dias": av["dias"]}
    if resultado == NAO_ENCONTRADA:
        # há o número, mas de outro fornecedor E com outro valor: não é esta linha
        obs = "o sistema só tem esse número em outra transação (%s, %s, ID %s)" % (
            t.get("fornecedor") or "?", formatar_valor(t.get("valor")), t.get("id") or "?")
        return {"resultado": NAO_ENCONTRADA, "acao": IMPORTAR, "situacao": "",
                "id": None, "valor_sistema": None, "valor_a_usar": None,
                "soma": False, "grupo": grupo, "candidata": None,
                "pagavel": False, "preencher": False, "dividir": False,
                "descricao_a_usar": None,
                "avaliacao": avaliacao, "motivo": obs, "observacao": obs}

    if resultado == REVISAR and not motivo:
        motivo = "fornecedor %s, valor %s (%s no sistema)" % (
            av["forn"], av["valor"], formatar_valor(t.get("valor")))
        if av["forn"] != "bate":
            motivo += ", fornecedor no sistema: %s" % (t.get("fornecedor") or "?")

    # Já paga em OUTRA data que a da planilha: não há o que baixar; provavelmente é a parcela
    # seguinte, que ainda não existe no sistema -> IMPORTAR. O lançamento pago fica só de referência.
    if sit == PAGO and av.get("parcela") == 2 and av["forn"] in ("bate", "parecido"):
        obs = ("a parcela do sistema (ID %s, %s) foi paga em %s, outra data que a da planilha: "
               "nada a baixar. Provável parcela seguinte → importar. Confira." % (
                   t.get("id"), formatar_valor(t.get("valor")), formatar_data(t.get("pagamento"))))
        if soma is not None:
            obs += " (soma das %d linhas, %s)" % (len(grupo), formatar_valor(soma))
        return {"resultado": REVISAR, "acao": IMPORTAR, "situacao": sit, "id": None,
                "valor_sistema": t.get("valor"), "valor_a_usar": None, "soma": soma is not None,
                "grupo": grupo, "candidata": t, "pagavel": False, "preencher": False,
                "dividir": False, "descricao_a_usar": None, "referencia_paga": True,
                "avaliacao": avaliacao, "motivo": obs, "observacao": obs}

    pagavel = sit == EM_ABERTO
    acao = _acao_sugerida(resultado, sit)
    if resultado == AJUSTAR_VALOR and sit == PAGO and not motivo:
        motivo = "já paga com %s (diferença de centavos): nada a baixar" % \
            formatar_valor(t.get("valor"))
    if sit == PAGO and resultado == REVISAR:
        motivo += "; lançamento já pago no sistema: nada a baixar. Escolha IGNORAR (se é este) ou IMPORTAR"
    if sit == PARCIAL:
        motivo = (motivo + "; " if motivo else "") + \
            "pagamento parcial no sistema (saldo %s): confira" % formatar_valor(t.get("saldo"))

    # "Valor a usar" é SEMPRE o da planilha; a diferença entra na baixa como juros ou desconto
    valor_usar = None
    soma_secundaria = soma is not None and not principal
    preencher = pagavel and acao != IGNORAR and not soma_secundaria
    if preencher:
        valor_usar = soma if soma is not None else l["valor"]
        base_sis = t.get("saldo") if t.get("saldo") is not None else t.get("valor")
        if valor_usar is not None and base_sis is not None:
            dif = round(float(valor_usar) - float(base_sis), 2)
            if abs(dif) >= cfg.LIMITE_AJUSTE:
                motivo = (motivo + "; " if motivo else "") + (
                    "valores diferentes: planilha %s × sistema %s. Na baixa, %s entram como %s" % (
                        formatar_valor(valor_usar), formatar_valor(base_sis),
                        formatar_valor(abs(dif)), "JUROS" if dif > 0 else "DESCONTO"))
    observacao = motivo

    if soma is not None:
        observacao = "soma das %d linhas (%s) corresponde a esta transação" % (
            len(grupo), formatar_valor(soma))
        if not principal and resultado in (JA_NO_SISTEMA, AJUSTAR_VALOR):
            acao, valor_usar = IGNORAR, None
            observacao += "; a baixa é feita na linha %d" % min(grupo)
        elif motivo:
            observacao = motivo + "; " + observacao
        if nota:
            observacao += "; " + nota
        if soma_secundaria and resultado == REVISAR:
            observacao += "; a baixa é feita na linha %d" % min(grupo)

    pn_l, pn_t = parcela_no(l.get("descricao")), parcela_no(t.get("descricao"))
    if pn_l and pn_t and pn_l[0] != pn_t[0] and observacao:
        observacao += " (parcela na planilha: %d/%d · no sistema: %d/%d)" % (pn_l + pn_t)

    return {
        "resultado": resultado, "acao": acao, "situacao": sit,
        "id": t.get("id"), "valor_sistema": t.get("valor"),
        "valor_a_usar": valor_usar, "soma": soma is not None, "grupo": grupo,
        "candidata": t, "pagavel": pagavel, "preencher": preencher, "dividir": False,
        "descricao_a_usar": None, "avaliacao": avaliacao,
        "motivo": motivo, "observacao": observacao,
    }
