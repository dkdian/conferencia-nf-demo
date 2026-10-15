# -*- coding: utf-8 -*-
"""
Acesso à planilha pelo próprio Excel (COM).

Por que não openpyxl: quando o arquivo está aberto no Excel, o Excel o trava e
nenhum outro processo consegue abri-lo -- openpyxl morre com PermissionError.
Falando com o Excel, funciona com o arquivo aberto ou fechado, enxerga o que
está na tela (inclusive alterações ainda não salvas) e, na hora de gravar, não
perde macros, tabelas, validações nem o painel de suplemento.
"""

from __future__ import annotations

import contextlib
from datetime import datetime, date

import pythoncom
import win32com.client as win32

import config as cfg

XL_UP = -4162  # xlUp


# --------------------------------------------------------- abrir / fechar ---

@contextlib.contextmanager
def com_workbook(gravar: bool = False):
    """Devolve a pasta de trabalho (workbook) da planilha.

    Planilha JÁ ABERTA no Excel do usuário: usa a janela dele e NÃO mexe em nada
    que não seja o conteúdo pedido -- não muda visibilidade, não salva (pode haver
    edições em andamento) e não fecha. Quem chama avisa que falta salvar.

    Planilha FECHADA: abre numa instância do Excel só nossa e invisível
    (DispatchEx), que nunca encosta nas janelas do usuário; se `gravar`, salva ao
    final. A instância é sempre encerrada. Em 28/09 uma conexão à instância
    compartilhada travou o Excel do usuário -- por isso nunca usamos
    Dispatch("Excel.Application") aqui.
    """
    pythoncom.CoInitialize()
    wb = achar_workbook_aberto()
    ja_estava_aberta = wb is not None
    excel = None
    try:
        if not ja_estava_aberta:
            if not cfg.PLANILHA.exists():
                raise SystemExit("Planilha não encontrada em:\n  %s" % cfg.PLANILHA)
            excel = win32.DispatchEx("Excel.Application")   # instância nova, só nossa
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: nada de macro ao abrir
            wb = excel.Workbooks.Open(str(cfg.PLANILHA), UpdateLinks=0)
        yield wb
        if gravar and not ja_estava_aberta:
            wb.Save()
    finally:
        if excel is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            excel.Quit()
        pythoncom.CoUninitialize()


@contextlib.contextmanager
def planilha_aberta(aba: str = "", gravar: bool = False):
    """Devolve a worksheet pedida (padrão: a aba de trabalho ativa, cfg.ABA)."""
    with com_workbook(gravar) as wb:
        yield wb.Worksheets(aba or cfg.ABA)


def com_tentativas(funcao, tentativas=6, espera_s=5):
    """O Excel recusa chamadas enquanto uma célula está em edição: tenta de novo."""
    import time
    import pywintypes
    for n in range(1, tentativas + 1):
        try:
            return funcao()
        except pywintypes.com_error as e:
            if n == tentativas or e.hresult not in (-2147418111, -2147417846):
                raise
            print("  Excel ocupado (célula em edição?). Tentando de novo em %ds..." % espera_s)
            time.sleep(espera_s)


class ExcelOcupado(RuntimeError):
    """A planilha está aberta no seu Excel, mas ele não respondeu (célula em edição, janela aberta)."""


def _monikers_da_planilha():
    """[(rot, moniker)] dos objetos registrados no Windows com o nome da planilha."""
    import urllib.parse
    alvo = cfg.PLANILHA.name.lower()
    rot = pythoncom.GetRunningObjectTable()
    contexto = pythoncom.CreateBindCtx(0)
    achados = []
    for moniker in rot.EnumRunning():
        try:
            nome = moniker.GetDisplayName(contexto, None)
        except Exception:
            continue
        arquivo = urllib.parse.unquote(nome).replace("\\", "/").split("/")[-1].lower()
        if arquivo == alvo:
            achados.append((rot, moniker))
    return achados


def achar_workbook_aberto(tentativas: int = 8, espera_s: float = 2.0):
    """Devolve a planilha se ela estiver aberta em QUALQUER instância do Excel.

    Procura na tabela de objetos em execução do Windows em vez de perguntar a
    "o" Excel: pode haver mais de uma instância (uma invisível, por exemplo), e
    a pergunta seria respondida pela errada. Compara pelo nome do arquivo, pois
    em pasta sincronizada na nuvem o Excel pode registrar o caminho como URL.

    Se a planilha ESTÁ registrada como aberta mas o Excel não responde (célula em edição,
    janela de diálogo), tenta de novo por alguns segundos e, se não adiantar, levanta
    ExcelOcupado. NUNCA devolve None nesse caso: quem recebesse None abriria uma segunda
    cópia do arquivo por cima da sua.
    """
    import time
    for n in range(tentativas):
        achados = _monikers_da_planilha()
        if not achados:
            return None
        for rot, moniker in achados:
            try:
                objeto = rot.GetObject(moniker)
                wb = win32.Dispatch(objeto.QueryInterface(pythoncom.IID_IDispatch))
                wb.Worksheets  # confirma que é mesmo uma pasta de trabalho
                return wb
            except Exception:
                continue
        if n < tentativas - 1:
            time.sleep(espera_s)
    raise ExcelOcupado(
        "O Excel está ocupado (uma célula em edição ou uma janela aberta?). Saia da edição "
        "(Esc), feche as janelas de diálogo do Excel e clique de novo. NADA foi feito.")


def executar(funcao) -> int:
    """Roda main() mostrando o ExcelOcupado como mensagem, não como traceback."""
    try:
        return funcao()
    except ExcelOcupado as e:
        print("\n  " + str(e))
        return 1


def ja_esta_aberta() -> bool:
    """A planilha está aberta no Excel agora? (só olha o registro do Windows; não conecta)"""
    pythoncom.CoInitialize()
    try:
        return bool(_monikers_da_planilha())
    finally:
        pythoncom.CoUninitialize()


# ------------------------------------------------------------- conversões ---

def _texto(valor) -> str:
    return "" if valor is None else str(valor).strip()


def _numero_documento(valor) -> str:
    """Normaliza o número do documento para texto, sem o '.0' que o Excel traz."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, int):
        return str(valor)
    return str(valor).strip()


def _valor_float(valor):
    """Aceita 1234.5, '1234,5' e '1.234,50'."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("R$", "").strip()
    if not texto:
        return None
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def formatos(ws) -> dict:
    """Códigos de formato de célula que ESTE Excel entende.

    O Excel em português usa códigos locais em NumberFormat ("dd/mm/aaaa",
    "#.##0,00"); com "dd/mm/yyyy" a célula mostraria "12/09/yyyy" (visto em
    ao testar). Os códigos vêm de Application.International (índice do Excel - 1).
    """
    app = ws.Application
    intl = app.International
    dia, mes, ano = intl[20], intl[19], intl[18]      # xlDayCode, xlMonthCode, xlYearCode
    hora, minuto = intl[21], intl[22]                 # xlHourCode, xlMinuteCode
    sep = intl[16]                                    # xlDateSeparator
    data = "%s%s%s%s%s" % (dia * 2, sep, mes * 2, sep, ano * 4)
    numero = "#%s##0%s00" % (app.ThousandsSeparator, app.DecimalSeparator)
    return {"data": data, "data_hora": "%s %s:%s" % (data, hora * 2, minuto * 2),
            "numero": numero}


def serial_excel(valor):
    """Data -> número de série do Excel (sem hora nem fuso).

    Passar datetime pelo COM grava a data com 3h de diferença (fuso do Brasil);
    o serial inteiro evita isso, e o formato de data da célula cuida da exibição.
    """
    if isinstance(valor, datetime) or hasattr(valor, "year"):
        base = date(1899, 12, 30)
        return float((date(valor.year, valor.month, valor.day) - base).days)
    return valor


def _data_iso(valor):
    if valor is None:
        return None
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")
    if hasattr(valor, "strftime"):          # pywintypes.datetime
        return valor.strftime("%Y-%m-%d")
    return _texto(valor) or None


# ---------------------------------------------------------------- leitura ---

def ultima_linha(ws) -> int:
    return int(ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row)


def ler_linhas_nf(ws) -> list:
    """Linhas da aba cujo Tipo Documento é 'Nota Fiscal' e que têm número."""
    fim = ultima_linha(ws)
    primeira = cfg.LINHA_CABECALHO + 1
    if fim < primeira:
        return []

    # Uma leitura só do bloco inteiro: COM célula a célula é lento demais.
    bloco = ws.Range(
        ws.Cells(primeira, 1),
        ws.Cells(fim, cfg.ULTIMA_COL_INSERCAO),
    ).Value
    if bloco and not isinstance(bloco[0], (tuple, list)):
        bloco = (bloco,)

    linhas = []
    for deslocamento, valores in enumerate(bloco or ()):
        numero_linha = primeira + deslocamento

        def coluna(indice, _v=valores):
            return _v[indice - 1] if indice - 1 < len(_v) else None

        tipo_doc = _texto(coluna(cfg.COL_TIPO_DOCUMENTO))
        if cfg.TIPO_DOCUMENTO_ALVO not in tipo_doc.lower():
            continue

        numero = _numero_documento(coluna(cfg.COL_NUMERO_DOCUMENTO))
        if not numero:
            print("  ! linha %d: 'Nota Fiscal' sem número do documento -- pulada"
                  % numero_linha)
            continue

        linhas.append({
            "linha": numero_linha,
            "numero": numero,
            "valor": _valor_float(coluna(cfg.COL_VALOR)),
            "vencimento": _data_iso(coluna(cfg.COL_VENCIMENTO)),
            "fornecedor": _texto(coluna(cfg.COL_FORNECEDOR)),
            "descricao": _texto(coluna(cfg.COL_DESCRICAO)),
            "data_pagamento": _data_iso(coluna(cfg.COL_DATA_PAGAMENTO)),
            "conta": _texto(coluna(cfg.COL_CONTA)),
            "forma": _texto(coluna(cfg.COL_FORMA)),
            "categoria": _texto(coluna(cfg.COL_CATEGORIA)),
            "centro_custo": _texto(coluna(cfg.COL_CENTRO_CUSTO)),
            "placa": _texto(coluna(cfg.COL_PLACA)),
        })

    return linhas


# ------------------------------------------------- Inserção: ler / backup ---

def _naive(valor):
    """datetime do COM (com fuso) -> datetime simples, que o openpyxl aceita."""
    if hasattr(valor, "year") and hasattr(valor, "hour"):
        return datetime(valor.year, valor.month, valor.day,
                        valor.hour, valor.minute, valor.second)
    return valor


def ler_insercao(ws):
    """(cabeçalho, linhas) da aba Inserção, colunas A..Q.

    Cada linha é {"linha": nº, "valores": [17 valores]}. Linhas totalmente em
    branco são descartadas.
    """
    fim = ultima_linha(ws)
    largura = cfg.ULTIMA_COL_INSERCAO
    cab = ws.Range(ws.Cells(cfg.LINHA_CABECALHO, 1),
                   ws.Cells(cfg.LINHA_CABECALHO, largura)).Value
    cabecalho = [_naive(v) for v in (cab[0] if cab else ())]
    primeira = cfg.LINHA_CABECALHO + 1
    linhas = []
    if fim >= primeira:
        bloco = ws.Range(ws.Cells(primeira, 1), ws.Cells(fim, largura)).Value
        if bloco and not isinstance(bloco[0], (tuple, list)):
            bloco = (bloco,)
        for desloc, valores in enumerate(bloco or ()):
            vals = [_naive(v) for v in valores]
            if any(v not in (None, "") for v in vals):
                linhas.append({"linha": primeira + desloc, "valores": vals})
    return cabecalho, linhas


def backup_aba(ws, rotulo: str):
    """Guarda os valores da aba em backup/<rotulo> <data-hora>.xlsx."""
    import openpyxl
    cfg.BACKUP.mkdir(parents=True, exist_ok=True)
    usado = ws.UsedRange.Value
    if usado is not None and not isinstance(usado, (tuple, list)):
        usado = ((usado,),)
    elif usado and not isinstance(usado[0], (tuple, list)):
        usado = (usado,)
    wb = openpyxl.Workbook()
    dest = wb.active
    dest.title = ws.Name[:31]
    for linha in usado or ():
        dest.append([_naive(v) for v in linha])
    caminho = cfg.BACKUP / ("%s %s.xlsx" % (rotulo, datetime.now().strftime("%Y%m%d_%H%M%S")))
    wb.save(caminho)
    return caminho


# -------------------------------------------- Inserção: acrescentar linhas ---

COL_COMPETENCIA = 5   # E - texto "mm/aaaa": sem o apóstrofo o Excel viraria data


def _para_celula(indice_col: int, valor):
    """Converte o valor do ETL no que a célula recebe, como a macro original."""
    if valor is None or valor == "":
        return None
    if indice_col == cfg.COL_VALOR and isinstance(valor, str):
        numero = _valor_float(valor)
        return numero if numero is not None else valor
    if indice_col == COL_COMPETENCIA and isinstance(valor, str):
        return "'" + valor
    if hasattr(valor, "year"):
        return serial_excel(valor)
    return valor


def acrescentar_insercao(ws, linhas_novas) -> int:
    """Escreve `linhas_novas` (listas de 17 valores) logo abaixo da última linha
    da aba Inserção. Não altera nada do que já existe. Devolve a 1ª linha usada."""
    if not linhas_novas:
        return ultima_linha(ws) + 1
    primeira = ultima_linha(ws) + 1
    largura = cfg.ULTIMA_COL_INSERCAO
    matriz = tuple(
        tuple(_para_celula(c, v) for c, v in enumerate(list(l)[:largura], 1))
        for l in linhas_novas)
    destino = ws.Range(ws.Cells(primeira, 1),
                       ws.Cells(primeira + len(matriz) - 1, largura))
    destino.Value = matriz
    fmt_data = formatos(ws)["data"]
    for coluna in (cfg.COL_VENCIMENTO, cfg.COL_DATA_PAGAMENTO):
        ws.Range(ws.Cells(primeira, coluna),
                 ws.Cells(primeira + len(matriz) - 1, coluna)).NumberFormat = fmt_data
    return primeira


# ------------------------------------------------ abas "Inserção", "Inserção - dd.mm" ---

def _sem_acento(t) -> str:
    import unicodedata
    t = unicodedata.normalize("NFKD", str(t or ""))
    return "".join(c for c in t if not unicodedata.combining(c)).upper().strip()


def eh_aba_insercao(nome: str) -> bool:
    n = _sem_acento(nome)
    return (n.startswith(_sem_acento(cfg.PREFIXO_ABA)) and "CONFER" not in n
            and not n.endswith(_sem_acento(cfg.SUFIXO_COMPLETA)))


def abas_insercao(wb) -> list:
    return [w.Name for w in wb.Worksheets if eh_aba_insercao(w.Name)]


def escolher_aba(nome: str = "", interativo: bool = True) -> str:
    """Define cfg.ABA: a aba que os botões vão usar.

    Com `nome`, usa essa. Senão: só uma aba "Inserção..." -> ela; várias -> pergunta.
    """
    with com_workbook() as wb:
        abas = abas_insercao(wb)
    if nome:
        if nome not in abas:
            raise SystemExit("A aba '%s' não existe. Abas disponíveis: %s" % (nome, ", ".join(abas) or "nenhuma"))
        cfg.ABA = nome
        return nome
    if not abas:
        raise SystemExit("Não há nenhuma aba 'Inserção' na planilha. Use o botão Buscar período.")
    if len(abas) == 1:
        cfg.ABA = abas[0]
        return cfg.ABA
    if not interativo:
        raise SystemExit("Há várias abas de inserção (%s): informe qual usar." % ", ".join(abas))
    print("\n  Há mais de uma aba de inserção na planilha. Qual usar?")
    for i, n in enumerate(abas, 1):
        print("    %d) %s" % (i, n))
    while True:
        r = input("\n  Número: ").strip()
        if r.isdigit() and 1 <= int(r) <= len(abas):
            cfg.ABA = abas[int(r) - 1]
            print("  Usando a aba '%s'." % cfg.ABA)
            return cfg.ABA
        print("  Digite um número da lista.")


def nome_para_aba_nova(wb, hoje=None) -> str:
    """'Inserção - dd.mm' (data em que rodou); se já existir, ' (2)', ' (3)'..."""
    hoje = hoje or datetime.now()
    base = "%s - %s" % (cfg.PREFIXO_ABA, hoje.strftime("%d.%m"))
    existentes = {w.Name.lower() for w in wb.Worksheets}
    nome, n = base, 1
    while nome.lower() in existentes:
        n += 1
        nome = "%s (%d)" % (base, n)
    return nome


# Colunas com lista suspensa / conferência de cadastro (nome definido no workbook)
LISTAS_COLUNAS = {10: "lst_Categoria", 11: "lst_ContaBancaria", 12: "lst_FormaPagamento",
                  16: "lst_CentroCusto", 17: "lst_Placa"}


def criar_aba_insercao(wb, cabecalho, linhas, hoje=None, nome=None, depois_de=None,
                       cadastro=True, extra_cabecalho=None):
    """Cria a aba 'Inserção - dd.mm' com o cabeçalho e as linhas dadas. Devolve a worksheet.

    Nunca mexe nas abas existentes. Formatação e listas ficam iguais às da Inserção
    (cabeçalho, larguras, autofiltro, datas, lista suspensa + vermelho para valor que não
    está no cadastro), mas as listas apontam para a aba Lista DESTA planilha (a Inserção
    antiga tinha nomes que apontavam para um arquivo em Downloads).
    """
    modelo = None
    for w in wb.Worksheets:
        if eh_aba_insercao(w.Name):
            modelo = w
            break
    nome = nome or nome_para_aba_nova(wb, hoje)
    if depois_de is not None:
        nova = wb.Worksheets.Add(None, depois_de)          # (Before, After)
    else:
        nova = wb.Worksheets.Add(Before=wb.Worksheets(1))
    nova.Name = nome
    largura = cfg.ULTIMA_COL_INSERCAO
    fmt = formatos(nova)

    # cabeçalho e larguras
    nova.Range(nova.Cells(1, 1), nova.Cells(1, largura)).Value = tuple(
        tuple(c if c not in (None,) else "" for c in cabecalho[:largura]) for _ in range(1))
    for c in range(1, largura + 1):
        try:
            nova.Columns(c).ColumnWidth = modelo.Columns(c).ColumnWidth if modelo else 20
        except Exception:
            nova.Columns(c).ColumnWidth = 20
    if modelo is not None:
        nova.Rows(1).RowHeight = modelo.Rows(1).RowHeight

    # dados
    fim = 1
    if linhas:
        fim = acrescentar_insercao(nova, linhas) + len(linhas) - 1
    ultima_fmt = max(fim, 1000)

    if extra_cabecalho:
        for k, titulo in enumerate(extra_cabecalho, largura + 1):
            nova.Cells(1, k).Value = titulo
            nova.Columns(k).ColumnWidth = 46
    # cadastro: lista suspensa + vermelho quando o valor não está na lista
    for col, nome_lista in (LISTAS_COLUNAS.items() if cadastro else ()):
        faixa = nova.Range(nova.Cells(2, col), nova.Cells(ultima_fmt, col))
        try:
            v = faixa.Validation
            v.Delete()
            v.Add(3, 2, 1, "=" + nome_lista)     # lista; alerta de aviso (não bloqueia)
            v.IgnoreBlank = True
            v.InCellDropdown = True
            v.ShowError = False
        except Exception:
            pass
    for col, nome_lista in (((10, "lst_Categoria"), (12, "lst_FormaPagamento"),
                             (16, "lst_CentroCusto"), (17, "lst_Placa")) if cadastro else ()):
        letra = chr(64 + col)
        faixa = nova.Range(nova.Cells(2, col), nova.Cells(ultima_fmt, col))
        for sep in (";", ","):        # separador de argumentos varia com o idioma do Excel
            try:
                cf = faixa.FormatConditions.Add(2, None, '=AND($%s2<>""%sCOUNTIF(%s%s$%s2)=0)' % (
                    letra, sep, nome_lista, sep, letra))
                cf.Interior.Color = 255
                break
            except Exception:
                continue
    for col in (cfg.COL_VENCIMENTO, cfg.COL_DATA_PAGAMENTO):
        nova.Range(nova.Cells(2, col), nova.Cells(ultima_fmt, col)).NumberFormat = fmt["data"]
    try:
        nova.Range(nova.Cells(1, 1), nova.Cells(1, largura)).AutoFilter()
    except Exception:
        pass
    return nova


def dias_da_aba(ws) -> set:
    """Datas (aaaa-mm-dd) que aparecem na coluna Data de pagamento da aba."""
    fim = ultima_linha(ws)
    if fim < 2:
        return set()
    bloco = ws.Range(ws.Cells(2, cfg.COL_DATA_PAGAMENTO), ws.Cells(fim, cfg.COL_DATA_PAGAMENTO)).Value
    if bloco is None:
        return set()
    if not isinstance(bloco, (tuple, list)):
        bloco = ((bloco,),)
    dias = set()
    for linha in bloco:
        v = linha[0] if isinstance(linha, (tuple, list)) else linha
        if hasattr(v, "strftime"):
            dias.add(v.strftime("%Y-%m-%d"))
    return dias


def dias_repetidos(wb, aba: str) -> dict:
    """{outra aba: [dias em comum]} entre `aba` e as demais abas de inserção."""
    meus = dias_da_aba(wb.Worksheets(aba))
    achados = {}
    for outra in abas_insercao(wb):
        if outra == aba:
            continue
        comuns = sorted(meus & dias_da_aba(wb.Worksheets(outra)))
        if comuns:
            achados[outra] = comuns
    return achados


def texto_dias_repetidos(achados: dict) -> str:
    linhas = []
    for outra, dias in achados.items():
        bonitos = ", ".join("%s/%s" % (d[8:10], d[5:7]) for d in dias)
        linhas.append("    - %s (dias %s)" % (outra, bonitos))
    return "\n".join(linhas)


def confirmar_sn(pergunta: str, sem_perguntas: bool = False) -> bool:
    """Pergunta S/N na janela preta. Padrão é NÃO (Enter = não)."""
    if sem_perguntas:
        return True
    r = input("\n  %s (S/N) > " % pergunta).strip().upper()
    return r in ("S", "SIM")


def avisar_dias_repetidos(aba: str, sem_perguntas: bool, acao: str) -> bool:
    """Avisa se `aba` tem dias que também estão em outra aba de inserção.

    Devolve True se pode continuar (não há repetição, ou você confirmou).
    """
    with com_workbook() as wb:
        achados = dias_repetidos(wb, aba)
    if not achados:
        return True
    print("\n  >> ATENÇÃO: a aba '%s' tem lançamentos de dias que também estão em outra aba:" % aba)
    print(texto_dias_repetidos(achados))
    return confirmar_sn("Você vai %s em lançamentos dessas datas. Tem certeza que quer continuar?" % acao,
                        sem_perguntas)
