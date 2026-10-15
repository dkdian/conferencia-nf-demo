# -*- coding: utf-8 -*-
"""
Instala na planilha da demonstração:

  * o módulo VBA  Conferencia_NF  (macros ConferirNotasFiscais, DarBaixaNoSistema e GerarArquivoInsercao)
  * a aba "Conferência NF" com os três botões numerados que disparam essas macros

Pode rodar de novo à vontade: recria o módulo e reaproveita a aba e os botões.
Exige o Excel com "Confiar no acesso ao modelo de objeto do projeto do VBA" ligado
(Arquivo > Opções > Central de Confiabilidade > Configurações de Macro).

    python instalar_botao.py                 instala na planilha da demonstração
    python instalar_botao.py --arquivo X     instala em outro .xlsm
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pythoncom
import win32com.client as win32

import aplicar_na_planilha as layout
import config as cfg
import planilha

NOME_MODULO = "Conferencia_NF"
BOTOES = [   # (nome do botão, macro, texto, cor BGR)
    ("BotaoConferirNF", "ConferirNotasFiscais", "1. Conferir notas no sistema", 0x50A000),
    ("BotaoDarBaixa", "DarBaixaNoSistema", "2. Dar baixa no sistema", 0x2050D0),
    ("BotaoGerarInsercao", "GerarArquivoInsercao", "3. Gerar arquivo de inserção", 0xA05000),
]

CODIGO_VBA = '''Option Explicit

' Cada botao chama um .bat da pasta da demonstracao, que roda o Python.
' A planilha em si nao faz conta nenhuma: ela e a tela de controle.

Private Const PASTA_AUTOMACAO As String = "@@PASTA@@"

Private Function BatExiste(nome As String) As Boolean
    BatExiste = (Dir(PASTA_AUTOMACAO & "\\" & nome) <> "")
    If Not BatExiste Then
        MsgBox "Nao encontrei o script em:" & vbCrLf & PASTA_AUTOMACAO, vbCritical, nome
    End If
End Function

' Botao 1: confere cada Nota Fiscal da aba de inserção no ERP e escreve a aba Conferência NF.
Public Sub ConferirNotasFiscais()
    If Not BatExiste("Conferir NF.bat") Then Exit Sub
    If MsgBox("Finalizou a planilha?" & vbCrLf & vbCrLf & _
        "Vou conferir no ERP as linhas com Nota Fiscal da aba Inserção." & vbCrLf & _
        "Uma janela preta vai abrir, e depois o navegador do robô." & vbCrLf & vbCrLf & _
        "Enquanto roda, não edite células desta planilha." & vbCrLf & _
        "Ao terminar, o resultado aparece na aba Conferência NF.", _
        vbYesNo + vbQuestion, "Conferir notas fiscais no ERP") <> vbYes Then Exit Sub
    Shell "cmd /c """ & PASTA_AUTOMACAO & "\\Conferir NF.bat""", vbNormalFocus
End Sub

' Botao 2: mostra uma PREVIA do que sera feito (nada e gravado) e so da baixa de verdade
' se voce digitar BAIXAR na janela preta.
Public Sub DarBaixaNoSistema()
    If Not BatExiste("Dar Baixa.bat") Then Exit Sub
    If MsgBox("Dar baixa no ERP das linhas marcadas DAR BAIXA na aba Conferência NF?" & vbCrLf & vbCrLf & _
        "Primeiro aparece uma PRÉVIA na janela preta, sem gravar nada." & vbCrLf & _
        "Você escolhe SIMULAR (testa e cancela) ou BAIXAR (grava de verdade)." & vbCrLf & _
        "Depois de baixar de verdade, as linhas baixadas saem da aba de inserção " & _
        "(ficam guardadas na aba '... completa')." & vbCrLf & _
        "Não edite células enquanto roda.", _
        vbYesNo + vbQuestion, "Dar baixa no ERP") <> vbYes Then Exit Sub
    Shell "cmd /c """ & PASTA_AUTOMACAO & "\\Dar Baixa.bat""", vbNormalFocus
End Sub

' Botao 3: monta o Inserção.xlsx de importacao a partir da aba de insercao, sem as linhas
' baixadas ou marcadas IGNORAR na Conferência NF.
Public Sub GerarArquivoInsercao()
    If Not BatExiste("Gerar Insercao.bat") Then Exit Sub
    If MsgBox("Gerar o Inserção.xlsx agora?" & vbCrLf & vbCrLf & _
        "Ele é montado a partir da aba Inserção, SEM as linhas baixadas ou " & _
        "marcadas IGNORAR na aba Conferência NF." & vbCrLf & _
        "Linhas com DAR BAIXA ainda não executada, ou sem Ação, ficam de fora " & _
        "e são listadas na janela preta.", _
        vbYesNo + vbQuestion, "Gerar arquivo de inserção") <> vbYes Then Exit Sub
    Shell "cmd /c """ & PASTA_AUTOMACAO & "\\Gerar Insercao.bat""", vbNormalFocus
End Sub
'''


def instalar_modulo(wb):
    componentes = wb.VBProject.VBComponents
    for c in componentes:
        if c.Name == NOME_MODULO:
            componentes.Remove(c)
            break
    modulo = componentes.Add(1)  # vbext_ct_StdModule
    modulo.Name = NOME_MODULO
    modulo.CodeModule.AddFromString(CODIGO_VBA.replace("@@PASTA@@", str(cfg.BASE)))
    return modulo


def pegar_ou_criar_aba(wb):
    for aba in wb.Worksheets:
        if aba.Name == cfg.ABA_RESULTADO:
            return aba
    abas = planilha.abas_insercao(wb)
    ref = wb.Worksheets(abas[-1]) if abas else wb.Worksheets(wb.Worksheets.Count)
    nova = wb.Worksheets.Add(None, ref)  # (Before, After)
    nova.Name = cfg.ABA_RESULTADO
    return nova


def _criar_botao(ws, nome, macro, texto, esquerda, topo, largura, cor):
    for forma in list(ws.Shapes):
        if forma.Name == nome:
            forma.Delete()
    botao = ws.Shapes.AddShape(5, esquerda, topo, largura, 40)  # retângulo arredondado
    botao.Name = nome
    botao.OnAction = macro
    botao.Fill.ForeColor.RGB = cor
    botao.Line.Visible = False
    quadro = botao.TextFrame2
    quadro.TextRange.Text = texto
    quadro.TextRange.Font.Size = 11
    quadro.TextRange.Font.Bold = True
    quadro.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    quadro.WordWrap = True
    quadro.VerticalAnchor = 3                  # centralizado
    quadro.TextRange.ParagraphFormat.Alignment = 2
    botao.Placement = 3                        # não muda de tamanho com as células


def instalar_botoes(ws):
    # Larguras iguais às da tabela (aplicar_na_planilha): os botões ficam no alto, sem cobrir a tabela.
    for i, largura in enumerate(layout.LARGURAS, 1):
        ws.Columns(i).ColumnWidth = largura
    esquerda = ws.Range("A1").Left + 6
    topo = ws.Range("D1").Top + 4
    larguras = (230, 210, 230)
    x = esquerda + ws.Range("E1").Left   # à direita do título, sem cobri-lo
    for (nome, macro, texto, cor), largura in zip(BOTOES, larguras):
        _criar_botao(ws, nome, macro, texto, x, topo, largura, cor)
        x += largura + 12
    if ws.Range("A1").Value in (None, ""):
        ws.Range("A1").Value = "Conferência das Notas Fiscais no ERP"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
        ws.Range("A2").Value = "Ainda não conferido. Clique no botão 1 quando finalizar a aba de inserção."


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--arquivo", help="instala em outro .xlsm em vez da planilha da demonstração")
    args = ap.parse_args(argv)
    alvo = Path(args.arquivo).resolve() if args.arquivo else cfg.PLANILHA
    cfg.PLANILHA = alvo

    pythoncom.CoInitialize()
    if planilha.achar_workbook_aberto() is not None:
        print("A planilha está ABERTA no Excel. Feche-a e rode de novo.")
        pythoncom.CoUninitialize()
        return 1

    # Instância do Excel só nossa e invisível: nunca encosta nas janelas do usuário.
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(alvo), UpdateLinks=0)
        instalar_modulo(wb)
        ws = pegar_ou_criar_aba(wb)
        instalar_botoes(ws)
        abas = planilha.abas_insercao(wb)
        if abas:
            wb.Worksheets(abas[0]).Activate()
        wb.Save()
        print("Módulo '%s' e botões instalados em: %s" % (NOME_MODULO, alvo.name))
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    sys.exit(main())
