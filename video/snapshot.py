# -*- coding: utf-8 -*-
"""Imagens da planilha feitas pelo próprio Excel (instância invisível), sem usar a área de transferência.

    snapshot("conferencia_a.png", "Conferência NF", "A1:J20")

O Excel exporta o intervalo para PDF e o PDF é rasterizado: fontes, cores e larguras são as reais.
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import pymupdf
from PIL import Image, ImageChops

import config as cfg
import planilha

SAIDA = Path(__file__).resolve().parent / "imagens"
SAIDA.mkdir(exist_ok=True)


def _aparar(im: Image.Image) -> Image.Image:
    bg = Image.new("RGB", im.size, (255, 255, 255))
    caixa = ImageChops.difference(im.convert("RGB"), bg).getbbox()
    return im.crop((max(caixa[0] - 6, 0), max(caixa[1] - 6, 0), caixa[2] + 6, caixa[3] + 6)) if caixa else im


def snapshots(pedidos: list, dpi: int = 300):
    """pedidos = [(arquivo_png, aba, intervalo, [colunas_a_ocultar])]; uma abertura do Excel para todos."""
    resultados = []
    with planilha.com_workbook() as wb:
        for nome, aba, area, ocultar in pedidos:
            ws = wb.Worksheets(aba)
            for c in ocultar:
                ws.Columns(c).Hidden = True
            ps = ws.PageSetup
            ps.PrintArea = area
            ps.Orientation = 2                      # paisagem
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 10
            ps.CenterHorizontally = False
            ps.PrintGridlines = False
            pdf = SAIDA / (Path(nome).stem + ".pdf")
            ws.ExportAsFixedFormat(0, str(pdf))
            doc = pymupdf.open(pdf)
            pix = doc[0].get_pixmap(dpi=dpi)
            png = SAIDA / nome
            pix.save(png)
            doc.close()
            pdf.unlink()
            _aparar(Image.open(png)).save(png)
            for c in ocultar:
                ws.Columns(c).Hidden = False
            resultados.append(png)
    return resultados


if __name__ == "__main__":
    for p in snapshots([("teste_semana.png", cfg.PLANILHA and "Inserção - 13.10", "A1:Q26", [])]):
        print(p, Image.open(p).size)
