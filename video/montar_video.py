# -*- coding: utf-8 -*-
"""
Monta o vídeo da demonstração a partir do material capturado por capturar.py:

  * terminal: a saída REAL de cada robô, com o horário real de cada trecho (a digitação de SIMULAR/BAIXAR é
    reproduzida na pausa do pedido, porque no teste automático a resposta chega pela entrada padrão);
  * navegador: o vídeo REAL gravado pelo Playwright enquanto o robô trabalha;
  * planilha: imagens exportadas pelo Excel entre os passos;
  * janelas de confirmação: o texto exato das macros dos botões.

Nada vem do computador de ninguém: não há captura de tela.   python video/montar_video.py [saida.mp4]
"""
from __future__ import annotations

import json
import math
import os
import re
import shutil
import subprocess
import sys
import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

AQUI = Path(__file__).resolve().parent
MAT, IMG = AQUI / "material", AQUI / "imagens"
TMP = AQUI / "_quadros"
FFMPEG = next(Path(os.environ["LOCALAPPDATA"], "Microsoft", "WinGet", "Packages").glob("Gyan.FFmpeg*/*/bin/ffmpeg.exe"))
W, H, FPS = 1280, 720, 15
FONTES = "C:/Windows/Fonts/"


def fonte(nome, tam):
    return ImageFont.truetype(FONTES + nome, tam)


UI, UIB, MONO, MONOB = "segoeui.ttf", "segoeuib.ttf", "consola.ttf", "consolab.ttf"
FUNDO, TOPO, TINTA, SUAVE = (15, 23, 36), (10, 16, 26), (232, 238, 246), (150, 164, 184)
VERDE, LARANJA, AZUL, AMARELO = (46, 160, 120), (232, 120, 60), (80, 150, 230), (255, 214, 90)


# ------------------------------------------------------------ gravação ---
class Saida:
    def __init__(self, destino: Path):
        self.destino = destino
        self.p = subprocess.Popen(
            [str(FFMPEG), "-y", "-hide_banner", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24",
             "-s", "%dx%d" % (W, H), "-r", str(FPS), "-i", "-", "-c:v", "libx264", "-preset", "medium", "-crf", "27",
             "-pix_fmt", "yuv420p", "-movflags", "+faststart", str(destino)], stdin=subprocess.PIPE)
        self.quadros = 0

    def escrever(self, im: Image.Image, vezes: int = 1):
        b = im.convert("RGB").tobytes()
        for _ in range(vezes):
            self.p.stdin.write(b)
        self.quadros += vezes

    def parar(self):
        self.p.stdin.close()
        self.p.wait()
        return self.quadros / FPS


# ------------------------------------------------------------- base ---
def base(passo: str, titulo: str) -> Image.Image:
    im = Image.new("RGB", (W, H), FUNDO)
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 62], fill=TOPO)
    d.ellipse([22, 14, 56, 48], fill=VERDE)
    f = fonte(UIB, 20)
    d.text((39 - d.textlength(passo, font=f) / 2, 17), passo, font=f, fill="white")
    d.text((72, 15), titulo, font=fonte(UIB, 24), fill="white")
    d.text((W - 24 - d.textlength("dados fictícios", font=fonte(UI, 15)), 22), "dados fictícios", font=fonte(UI, 15), fill=SUAVE)
    return im


def legenda(im: Image.Image, texto: str, y=None):
    d = ImageDraw.Draw(im)
    f = fonte(UI, 21)
    linhas = textwrap.wrap(texto, 112)
    y = y if y is not None else H - 20 - 30 * len(linhas)
    for l in linhas:
        d.text((40, y), l, font=f, fill=TINTA)
        y += 30


def ajustar(img: Image.Image, w: int, h: int) -> Image.Image:
    esc = min(w / img.width, h / img.height)
    return img.resize((max(1, int(img.width * esc)), max(1, int(img.height * esc))), Image.LANCZOS)


# ------------------------------------------------------------ cenas estáticas ---
def cena_cartao(s: Saida, linhas: list, dur: float, destaque: str = ""):
    im = Image.new("RGB", (W, H), TOPO)
    d = ImageDraw.Draw(im)
    y = 150 if len(linhas) > 4 else 220
    for i, (txt, tam, cor) in enumerate(linhas):
        f = fonte(UIB if tam >= 24 else UI, tam)
        d.text(((W - d.textlength(txt, font=f)) / 2, y), txt, font=f, fill=cor)
        y += int(tam * 1.55)
    d.rectangle([0, H - 8, W, H], fill=VERDE)
    s.escrever(im, int(dur * FPS))


def cena_folha(s: Saida, passo: str, titulo: str, arq: str, texto: str, dur: float):
    im = base(passo, titulo)
    folha = ajustar(Image.open(IMG / arq).convert("RGB"), W - 80, H - 62 - 120)
    x, y = (W - folha.width) // 2, 62 + 22
    d = ImageDraw.Draw(im)
    d.rectangle([x - 3, y - 3, x + folha.width + 3, y + folha.height + 3], fill=(60, 72, 92))
    im.paste(folha, (x, y))
    legenda(im, texto, y=y + folha.height + 18)
    s.escrever(im, int(dur * FPS))


def cena_antes_depois(s: Saida, passo: str, titulo: str, a: str, b: str, texto: str, dur: float):
    imA = Image.open(IMG / a).convert("RGB")
    imB = Image.open(IMG / b).convert("RGB")
    fa, fb = ajustar(imA, W - 80, H - 62 - 120), ajustar(imB, W - 80, H - 62 - 120)
    total = int(dur * FPS)
    for i in range(total):
        t = i / total
        mix = 0.0 if t < 0.35 else (1.0 if t > 0.55 else (t - 0.35) / 0.2)
        fm = Image.blend(fa, fb.resize(fa.size), mix) if fa.size != fb.size or True else fa
        im = base(passo, titulo)
        x, y = (W - fm.width) // 2, 84
        ImageDraw.Draw(im).rectangle([x - 3, y - 3, x + fm.width + 3, y + fm.height + 3], fill=(60, 72, 92))
        im.paste(fm, (x, y))
        legenda(im, texto, y=y + fm.height + 18)
        s.escrever(im)


def desenhar_dialogo(titulo: str, texto: str, botoes=("Sim", "Não"), largura=780) -> Image.Image:
    f, ft, fb = fonte(UI, 19), fonte(UI, 16), fonte(UI, 18)
    cols = int((largura - 130) / 9.2)
    linhas = []
    for par in texto.split("\n"):
        linhas += textwrap.wrap(par, cols) or [""]
    h = 44 + 28 + len(linhas) * 27 + 24 + 70
    im = Image.new("RGB", (largura, h), "white")
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, largura - 1, h - 1], outline=(120, 120, 120))
    d.text((14, 11), titulo, font=ft, fill=(0, 0, 0))
    d.ellipse([28, 66, 76, 114], fill=(0, 103, 192))
    d.text((43, 66), "?", font=fonte(UIB, 34), fill="white")
    y = 68
    for l in linhas:
        d.text((100, y), l, font=f, fill=(0, 0, 0))
        y += 27
    d.rectangle([1, h - 70, largura - 2, h - 2], fill=(240, 240, 240))
    bx = largura - 24
    for i, b in enumerate(reversed(botoes)):
        w = 118
        principal = i == len(botoes) - 1
        d.rounded_rectangle([bx - w, h - 53, bx, h - 18], 4, fill=(253, 253, 253),
                            outline=(0, 103, 192) if principal else (173, 173, 173), width=3 if principal else 1)
        d.text((bx - w / 2 - d.textlength(b, font=fb) / 2, h - 48), b, font=fb, fill=(0, 0, 0))
        bx -= w + 12
    return im


def cena_dialogo(s: Saida, passo: str, titulo: str, dlg: Image.Image, texto: str, dur: float):
    im = base(passo, titulo)
    d = ImageDraw.Draw(im)
    x, y = (W - dlg.width) // 2, 110
    d.rectangle([x + 8, y + 8, x + dlg.width + 8, y + dlg.height + 8], fill=(8, 12, 20))
    im.paste(dlg, (x, y))
    legenda(im, texto, y=H - 70)
    s.escrever(im, int(dur * FPS))


# ------------------------------------------------------------- terminal ---
def cor_da_linha(l: str):
    s = l.strip()
    if s.startswith("[OK") or "SIMULAÇÃO OK" in s or s.startswith("BAIXADA") or s.startswith("Já no sistema") \
            or s.startswith("Ajustar valor") or s.startswith("Não encontradas") or s.startswith("Para revisar") \
            or s.startswith("Entraram no arquivo") or s.startswith("Arquivo:"):
        return (120, 220, 130)
    if "PROBLEMA" in s or s.startswith("[NÃO") or "ATENÇÃO" in s or s.startswith("MODO REAL") or "NÃO foi salvo" in s \
            or "BLOQUEADO" in s or s.startswith("aviso:") or "NÃO entraram" in s or s.startswith("linha ") and "->" in s and "DAR BAIXA" in s:
        return (255, 150, 110)
    if s.startswith("PRÉVIA") or s.startswith("MODO SIMULAÇÃO"):
        return (130, 190, 255)
    return (204, 204, 204)


class Terminal:
    def __init__(self, resumo: dict, banner: str, final: bool = True):
        self.eventos = [[0.0, banner]] + [list(e) for e in resumo["eventos"]]
        self.fim_real = resumo["duracao"]
        if final:
            self.eventos.append([self.fim_real + 0.4, "\nConcluido. Pode fechar esta janela.\nPressione qualquer tecla para continuar . . .\n"])
        self.resp = resumo.get("entrada") or ""
        self.t_prompt = None
        for t, txt in self.eventos:
            if self.resp and txt.endswith("> ") and self.t_prompt is None:
                self.t_prompt = t
        acumulado, partes = [], ""
        for t, txt in self.eventos:
            partes += txt
            acumulado.append((t, partes))
        self.acum = acumulado
        # depois do prompt, a resposta digitada passa a fazer parte do texto
        if self.t_prompt is not None:
            novo = []
            inserido = False
            for t, txt in acumulado:
                if self.t_prompt is not None and t > self.t_prompt and not inserido:
                    inserido = True
                novo.append((t, txt))
            self.acum = novo

    def texto(self, r: float, digitado: float | None) -> str:
        atual = ""
        for t, txt in self.acum:
            if t <= r:
                atual = txt
            else:
                break
        if self.t_prompt is not None:
            if r > self.t_prompt:
                # a resposta entra logo depois do prompt, na mesma linha
                i = atual.index("\n  > ") + len("\n  > ") if "\n  > " in atual else len(atual)
                atual = atual[:i] + self.resp + "\n" + atual[i:]
            elif digitado is not None:
                atual = atual + self.resp[:int(round(len(self.resp) * digitado))]
        return atual


def linhas_quebradas(texto: str, cols: int) -> list:
    saida = []
    for l in texto.split("\n"):
        if len(l) <= cols:
            saida.append(l)
        else:
            pref = len(l) - len(l.lstrip())
            for k, parte in enumerate(textwrap.wrap(l, cols - 2, subsequent_indent=" " * (pref + 2)) or [""]):
                saida.append(parte)
    return saida


def desenhar_terminal(texto: str, w: int, h: int, tam: int, cursor: bool) -> Image.Image:
    f, fb = fonte(MONO, tam), fonte(MONOB, tam)
    cw = f.getlength("M")
    lh = int(tam * 1.42)
    im = Image.new("RGB", (w, h), (12, 12, 12))
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, w, 30], fill=(236, 238, 241))
    d.rectangle([10, 8, 24, 22], fill=(12, 12, 12))
    d.text((32, 6), "Prompt de Comando", font=fonte(UI, 15), fill=(40, 44, 52))
    cols = int((w - 24) / cw)
    ls = linhas_quebradas(texto.rstrip("\n") if not texto.endswith("\n\n") else texto, cols)
    vis = (h - 44) // lh
    ls = ls[-vis:]
    y = 38
    for i, l in enumerate(ls):
        cor = cor_da_linha(l)
        d.text((12, y), l, font=fb if cor != (204, 204, 204) else f, fill=cor)
        y += lh
    if cursor and ls:
        x = 12 + cw * len(ls[-1])
        d.rectangle([x, y - lh + 3, x + cw * 0.7, y - 4], fill=(204, 204, 204))
    return im


BANNERS = {
    "rotina": "",
    "conferir": "\n==========================================================\n  Conferencia de Notas Fiscais - demonstracao\n==========================================================\n\n",
    "baixa": "\n==========================================================\n  Dar baixa no ERP - demonstracao\n==========================================================\n\n",
    "gerar": "\n==========================================================\n  Gerar arquivo de insercao - demonstracao\n==========================================================\n\n",
}


def carregar(nome: str) -> dict:
    return json.loads((MAT / nome / "resumo.json").read_text(encoding="utf-8"))


def quadros_navegador(nome: str, resumo: dict, largura: int = 576):
    pasta = TMP / ("nav_" + nome)
    shutil.rmtree(pasta, ignore_errors=True)
    pasta.mkdir(parents=True)
    if not resumo.get("video"):
        return pasta, 0
    subprocess.run([str(FFMPEG), "-y", "-hide_banner", "-loglevel", "error", "-i", str(MAT / nome / resumo["video"]),
                    "-vf", "fps=10,scale=%d:-1" % largura, str(pasta / "%05d.png")], check=True)
    return pasta, len(list(pasta.glob("*.png")))


def atividade_navegador(pasta: Path, n: int) -> list:
    """Para cada quadro (10 por segundo) do vídeo do navegador: True se a tela mudou em relação ao anterior."""
    import numpy as np
    saida, anterior = [], None
    for i in range(1, n + 1):
        atual = np.asarray(Image.open(pasta / ("%05d.png" % i)).convert("L").resize((96, 56)), dtype=np.int16)
        saida.append(anterior is not None and float(np.abs(atual - anterior).mean()) > 0.35)
        anterior = atual
    return saida


def moldura_navegador(quadro: Image.Image | None, w=576, depois: bool = False) -> Image.Image:
    alt = quadro.height if quadro is not None else int(w * 0.75)
    im = Image.new("RGB", (w, alt + 30), (240, 242, 246))
    d = ImageDraw.Draw(im)
    for i, c in enumerate(((237, 106, 94), (245, 191, 79), (98, 197, 84))):
        d.ellipse([10 + i * 20, 9, 22 + i * 20, 21], fill=c)
    d.rounded_rectangle([86, 5, w - 12, 26], 10, fill="white")
    d.text((98, 6), "navegador do robô · ERP fictício", font=fonte(UI, 14), fill=(90, 100, 114))
    if quadro is not None:
        im.paste(quadro, (0, 30))
    else:
        d.rectangle([0, 30, w, alt + 30], fill=(26, 36, 52))
        t = "navegador fechado: o robô terminou" if depois else "o robô ainda não abriu o navegador"
        d.text(((w - d.textlength(t, font=fonte(UI, 16))) / 2, alt / 2 + 20), t, font=fonte(UI, 16), fill=SUAVE)
    return im


def tempos(total: float, vel, holds: list):
    """Gera (r, digitado) por quadro de vídeo. vel(r)=velocidade de reprodução; holds=[(r, segundos)]."""
    saida, r, usados = [], 0.0, set()
    while r < total:
        for i, (rh, secs) in enumerate(holds):
            if i not in usados and r >= rh:
                usados.add(i)
                n = int(secs * FPS)
                for k in range(n):
                    saida.append((rh, min(1.0, k / max(1, int(n * 0.55)))))
        saida.append((r, None))
        r += vel(r) / FPS
    return saida


def cena_robo(s: Saida, passo: str, titulo: str, nome: str, banner: str, vel_nav: float, hold_secs: float,
              pontos: list, solo: bool = False):
    resumo = carregar(nome)
    term = Terminal(resumo, banner)
    pasta, n_nav = quadros_navegador(nome, resumo)
    t_pag = resumo["t_pagina"] or 0
    dur_nav = n_nav / 10.0
    holds = [(term.t_prompt, hold_secs)] if term.t_prompt is not None else []
    instantes = sorted({e[0] for e in term.eventos})
    ativo = atividade_navegador(pasta, n_nav)          # segundos (da página) em que a tela do navegador mudou

    def vel(r):
        """Esperas (do Excel, por exemplo) são puladas; o que se mexe na tela e as linhas novas ficam legíveis."""
        v = 12.0
        if n_nav and 0 <= r - t_pag <= dur_nav + 0.5:
            k = int((r - t_pag) * 10)
            if any(ativo[j] for j in range(max(0, k - 6), min(len(ativo), k + 7))):
                v = vel_nav
        if any(abs(r - e) < 0.6 for e in instantes):
            v = min(v, 1.5)
        return v
    linha = tempos(term.eventos[-1][0] + 1.0, vel, holds)
    cache_nav = {}
    for i, (r, dig) in enumerate(linha):
        im = base(passo, titulo)
        if solo:
            im.paste(desenhar_terminal(term.texto(r, dig), W - 32, H - 62 - 30, 19, (i // 8) % 2 == 0), (16, 78))
        else:
            im.paste(desenhar_terminal(term.texto(r, dig), 656, H - 62 - 30, 14, (i // 8) % 2 == 0), (16, 78))
            pt = r - t_pag
            quadro = None
            if 0 <= pt <= dur_nav + 0.3 and n_nav:
                idx = min(n_nav, int(pt * 10) + 1)
                if idx not in cache_nav:
                    cache_nav[idx] = Image.open(pasta / ("%05d.png" % idx)).convert("RGB")
                quadro = cache_nav[idx]
            im.paste(moldura_navegador(quadro, depois=pt > dur_nav), (688, 78))
            d = ImageDraw.Draw(im)
            y = 78 + 432 + 30 + 14
            for txt in pontos:
                for k, l in enumerate(textwrap.wrap(txt, 64)):
                    d.text((692 if k == 0 else 710, y), ("•  " + l) if k == 0 else l, font=fonte(UI, 16), fill=TINTA)
                    y += 22
                y += 3
        s.escrever(im)



def cena_navegador(s: Saida, passo: str, titulo: str, nome: str, texto: str, largura: int = 1000):
    """O vídeo real do navegador do robô, em tamanho grande, no tempo real."""
    resumo = carregar(nome)
    pasta, n_nav = quadros_navegador(nome, resumo, largura)
    cache = {}
    for i in range(n_nav):
        im = base(passo, titulo)
        quadro = Image.open(pasta / ("%05d.png" % (i + 1))).convert("RGB")
        moldura = moldura_navegador(quadro, w=quadro.width)
        x = (W - moldura.width) // 2
        d = ImageDraw.Draw(im)
        d.rectangle([x - 3, 74, x + moldura.width + 3, 74 + moldura.height + 3], fill=(60, 72, 92))
        im.paste(moldura, (x, 77))
        legenda(im, texto, y=H - 44)
        s.escrever(im, 3 if i % 2 == 0 else 2)       # 10 quadros por segundo de origem -> 15 fps


# --------------------------------------------------------------- roteiro ---
def numeros():
    def pega(nome, padrao):
        txt = "".join(t for _, t in carregar(nome)["eventos"])
        return re.search(padrao, txt)
    conf = "".join(t for _, t in carregar("conferir")["eventos"])
    real = "".join(t for _, t in carregar("baixa_real")["eventos"])
    ger = "".join(t for _, t in carregar("gerar")["eventos"])
    g = lambda pat, txt: int(re.search(pat, txt).group(1))
    return {
        "ja": g(r"Já no sistema\s*:\s*(\d+)", conf), "ajustar": g(r"Ajustar valor\s*:\s*(\d+)", conf),
        "nao": g(r"Não encontradas\s*:\s*(\d+)", conf), "revisar": g(r"Para revisar\s*:\s*(\d+)", conf),
        "baixadas": len(re.findall(r"BAIXADA —", real)), "no_arquivo": g(r"Entraram no arquivo\s*:\s*(\d+)", ger),
        "pendentes": g(r"(\d+) linha\(s\) NÃO entraram", ger),
        "semana": g(r"Aba criada: .*?\((\d+) linhas\)", "".join(x for _, x in carregar("rotina")["eventos"])),
        "com_nf": g(r"(\d+) linhas com Nota Fiscal", conf),
    }


def _medir(funcao):
    def interna(s, *args, **kw):
        antes = s.quadros
        funcao(s, *args, **kw)
        print("  %-16s %-48s %5.1f s" % (funcao.__name__, str(args[1])[:48], (s.quadros - antes) / FPS))
    return interna


cena_cartao, cena_folha, cena_antes_depois = _medir(cena_cartao), _medir(cena_folha), _medir(cena_antes_depois)
cena_dialogo, cena_robo, cena_navegador = _medir(cena_dialogo), _medir(cena_robo), _medir(cena_navegador)


def main():
    destino = Path(sys.argv[1]) if len(sys.argv) > 1 else AQUI / "demonstracao.mp4"
    shutil.rmtree(TMP, ignore_errors=True)
    TMP.mkdir()
    n = numeros()
    print("números reais:", n)
    s = Saida(destino)

    cena_cartao(s, [("Conferência de notas fiscais e baixa no ERP", 44, "white"),
                    ("Demonstração com dados fictícios", 26, AMARELO),
                    ("Três botões numa planilha: conferir, dar baixa e gerar o arquivo de importação.", 21, SUAVE),
                    ("O robô e o ERP que você vai ver são de verdade; só os dados são inventados.", 21, SUAVE)], 4.5)

    cena_cartao(s, [("Toda terça-feira, às 10h", 46, "white"),
                    ("o robô acorda sozinho (Agendador de Tarefas do Windows)", 25, AMARELO),
                    ("Baixa os relatórios do Drive, organiza a pasta, cria a aba da semana e avisa o gestor.", 21, SUAVE)], 4)
    cena_robo(s, "1", "Rotina de terça: do Drive à aba da semana", "rotina", BANNERS["rotina"], 1.0, 0, [], solo=True)
    cena_folha(s, "1", "A aba da semana, criada pelo robô", "s1_semana.png",
               "%d lançamentos pagos na semana, %d com nota fiscal. Nenhuma aba existente é alterada: a aba nasce nova, com a data." % (n["semana"], n["com_nf"]), 6.5)
    cena_navegador(s, "2", "O gestor recebe o aviso e valida", "gestor",
                   "Antes de lançar qualquer coisa, o gestor confere as informações e responde OK no chat.")

    cena_dialogo(s, "3", "Botão 1: conferir notas no ERP", desenhar_dialogo(
        "Conferir notas fiscais no ERP",
        "Finalizou a planilha?\n\nVou conferir no ERP as linhas com Nota Fiscal da aba Inserção.\n"
        "Uma janela preta vai abrir, e depois o navegador do robô.\n\nEnquanto roda, não edite células desta planilha.\n"
        "Ao terminar, o resultado aparece na aba Conferência NF."),
        "Um clique em Sim e o robô começa. O que ele faz é só consultar: nada é gravado no ERP neste passo.", 3)

    r = carregar("conferir")
    cena_robo(s, "3", "O robô procura cada nota no ERP", "conferir", BANNERS["conferir"],
              1.6, 0,
              ["Busca em lotes de 10 números, de forma exata.",
               "O que o lote não acha é buscado de novo, por trecho: números gravados com ano e zeros na frente não viram \"não encontrada\" por engano.",
               "O robô não lê a tela: escuta as respostas do próprio ERP."])

    cena_folha(s, "4", "Cada linha recebe um resultado e uma ação sugerida", "s2_azul.png",
               "Azul: o que veio da planilha. Já no sistema: %d · Ajustar valor: %d · Não encontradas: %d · Para revisar: %d."
               % (n["ja"], n["ajustar"], n["nao"], n["revisar"]), 6.5)
    cena_folha(s, "4", "O que o ERP tem (laranja)", "s2_laranja.png",
               "Laranja: o lançamento mais parecido que o robô achou no ERP, com situação (em aberto ou pago).", 5.5)
    cena_folha(s, "4", "A decisão (amarelo): resultado e ação sugerida", "s2_decisao.png",
               "DAR BAIXA, IMPORTAR ou IGNORAR. Nas linhas em REVISAR o robô não sugere nada: quem decide é a pessoa.", 6)

    cena_antes_depois(s, "5", "A pessoa decide só o que o robô não tem certeza", "s2_decisao.png", "s3_decisao.png",
                      "Três linhas em branco (valor diferente, fornecedor parecido, número com ano e zeros): a pessoa escolhe DAR BAIXA.", 6)

    cena_dialogo(s, "6", "Botão 2: dar baixa no ERP", desenhar_dialogo(
        "Dar baixa no ERP",
        "Dar baixa no ERP das linhas marcadas DAR BAIXA na aba Conferência NF?\n\n"
        "Primeiro aparece uma PRÉVIA na janela preta, sem gravar nada.\n"
        "Você escolhe SIMULAR (testa e cancela) ou BAIXAR (grava de verdade).\n"
        "Depois de baixar de verdade, as linhas baixadas saem da aba de inserção (ficam guardadas na aba '... completa').\n"
        "Não edite células enquanto roda."),
        "A única etapa que grava no ERP tem três travas: prévia, simulação e a palavra BAIXAR digitada pela pessoa.", 3.2)

    r = carregar("baixa_sim")
    cena_robo(s, "6", "Prévia e simulação: nada é gravado", "baixa_sim", BANNERS["baixa"],
              5.5, 9.0,
              ["Prévia: valor, data, conta e forma de cada linha. A nota 2250 está sem forma de pagamento e é recusada.",
               "Diferença de valor entra como juros ou desconto.",
               "Simulação: preenche cada janela e cancela, com a gravação bloqueada na rede."])

    r = carregar("baixa_real")
    cena_robo(s, "7", "Baixa de verdade: confere cada pagamento", "baixa_real", BANNERS["baixa"],
              5.5, 5.0,
              ["Só começa depois que a pessoa digita BAIXAR.",
               "Depois de cada Confirmar, o robô consulta o ERP: o lançamento precisa estar PAGO, na data certa.",
               "No primeiro erro ele para. As linhas baixadas saem da aba, com cópia."])

    cena_folha(s, "7", "Na planilha: o registro de cada baixa", "s4_conferencia.png",
               "As linhas baixadas saem da aba de inserção (\"retirada\") e ficam registradas, uma a uma.", 5.5)
    cena_folha(s, "7", "Status, data e hora de cada baixa", "s4_execucao.png",
               "BAIXADA, com valor, data, conta e forma. A nota sem forma de pagamento ficou NÃO EXECUTADA, com o motivo.", 5.5)
    cena_folha(s, "7", "Nada se perde: a aba \"completa\"", "s4_completa.png",
               "Cópia de tudo, com o que foi baixado marcado em verde. A aba de trabalho fica só com o que ainda precisa subir.", 6)

    cena_dialogo(s, "8", "Botão 3: gerar o arquivo de importação", desenhar_dialogo(
        "Gerar arquivo de inserção",
        "Gerar o Inserção.xlsx agora?\n\nEle é montado a partir da aba Inserção, SEM as linhas baixadas ou marcadas IGNORAR "
        "na aba Conferência NF.\nLinhas com DAR BAIXA ainda não executada, ou sem Ação, ficam de fora e são listadas na janela preta."),
        "O arquivo que vai para o ERP não repete o que já foi baixado: acabou a limpeza manual.", 3)
    cena_robo(s, "8", "O arquivo sai sem duplicatas", "gerar", BANNERS["gerar"], 1.0, 0, [], solo=True)
    cena_folha(s, "8", "O arquivo de importação", "s5_arquivo.png",
               "%d linhas: as sem nota fiscal e as notas que não estão no ERP. A nota pendente fica de fora, com aviso." % n["no_arquivo"], 7)

    cena_cartao(s, [("Resultado desta semana de demonstração", 38, "white"),
                    ("Terça às 10h: relatórios baixados, pasta organizada, aba criada e gestor avisado, sozinhos", 24, (130, 225, 160)),
                    ("%d notas baixadas pelo robô, com prévia, simulação e conferência depois de gravar" % n["baixadas"], 24, (130, 225, 160)),
                    ("%d linhas no arquivo de importação, sem duplicatas" % n["no_arquivo"], 24, (130, 225, 160)),
                    ("%d linha pendente listada com o motivo, em vez de passar despercebida" % n["pendentes"], 24, (130, 225, 160)),
                    ("A pessoa só decidiu 3 casos; o resto foi conferência e registro.", 24, TINTA),
                    ("Código e testes: ver o repositório da demonstração.", 20, SUAVE)], 7)
    dur = s.parar()
    print("vídeo pronto: %s (%.1f s, %.1f MB)" % (destino, dur, destino.stat().st_size / 1e6))
    shutil.rmtree(TMP, ignore_errors=True)


if __name__ == "__main__":
    main()
