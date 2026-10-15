# -*- coding: utf-8 -*-
"""
Roda a demonstração DE VERDADE, passo a passo, e guarda o material para montar o vídeo:

  * a saída do terminal de cada robô, com o horário de cada trecho;
  * o vídeo do navegador do robô (gravado pelo Playwright, sem capturar a tela do computador);
  * imagens da planilha entre os passos (exportadas pelo Excel numa instância invisível).

Nada aqui mexe no mouse, no teclado ou na tela de ninguém.

    python video/capturar.py        (leva alguns minutos; recomeça a demonstração do zero)
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
PASTA = Path(__file__).resolve().parent / "material"
sys.path[:0] = [str(RAIZ), str(Path(__file__).resolve().parent)]

import config as cfg
import snapshot

ENV = dict(os.environ, PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8", DEMO_HEADLESS="1", DEMO_RITMO="0.6")


def rodar(nome: str, script: str, entrada: str = "", viewport: str = "") -> dict:
    """Roda um robô e guarda terminal + vídeo do navegador. Devolve o resumo."""
    pasta = PASTA / nome
    shutil.rmtree(pasta, ignore_errors=True)
    pasta.mkdir(parents=True)
    env = dict(ENV, DEMO_GRAVAR=str(pasta), DEMO_T0=str(pasta / "t0.txt"))
    if viewport:
        env["DEMO_VIEWPORT"] = viewport
    inicio = time.time()
    proc = subprocess.Popen([sys.executable, script], cwd=str(RAIZ), env=env, stdin=subprocess.PIPE,
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    eventos = []

    def ler():
        import codecs
        dec = codecs.getincrementaldecoder("utf-8")()
        while True:
            dados = os.read(proc.stdout.fileno(), 4096)
            if not dados:
                break
            txt = dec.decode(dados)
            if txt:
                eventos.append([round(time.time() - inicio, 3), txt])
    th = threading.Thread(target=ler, daemon=True)
    th.start()
    if entrada:
        proc.stdin.write(entrada.encode("utf-8"))
    proc.stdin.close()
    codigo = proc.wait(timeout=600)
    th.join(timeout=5)
    fim = round(time.time() - inicio, 3)
    t0 = None
    if (pasta / "t0.txt").exists():
        t0 = round(float((pasta / "t0.txt").read_text()) - inicio, 3)
    videos = sorted(pasta.glob("*.webm"))
    resumo = {"nome": nome, "codigo": codigo, "duracao": fim, "t_pagina": t0,
              "video": videos[-1].name if videos else None, "eventos": eventos, "entrada": entrada.strip()}
    (pasta / "resumo.json").write_text(json.dumps(resumo, ensure_ascii=False), encoding="utf-8")
    print("  %-12s codigo=%s duracao=%.1fs eventos=%d video=%s" % (nome, codigo, fim, len(eventos), resumo["video"]))
    if codigo not in (0, 1):
        raise SystemExit("o robô %s terminou com erro (%s)" % (nome, codigo))
    return resumo


def imagens(pedidos):
    for p in snapshot.snapshots(pedidos):
        print("  imagem", p.name)


def main():
    PASTA.mkdir(parents=True, exist_ok=True)
    print("1. Recomeçando a demonstração do zero...")
    subprocess.run([sys.executable, "montar_demo.py", "--sem-rotina"], cwd=str(RAIZ), env=ENV, check=True,
                   stdout=subprocess.DEVNULL)

    print("2. Rotina de terça: Drive -> pasta -> aba nova -> aviso ao gestor...")
    rodar("rotina", "rotina_terca.py")
    imagens([("s1_semana.png", "Inserção - 13.10", "B1:M26", [5, 7, 10, 11])])
    print("   O gestor valida no chat...")
    rodar("gestor", "gestor_responde.py", viewport="1000x560")

    print("3. Botão 1: conferir notas...")
    rodar("conferir", "conferir_nf.py")
    imagens([("s2_azul.png", "Conferência NF", "A1:J24", []),
             ("s2_laranja.png", "Conferência NF", "K5:R24", []),
             ("s2_decisao.png", "Conferência NF", "S5:AA24", [])])

    print("4. A pessoa decide as linhas em branco...")
    subprocess.run([sys.executable, "dados/aplicar_decisoes.py"], cwd=str(RAIZ), env=ENV, check=True)
    imagens([("s3_decisao.png", "Conferência NF", "S5:AA24", [])])

    print("5. Botão 2: simulação e baixa real...")
    rodar("baixa_sim", "dar_baixa.py", "SIMULAR\n")
    rodar("baixa_real", "dar_baixa.py", "BAIXAR\n")
    imagens([("s4_conferencia.png", "Conferência NF", "A1:J24", []),
             ("s4_execucao.png", "Conferência NF", "AB5:AD24", []),
             ("s4_completa.png", "Inserção - 13.10 completa", "A1:R27", [4, 5, 6, 7, 8, 10, 11, 12, 14, 15, 16, 17])])

    print("6. Botão 3: gerar o arquivo de importação...")
    rodar("gerar", "gerar_insercao.py")
    cfg.PLANILHA = cfg.ARQ_INSERCAO
    imagens([("s5_arquivo.png", "Inserção", "B1:M13", [4, 5, 6, 7, 8, 10, 11])])

    estado = subprocess.run([sys.executable, "-c", "import urllib.request;print(urllib.request.urlopen('http://127.0.0.1:8800/api/estado').read().decode())"],
                            capture_output=True, text=True).stdout.strip()
    (PASTA / "estado_final.json").write_text(estado, encoding="utf-8")
    print("Estado final do ERP:", estado)
    print("Pronto. Material em", PASTA)


if __name__ == "__main__":
    main()
