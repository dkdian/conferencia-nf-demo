# -*- coding: utf-8 -*-
"""Garante que o ERP fictício está no ar (sobe em segundo plano se precisar)."""

import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import config as cfg

SERVIDOR = Path(__file__).resolve().parent / "servidor.py"


def no_ar() -> bool:
    try:
        with urllib.request.urlopen(cfg.URL_ERP + "/api/saude", timeout=2) as r:
            return r.status == 200
    except OSError:
        return False


def garantir_servidor() -> None:
    if no_ar():
        return
    print("  Subindo o ERP fictício em %s ..." % cfg.URL_ERP)
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen([sys.executable, str(SERVIDOR), "--porta", str(cfg.PORTA_ERP)],
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags)
    for _ in range(40):
        if no_ar():
            return
        time.sleep(0.25)
    raise SystemExit("Não consegui subir o ERP fictício em %s." % cfg.URL_ERP)


def reiniciar() -> None:
    """Devolve os lançamentos ao estado inicial (nenhum pago pela demonstração)."""
    garantir_servidor()
    req = urllib.request.Request(cfg.URL_ERP + "/api/reiniciar", data=b"{}", method="POST",
                                 headers={"Content-Type": "application/json"})
    urllib.request.urlopen(req, timeout=5).read()
