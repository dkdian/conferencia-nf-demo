# -*- coding: utf-8 -*-
"""
Faz o papel do gestor na demonstração: abre o chat, lê o aviso da rotina de terça e responde OK.

Na vida real quem faz isso é uma pessoa. Aqui o robô só existe para a demonstração poder rodar sem ninguém.

    python gestor_responde.py
"""

from __future__ import annotations

import sys
import time

from playwright.sync_api import sync_playwright

import config as cfg
import conferir_nf as tela
from erp_ficticio import garantir


def main() -> int:
    garantir.garantir_servidor()
    print("Abrindo o chat da equipe como gestor...")
    with sync_playwright() as p:
        navegador, pag = tela.abrir_navegador(p)
        try:
            pag.goto(cfg.URL_ERP + "/chat?hora=10:22", wait_until="domcontentloaded")
            pag.locator(".cartao").first.wait_for(timeout=10000)
            print("Aviso da rotina recebido: o gestor confere as informações.")
            time.sleep(4.0 * max(cfg.RITMO_S, 0.5))
            pag.locator("#texto").click()
            pag.keyboard.type("OK, informações validadas. Pode seguir com a conferência das notas.", delay=55)
            time.sleep(0.8)
            pag.locator("#enviar").click()
            pag.locator(".msg .balao").first.wait_for(timeout=8000)
            print("Resposta enviada: OK.")
            time.sleep(3.0 * max(cfg.RITMO_S, 0.5))
        finally:
            navegador.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
