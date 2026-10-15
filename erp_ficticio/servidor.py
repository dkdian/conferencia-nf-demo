# -*- coding: utf-8 -*-
"""ERP de mentira para a demonstração. Só biblioteca padrão do Python.

    python erp_ficticio/servidor.py            sobe em http://127.0.0.1:8800
    python erp_ficticio/servidor.py --porta N

O estado fica na memória: ao reiniciar, os lançamentos voltam ao início (ou use POST /api/reiniciar).
"""

from __future__ import annotations

import argparse
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import modelo

PASTA_ESTATICA = Path(__file__).resolve().parent / "static"
TIPOS = {".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8",
         ".css": "text/css; charset=utf-8"}

TRAVA = threading.Lock()
ESTADO = {"tx": modelo.novo_estado(), "chat": []}


class Manipulador(BaseHTTPRequestHandler):
    def log_message(self, formato, *args):          # sem ruído no terminal
        pass

    def _json(self, corpo, status=200):
        dados = json.dumps(corpo, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(dados)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(dados)

    def _corpo(self) -> dict:
        tamanho = int(self.headers.get("Content-Length") or 0)
        try:
            return json.loads(self.rfile.read(tamanho) or b"{}")
        except ValueError:
            return {}

    def do_GET(self):
        url = urlparse(self.path)
        if url.path == "/api/saude":
            return self._json({"ok": True})
        if url.path == "/api/config":
            return self._json({"hoje": modelo.dados_demo.HOJE.isoformat(),
                               "contas": modelo.dados_demo.CONTAS, "formas": modelo.dados_demo.FORMAS})
        if url.path == "/api/periodo-fechado":
            data = (parse_qs(url.query).get("data") or [""])[0]
            return self._json({"closed": modelo.periodo_fechado(data)})
        if url.path == "/api/chat":
            with TRAVA:
                return self._json(ESTADO["chat"])
        if url.path == "/api/estado":
            with TRAVA:
                pagos = [t for t in ESTADO["tx"] if t["status"] == "PAID"]
                return self._json({"lancamentos": len(ESTADO["tx"]), "pagos": len(pagos),
                                   "codigos_pagos": sorted(t["code"] for t in pagos)})
        caminho = "index.html" if url.path in ("/", "/financeiro/transacoes") else url.path.lstrip("/")
        if url.path == "/chat":
            caminho = "chat.html"
        arquivo = (PASTA_ESTATICA / caminho).resolve()
        if PASTA_ESTATICA not in arquivo.parents and arquivo != PASTA_ESTATICA / "index.html":
            return self._json({"erro": "não encontrado"}, 404)
        if not arquivo.is_file():
            return self._json({"erro": "não encontrado"}, 404)
        dados = arquivo.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", TIPOS.get(arquivo.suffix, "application/octet-stream"))
        self.send_header("Content-Length", str(len(dados)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(dados)

    def do_POST(self):
        caminho = urlparse(self.path).path
        corpo = self._corpo()
        with TRAVA:
            if caminho == "/api/transacoes/listar":
                return self._json(modelo.buscar(
                    ESTADO["tx"], corpo.get("document_number", ""), corpo.get("start_date"),
                    corpo.get("end_date"), corpo.get("page", 1), corpo.get("page_size", 10)))
            if caminho == "/api/transacoes/pagar":
                resultado = modelo.pagar(ESTADO["tx"], corpo)
                return self._json(resultado, 200 if resultado["ok"] else 400)
            if caminho == "/api/chat":
                msg = {k: corpo.get(k) for k in ("autor", "papel", "tipo", "titulo", "fatos", "texto", "hora")}
                msg["id"] = len(ESTADO["chat"]) + 1
                ESTADO["chat"].append(msg)
                return self._json({"ok": True, "id": msg["id"]})
            if caminho == "/api/reiniciar":
                ESTADO["chat"] = []
                ESTADO["tx"] = modelo.novo_estado()
                return self._json({"ok": True})
        return self._json({"erro": "não encontrado"}, 404)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--porta", type=int, default=8800)
    args = ap.parse_args()
    servidor = ThreadingHTTPServer(("127.0.0.1", args.porta), Manipulador)
    print("ERP fictício em http://127.0.0.1:%d  (Ctrl+C para parar)" % args.porta, flush=True)
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
