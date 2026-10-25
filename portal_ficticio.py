"""
Substituto da classe PortalUnico que responde com os dados de exemplos/portal_ficticio.json.

Imita o que importa da API real: a consulta de LPCOs filtra por data de registro e NCM
e devolve os resultados em paginas (parametro offset); o detalhe vem por numero; e as
exportacoes do catalogo chegam como .zip com um JSON dentro.
"""

import io
import json
import os
import zipfile

ARQ_EXEMPLO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exemplos", "portal_ficticio.json")


class PortalFicticio:
    def __init__(self, arquivo=ARQ_EXEMPLO, por_pagina=2):
        with open(arquivo, encoding="utf-8") as f:
            self.dados = json.load(f)
        self.por_pagina = por_pagina
        self.chamadas = []

    def get(self, caminho, params=None, sistema="talpco", binario=False):
        self.chamadas.append(caminho)
        params = params or {}
        if caminho == "/ext/lpco/consulta":
            return self._consulta(params)
        if caminho.startswith("/ext/lpco/"):
            det = self.dados["detalhes"].get(caminho.rsplit("/", 1)[-1])
            return (200, det) if det else (204, None)
        if sistema == "catp":
            return 200, self._exportacao(caminho)
        raise RuntimeError(f"GET {caminho} -> 404")

    def _consulta(self, params):
        ncm = params.get("ncm-inicial")
        achados = [
            l for l in self.dados["consulta"]
            if params["data-inicial-registro"] <= l["dataRegistro"] <= params["data-final-registro"]
            and (not ncm or any(it["ncm"] == ncm for it in self.dados["detalhes"][l["numero"]]["listaNcm"]))
        ]
        pagina = achados[int(params.get("offset", 0)):][: self.por_pagina]
        return (200, pagina) if pagina else (204, None)

    def _exportacao(self, caminho):
        if "/produto/" in caminho:
            registros = self.dados["produtos"]
        elif "/fabricante/" in caminho:
            registros = self.dados["vinculos"]
        else:
            registros = self.dados["operadores"]
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr("dados.json", json.dumps(registros, ensure_ascii=False))
        return buf.getvalue()

    def fechar(self):
        pass
