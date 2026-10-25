"""
Roda as duas extracoes com o portal ficticio, sem credencial nem acesso a internet.

Gera em saida_demo/ os mesmos dois Excel da versao real:
LPCOs_demo.xlsx (uma linha por item de LPCO) e Catalogo_demo.xlsx (um produto por linha).

Uso: python demo_sem_portal.py
"""

import os

import extrair_catalogo
import extrair_lpcos
from portal_ficticio import PortalFicticio

SAIDA = os.path.join(extrair_lpcos.PASTA, "saida_demo")

CONFIG = {
    "cnpj": "12345678000190",
    "ambiente": "validacao",
    "data_inicial": "2026-01-01",
    "data_final": "2026-06-30",
    "dias_por_janela": 30,
    "ncms": ["4011.10.00", "4011.20.90"],
}


def main():
    os.makedirs(SAIDA, exist_ok=True)
    # Cache e exportacoes do catalogo ficam na pasta da demo, longe dos arquivos reais
    extrair_lpcos.ARQ_CACHE = os.path.join(SAIDA, "cache_detalhes_lpco.json")
    extrair_catalogo.PASTA = SAIDA

    api = PortalFicticio()
    cache = {}
    lpcos = extrair_lpcos.listar_lpcos(api, CONFIG)
    extrair_lpcos.detalhar_lpcos(api, lpcos, cache)
    linhas = extrair_lpcos.montar_linhas(lpcos, cache, extrair_lpcos.ncms_do_config(CONFIG))
    arq_lpcos = os.path.join(SAIDA, "LPCOs_demo.xlsx")
    extrair_lpcos.gerar_excel(linhas, arq_lpcos, CONFIG)
    extrair_lpcos.log(f"{arq_lpcos}: {len(lpcos)} LPCOs, {len(linhas)} linhas")

    produtos = extrair_catalogo.produtos_das_lpcos(arq_lpcos)
    catalogos = {raiz: extrair_catalogo.baixar_catalogo(api, raiz) for raiz in {r for r, _ in produtos}}
    linhas_cat = extrair_catalogo.montar_linhas(produtos, catalogos)
    arq_catalogo = os.path.join(SAIDA, "Catalogo_demo.xlsx")
    extrair_catalogo.gerar_excel(linhas_cat, arq_catalogo, os.path.basename(arq_lpcos))
    extrair_lpcos.log(f"{arq_catalogo}: {len(linhas_cat)} produtos")


if __name__ == "__main__":
    main()
