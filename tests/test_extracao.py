from datetime import date

import pytest

import extrair_catalogo
import extrair_lpcos
from portal_ficticio import PortalFicticio

CONFIG = {
    "cnpj": "12345678000190",
    "data_inicial": "2026-01-01",
    "data_final": "2026-06-30",
    "dias_por_janela": 30,
    "ncms": ["4011.10.00", "4011.20.90"],
}


@pytest.fixture
def cache_temporario(tmp_path, monkeypatch):
    monkeypatch.setattr(extrair_lpcos, "ARQ_CACHE", str(tmp_path / "cache.json"))
    return {}


def test_janelas_cobrem_o_periodo_sem_sobrepor():
    janelas = list(extrair_lpcos.janelas_de_datas(date(2026, 1, 1), date(2026, 3, 1), 30))
    assert janelas[0] == (date(2026, 1, 1), date(2026, 1, 30))
    assert janelas[-1][1] == date(2026, 3, 1)
    for (_, fim), (inicio, _) in zip(janelas, janelas[1:]):
        assert (inicio - fim).days == 1


def test_ncm_aceita_com_e_sem_pontos():
    assert extrair_lpcos.ncms_do_config({"ncms": ["4011.10.00", "40112090"]}) == ["40111000", "40112090"]


def test_ncm_invalida_interrompe():
    with pytest.raises(SystemExit):
        extrair_lpcos.ncms_do_config({"ncms": ["4011.10"]})


def test_lista_todas_as_paginas_sem_repetir():
    lpcos = extrair_lpcos.listar_lpcos(PortalFicticio(por_pagina=1), CONFIG)
    numeros = [l["numero"] for l in lpcos]
    assert len(numeros) == len(set(numeros)) == 8


def test_lpco_em_situacao_final_e_em_cache_nao_e_consultada_de_novo(cache_temporario):
    api = PortalFicticio()
    lpcos = extrair_lpcos.listar_lpcos(api, CONFIG)
    extrair_lpcos.detalhar_lpcos(api, lpcos, cache_temporario)
    api.chamadas.clear()

    extrair_lpcos.detalhar_lpcos(api, lpcos, cache_temporario)
    # Só a que está em análise pode ter mudado
    assert api.chamadas == ["/ext/lpco/I2600000103"]


def test_uma_linha_por_item_das_ncms_filtradas(cache_temporario):
    api = PortalFicticio()
    lpcos = extrair_lpcos.listar_lpcos(api, CONFIG)
    extrair_lpcos.detalhar_lpcos(api, lpcos, cache_temporario)
    linhas = extrair_lpcos.montar_linhas(lpcos, cache_temporario, extrair_lpcos.ncms_do_config(CONFIG))

    assert len(linhas) == 11
    # O item de NCM 4011.90.90 da LPCO 103 fica de fora do filtro
    assert all(l[10] in ("40111000", "40112090") for l in linhas)
    sem_produto = [l for l in linhas if l[0] == "I2600000107"]
    assert sem_produto[0][-1] == "item sem produto do catalogo"


def test_catalogo_aponta_divergencias(tmp_path, cache_temporario, monkeypatch):
    monkeypatch.setattr(extrair_catalogo, "PASTA", str(tmp_path))
    api = PortalFicticio()
    lpcos = extrair_lpcos.listar_lpcos(api, CONFIG)
    extrair_lpcos.detalhar_lpcos(api, lpcos, cache_temporario)
    arq = tmp_path / "lpcos.xlsx"
    extrair_lpcos.gerar_excel(extrair_lpcos.montar_linhas(lpcos, cache_temporario), str(arq), CONFIG)

    produtos = extrair_catalogo.produtos_das_lpcos(str(arq))
    catalogos = {"12345678": extrair_catalogo.baixar_catalogo(api, "12345678")}
    linhas = {l[0]: l for l in extrair_catalogo.montar_linhas(produtos, catalogos)}

    assert linhas[999][-1] == "produto nao encontrado no catalogo"
    assert "versao atual do catalogo diferente" in linhas[203][-1]
    assert "sem operador estrangeiro" in linhas[203][-1]
    assert linhas[202][8] == "OP0001\nOP0003"
    assert linhas[101][15:20] == ["295/80R22.5", "154/149M", "18PR", "MARCAUM", "MD200"]
