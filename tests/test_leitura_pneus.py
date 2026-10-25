import pytest

from leitura_pneus import ler_medida_carga_lonas, ler_produto_catalogo, modelo_base


@pytest.mark.parametrize("texto, esperado", [
    ("PNEU 295/80R22.5 154/149M 18PR", ("295/80R22.5", "154/149M", "18PR")),
    ("PNEU 295/80 R 22,5 154/149 M 18 PR", ("295/80R22.5", "154/149M", "18PR")),
    ("205/55R16 91V XL", ("205/55R16", "91V XL", "")),
    ("LT225/75R16 115/112S 10PR", ("LT225/75R16", "115/112S", "10PR")),
    ("195/75R16C 107/105R", ("195/75R16C", "107/105R", "")),
    ("12.00R24 160/157K 20PR", ("12.00R24", "160/157K", "20PR")),
    ("31X10.50R15LT 109S", ("31X10.50R15LT", "109S", "")),
    ("18.4-30 14PR TT", ("18.4-30", "", "14PR")),
    ("275/80R22.5-18", ("275/80R22.5", "", "18PR")),
    ("175/75R1385T", ("175/75R13", "85T", "")),
])
def test_medida_carga_lonas(texto, esperado):
    assert ler_medida_carga_lonas(texto) == esperado


def test_sem_medida():
    assert ler_medida_carga_lonas("CAMARA DE AR PARA BICICLETA") == ("", "", "")


def test_procura_a_medida_no_segundo_texto():
    assert ler_medida_carga_lonas("PNEU NOVO", "MEDIDA 205/55R16 91V") == ("205/55R16", "91V", "")


def test_produto_completo():
    campos = ler_produto_catalogo(
        "PNEU NOVO RADIAL 295/80R22.5 154/149M 18PR",
        "PNEU NOVO, MARCA: MARCAUM, MODELO: MD200, REGISTRO INMETRO 001234/2025, "
        "CERTIFICADO: OCP-0042/25, FAMILIA INMETRO: 1A2B3C",
    )
    assert campos["medida"] == "295/80R22.5"
    assert campos["carga"] == "154/149M"
    assert campos["lonas"] == "18PR"
    assert campos["marca"] == "MARCAUM"
    assert campos["modelo"] == "MD200"
    assert campos["registro"] == "001234/2025"
    assert campos["certificado"] == "OCP-0042/25"
    assert campos["familia"] == "1A2B3C"


def test_xl_no_modelo_vai_para_a_carga():
    campos = ler_produto_catalogo("PNEU 205/55R16 91V", "MARCA: MARCATRES, MODELO: EV ECO XL")
    assert campos["carga"] == "91V XL"


def test_modelo_sem_a_palavra_modelo():
    campos = ler_produto_catalogo("13.6-38 14PR TDB120 TT", "")
    assert campos["modelo"] == "TDB120"


@pytest.mark.parametrize("modelo, esperado", [
    ("NEO CURVE P1 / 176720", "NEOCURVEP1"),
    ("MD200 (295/80R22.5)", "MD200"),
    ("EV ECO XL", "EVECO"),
    ("TR-PRO 500 7891234567890", "TRPRO500"),
])
def test_modelo_base(modelo, esperado):
    assert modelo_base(modelo) == esperado
