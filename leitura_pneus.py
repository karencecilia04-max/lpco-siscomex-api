"""
Leitura das caracteristicas de pneu a partir de textos livres
(descricao do Catalogo de Produtos e itens das invoices).

Campos: medida, indice de carga/velocidade, lonas (PR), marca, modelo,
registro INMETRO, certificado INMETRO e familia INMETRO.
"""

import re
import unicodedata


def sem_acento(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    return "".join(c for c in texto if not unicodedata.combining(c)).upper()


def so_alfanum(texto):
    return re.sub(r"[^0-9A-Z]", "", sem_acento(texto))


# ---------------------------------------------------------------------------
# Medida, indice de carga e lonas
# ---------------------------------------------------------------------------

# 295/80R22.5 | 295/80 R 22,5 | 205/50ZR17 | LT225/75R16 | 195/75R16C | 275/80R22 5
# e tambem colado no resto: 175/75R1385T | 235/75R17.518PR143/141JTL
RE_RADIAL = re.compile(
    r"(?<![0-9])(LT|P)?\s?(\d{3})\s?/\s?(\d{2})\s?(Z?R)\s?(\d{2})(?:\s?[.,]\s?(5)|\s(5)(?=\s))?(C(?![A-Z])|LT)?")
# 185R14 | 12.00R24 (sem perfil)
RE_RADIAL_SEM_PERFIL = re.compile(
    r"(?<![0-9/.])(LT|P)?(\d{3}|\d{1,2}[.,]\d{2})\s?(Z?R)\s?(\d{2})(?:[.,](5))?(C(?![A-Z])|LT)?")
# 31X10.50R15LT
RE_FLOTACAO = re.compile(r"(?<![0-9])(\d{2})\s?X\s?(\d{1,2})[.,](\d{1,2})\s?R\s?(\d{2})(LT)?")
# Diagonais/agricolas/moto: 16.9-24 | 10-16.5 | 11L-15 | 10.5/80-18 | 7.50-16 | 6.00-9 | 250/80-18
# | 80/100-14 | 1300-24 (= 13.00-24)
RE_DIAGONAL = re.compile(
    r"(?<![0-9/.])(\d{2}00|\d{1,3}(?:[.,]\d{1,2})?L?(?:/\d{2,3})?)\s?-\s?(\d{1,2}(?:[.,]5)?)(?![0-9])")

# 154/149M | 154/149 M | 91H | 93WXL | 107H XL | 169D/172A8
RE_CARGA = re.compile(
    r"(?<![0-9A-Z/.,])(\d{2,3}(?:/\d{2,3})?)\s?([A-HJ-NP-Z]\d?)\s?(XL)?(?=$|[^0-9A-Z]|TL|TT|RF)")
RE_LONAS = re.compile(r"(?<![0-9])(\d{1,2})\s?-?\s?PR(?![A-Z])")  # 16PR, 16 PR, 16-PR
RE_LONAS_HIFEN = re.compile(r"^-(\d{2})(?![0-9])")  # 275/80R22.5-18 = 18 lonas
FIM_TRECHO = re.compile(r",|;|\bMARCA\b|\bMODELO\b|\bCERTIFICADO\b|\bFAMILIA\b|\bREGISTRO\b| - ")


def ler_medida(texto):
    """Retorna (medida normalizada, posicao onde a medida termina) ou ("", -1)."""
    t = sem_acento(texto)
    m = RE_RADIAL.search(t)
    if m:
        prefixo, larg, perfil, r, aro, meio1, meio2, sufixo = m.groups()
        aro += ".5" if (meio1 or meio2) else ""
        return f"{prefixo or ''}{larg}/{perfil}{r}{aro}{sufixo or ''}", m.end()
    # Flotacao antes do radial sem perfil: senao 31X10.50R15 vira so 10.50R15
    m = RE_FLOTACAO.search(t)
    if m:
        a, b, c, aro, lt = m.groups()
        return f"{a}X{b}.{c}R{aro}{lt or ''}", m.end()
    m = RE_RADIAL_SEM_PERFIL.search(t)
    if m:
        prefixo, larg, r, aro, meio, sufixo = m.groups()
        return f"{prefixo or ''}{larg}{r}{aro}{'.5' if meio else ''}{sufixo or ''}", m.end()
    m = RE_DIAGONAL.search(t)
    if m:
        largura = m.group(1).replace(",", ".")
        if re.fullmatch(r"\d{2}00", largura):
            largura = f"{largura[:2]}.00"
        return f"{largura}-{m.group(2).replace(',', '.')}", m.end()
    return "", -1


def ler_medida_carga_lonas(*textos):
    """Procura a medida no primeiro texto que a tiver e le carga e lonas logo depois dela."""
    for texto in textos:
        medida, fim = ler_medida(texto)
        if not medida:
            continue
        t = sem_acento(texto)[fim:]
        corte = FIM_TRECHO.search(t)
        trecho = t[:corte.start()] if corte else t
        # Texto colado ("4PR99VTL"): separa numeros de letras para achar a carga
        separado = re.sub(r"(?<=\d)(?=[A-Z])|(?<=PR)(?=\d)", " ", trecho)
        carga = ""
        m = RE_CARGA.search(trecho) or RE_CARGA.search(separado)
        if m:
            xl = " XL" if (m.group(3) or re.search(r"\bXL\b", trecho)) else ""
            carga = f"{m.group(1)}{m.group(2)}{xl}"
        lonas = RE_LONAS.search(trecho) or RE_LONAS_HIFEN.search(trecho)
        return medida, carga, f"{int(lonas.group(1))}PR" if lonas else ""
    return "", "", ""


# ---------------------------------------------------------------------------
# Marca, modelo e dados INMETRO
# ---------------------------------------------------------------------------

PARADA = (r"(?=\s*(?:,|;|\s-\s|\.\s|\.?$|\n|\bCERTIFICADO\b|\bFAMILIA\b|\bREGISTRO\b|\bNUMERO\b"
          r"|\bMODELO\b|\bMARCA\b|\bSEM CAMARA|\bCOM CAMARA|\bREF\b|\bTIPO\b|\bCCLASSTRIB\b))")
RE_MARCA = re.compile(r"\bMARCA\s*:?\s*(.+?)" + PARADA)
RE_MODELO = re.compile(r"\bMODELO\s*:?\s*(.+?)" + PARADA)
RE_REGISTRO = re.compile(r"\bREGISTRO\b[^0-9]{0,45}?(\d{3,6}\s?/\s?\d{3,4})")
RE_CERTIFICADO = re.compile(
    r"\bCERTIFICADO\b(?:\s+DE\s+CONFORMIDADE)?(?:\s+INMETRO)?\s*:?\s*([0-9A-Z][0-9A-Z .\-/]*?)"
    r"(?=\s*(?:,|;|\s-\s|\.\s|\.?$|\bREGISTRO\b|\bFAMILIA\b|\bNUMERO\b))")
RE_FAMILIA = re.compile(r"\bFAMILIA(?:\s+INMETRO)?\s*:?\s*(?:\(F\d+\)\s*)?([0-9][0-9A-Z]{3,11})")

# Codigos que os cadastros costumam colar no fim do modelo (EAN, codigo do fornecedor, lote)
RE_CODIGO_SOLTO = re.compile(r"^(?:\d{4,}|(?=[A-Z0-9_]*\d)(?=[A-Z0-9_]*[A-Z])[A-Z0-9_]{9,})$")


def limpar(texto):
    return re.sub(r"\s+", " ", texto).strip(" .:-/")


def ler_marca(*textos):
    for texto in textos:
        m = RE_MARCA.search(sem_acento(texto))
        if m and limpar(m.group(1)):
            return limpar(m.group(1))
    return ""


def ler_modelo(*textos):
    for texto in textos:
        m = RE_MODELO.search(sem_acento(texto))
        if m and limpar(m.group(1)):
            return limpar(m.group(1))
    return ""


def modelo_base(modelo):
    """Modelo sem os codigos colados no fim: 'NEO CURVE P1 / 176720' -> 'NEOCURVEP1'.
    Tira tambem a medida entre parenteses e o XL (que entra no indice de carga)."""
    t = re.sub(r"\([^)]*\)", " ", sem_acento(modelo))
    t = re.split(r"\s/\s|\s-\s", t)[0]
    tokens = [x for x in re.split(r"\s+", t) if x and x != "XL"]
    while len(tokens) > 1 and RE_CODIGO_SOLTO.match(tokens[-1]):
        tokens.pop()
    return so_alfanum(" ".join(tokens))


def ler_inmetro(texto):
    t = sem_acento(texto)
    reg = RE_REGISTRO.search(t)
    cert = RE_CERTIFICADO.search(t)
    fam = RE_FAMILIA.search(t)
    return (
        re.sub(r"\s", "", reg.group(1)) if reg else "",
        limpar(cert.group(1)) if cert else "",
        fam.group(1) if fam else "",
    )


SIGLAS_TECNICAS = {"TT", "TL", "HD", "SB", "DOT", "EU", "M+S", "RF", "VF", "IF", "CFO"}


def modelo_no_trecho(texto):
    """Cadastros antigos sem a palavra MODELO: '13.6-38 14PR TDB120 TT' -> 'TDB120'."""
    medida, fim = ler_medida(texto)
    if not medida:
        return ""
    t = sem_acento(texto)[fim:]
    corte = FIM_TRECHO.search(t)
    t = re.sub(r"\([^)]*\)\S*", " ", t[:corte.start()] if corte else t)
    tokens = [x for x in re.split(r"[\s_]+", t) if x
              and x not in SIGLAS_TECNICAS
              and not RE_LONAS.fullmatch(x) and not re.fullmatch(r"-\d+|PR", x)
              and not RE_CARGA.fullmatch(x)]
    return " ".join(tokens)


def ajustar_xl(campos):
    """XL as vezes vem no modelo ('PS71 XL') e nao junto do indice de carga."""
    if campos["carga"] and "XL" not in campos["carga"] and re.search(r"\bXL\b", sem_acento(campos["modelo"])):
        campos["carga"] += " XL"
    return campos


def ler_produto_catalogo(denominacao, descricao):
    medida, carga, lonas = ler_medida_carga_lonas(denominacao, descricao)
    registro, certificado, familia = ler_inmetro(descricao or denominacao)
    if not (registro or certificado or familia):
        registro, certificado, familia = ler_inmetro(denominacao)
    modelo = ler_modelo(descricao, denominacao) or modelo_no_trecho(denominacao)
    return ajustar_xl({
        "medida": medida,
        "carga": carga,
        "lonas": lonas,
        "marca": ler_marca(denominacao, descricao),
        "modelo": modelo,
        "modelo_base": modelo_base(modelo),
        "registro": registro,
        "certificado": certificado,
        "familia": familia,
    })
