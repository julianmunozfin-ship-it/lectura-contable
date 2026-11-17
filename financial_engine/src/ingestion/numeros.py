"""Limpieza de cifras contables en formato latino (1.224.609,81) o anglosajón (1,224,609.81)."""
from __future__ import annotations

import re

# Una cifra: signo opcional, paréntesis opcionales, dígitos con separadores y, opcionalmente, decimales.
# No exige espacio antes porque en algunos PDF el número queda pegado al nombre ("RAZONABLE335,584,769.09").
PATRON_CIFRA = re.compile(
    r"(?<![\d.,])\(?-?\$?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{1,2})\)?-?(?![\d])"
    r"|(?<![\d.,])\(?-?\$?\d+[.,]\d{1,2}\)?-?(?![\d])"
    r"|(?<![\d.,])\(?-?\$?\d{1,3}(?:[.,]\d{3})+\)?-?(?![\d.,])"
    r"|(?<![\w.,])-?0(?![\d.,])"
)


def a_float(texto: str | float | int | None) -> float:
    """Convierte una cifra en texto a float, detectando el separador decimal.

    Reglas:
    - Con '.' y ',' a la vez: el último separador que aparece es el decimal.
    - Con un solo tipo de separador: si va seguido de exactamente 3 dígitos (y se repite
      o hay más de 3 dígitos a la izquierda) es de miles; si va seguido de 1 o 2 dígitos, es decimal.
    - Negativos: signo '-', paréntesis (1.000) o signo al final 1.000-.
    """
    if texto is None:
        return 0.0
    if isinstance(texto, (int, float)):
        return float(texto)
    s = str(texto).strip().replace("$", "").replace(" ", "").replace(" ", "")
    s = s.rstrip("DdCcRr")  # sufijos como "370,152,937.51D" o "CR"
    if s in ("", "-", "--"):
        return 0.0
    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo, s = True, s[1:-1]
    if s.endswith("-"):
        negativo, s = True, s[:-1]
    if s.startswith("-"):
        negativo, s = (not negativo), s[1:]
    if "," in s and "." in s:
        decimal = "," if s.rfind(",") > s.rfind(".") else "."
    elif "," in s or "." in s:
        sep = "," if "," in s else "."
        partes = s.split(sep)
        if len(partes) > 2 or (len(partes[-1]) == 3 and len(partes) == 2 and len(partes[0]) > 0 and partes[0] != "0"):
            decimal = None  # separador de miles
        else:
            decimal = sep
    else:
        decimal = None
    if decimal is None:
        s = s.replace(",", "").replace(".", "")
    else:
        miles = "." if decimal == "," else ","
        s = s.replace(miles, "").replace(decimal, ".")
    valor = float(s)
    return -valor if negativo else valor


def extraer_cifras(texto: str) -> list[tuple[int, int, str]]:
    """Devuelve (inicio, fin, texto) de cada cifra encontrada en una línea."""
    return [(m.start(), m.end(), m.group(0)) for m in PATRON_CIFRA.finditer(texto)]
