"""Lector de balances de prueba en PDF (texto digital, no escaneado).

Estrategia:
1. Se leen los caracteres de cada página con su posición (pdfplumber).
2. Se reconstruyen las palabras siguiendo el orden del flujo del PDF. Así no se mezclan
   los caracteres cuando un nombre largo se monta sobre la columna de cifras.
3. Se agrupan las palabras por renglón (coordenada vertical).
4. Cada cifra se asigna a su columna (saldo inicial, débitos, créditos, saldo final)
   según la posición horizontal de los títulos del encabezado.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber

from .numeros import PATRON_CIFRA, a_float

COLUMNAS = ("saldo_inicial", "debitos", "creditos", "saldo_final")
_CIFRA_COMPLETA = re.compile(r"^(?:" + PATRON_CIFRA.pattern + r")[A-Za-z]{0,2}$")
_CODIGO = re.compile(r"^\d{1,10}$")


def _sin_tildes(t: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn").upper()


@dataclass
class Palabra:
    texto: str
    x0: float
    x1: float
    top: float


@dataclass
class Registro:
    codigo: str
    nombre_partes: list[str] = field(default_factory=list)
    valores: dict[str, float] = field(default_factory=dict)
    pagina: int = 0


def _palabras_en_orden_de_flujo(pagina) -> list[Palabra]:
    palabras: list[Palabra] = []
    actual: list[dict] = []

    def cerrar():
        if actual:
            texto = "".join(c["text"] for c in actual).strip()
            if texto:
                palabras.append(Palabra(texto, actual[0]["x0"], actual[-1]["x1"], actual[0]["top"]))
        actual.clear()

    for c in pagina.chars:
        if c["text"].isspace():
            cerrar()
            continue
        if actual:
            prev = actual[-1]
            ancho = max(prev["x1"] - prev["x0"], 1.0)
            mismo_renglon = abs(c["top"] - prev["top"]) < 2.0
            contiguo = -0.6 * ancho <= c["x0"] - prev["x1"] <= 0.6 * ancho
            if not (mismo_renglon and contiguo):
                cerrar()
        actual.append(c)
    cerrar()
    return palabras


def _renglones(palabras: list[Palabra], tolerancia: float = 3.0) -> list[list[Palabra]]:
    renglones: list[list[Palabra]] = []
    for p in sorted(palabras, key=lambda w: (round(w.top, 1), w.x0)):
        if renglones and abs(renglones[-1][0].top - p.top) <= tolerancia:
            renglones[-1].append(p)
        else:
            renglones.append([p])
    return renglones


def _detectar_columnas(renglon: list[Palabra]) -> dict[str, tuple[float, float]] | None:
    """Ubica las columnas numéricas a partir de un renglón de encabezado."""
    txt = [_sin_tildes(p.texto) for p in renglon]
    cols: dict[str, tuple[float, float]] = {}
    for i, t in enumerate(txt):
        p = renglon[i]
        if t in ("INICIAL", "ANTERIOR"):
            ini = renglon[i - 1] if i > 0 and txt[i - 1] == "SALDO" else p
            cols["saldo_inicial"] = (ini.x0, p.x1)
        elif t.startswith("DEBITO") or t == "DEBE":
            cols["debitos"] = (p.x0, p.x1)
        elif t.startswith("CREDITO") or t == "HABER":
            cols["creditos"] = (p.x0, p.x1)
        elif t in ("NUEVO", "FINAL", "ACTUAL"):
            sig = renglon[i + 1] if i + 1 < len(renglon) and txt[i + 1] == "SALDO" else p
            ant = renglon[i - 1] if i > 0 and txt[i - 1] == "SALDO" else p
            cols["saldo_final"] = (min(p.x0, ant.x0), max(p.x1, sig.x1))
    return cols if len(cols) == 4 else None


def _columna_mas_cercana(p: Palabra, columnas: dict[str, tuple[float, float]]) -> str:
    def distancia(rango):
        a, b = rango
        if p.x1 >= a and p.x0 <= b:  # se traslapan
            return 0.0, abs(p.x1 - b)
        return min(abs(p.x0 - b), abs(p.x1 - a)), abs(p.x1 - b)
    return min(columnas, key=lambda k: distancia(columnas[k]))


def _es_cifra(t: str) -> bool:
    return bool(_CIFRA_COMPLETA.match(t))


def leer_pdf(ruta: str | Path) -> tuple[list[dict], dict]:
    """Devuelve (filas, metadatos). Cada fila tiene codigo, nombre y las 4 columnas numéricas."""
    registros: list[Registro] = []
    meta: dict = {"empresa": "", "nit": "", "periodo": "", "advertencias": []}
    columnas = None
    with pdfplumber.open(str(ruta)) as pdf:
        if not any(p.chars for p in pdf.pages):
            raise ValueError("El PDF no tiene texto digital (parece escaneado). Conviértalo con OCR o use Excel/CSV.")
        for n_pag, pagina in enumerate(pdf.pages, start=1):
            for renglon in _renglones(_palabras_en_orden_de_flujo(pagina)):
                texto_renglon = " ".join(p.texto for p in sorted(renglon, key=lambda w: w.x0))
                norm = _sin_tildes(texto_renglon)
                cols = _detectar_columnas(sorted(renglon, key=lambda w: w.x0))
                if cols:
                    columnas = cols
                    continue
                if n_pag == 1 and not registros:
                    # Encabezado del informe: empresa, NIT y periodo
                    m = re.search(r"NIT[:\s.]*([\d.\-]+)", norm)
                    if m:
                        meta["nit"] = m.group(1).strip(".")
                        continue
                    m = re.search(r"(DE|DEL)\s+(\d{1,2}/\w{3,9}/\d{4}|\d{1,2}/\d{1,2}/\d{4}|\d{4}-\d{2}-\d{2})\s+(A|AL)\s+(\S+)", norm)
                    if m:
                        meta["periodo"] = f"{m.group(2)} a {m.group(4).strip(')')}"
                        continue
                    if not meta["empresa"] and not _CODIGO.match(renglon[0].texto) and "BALANCE" not in norm:
                        meta["empresa"] = texto_renglon.strip()
                        continue
                if re.search(r"PAGINA\s+\d+|BALANCE DE PRUEBA|ESTE INFORME SE ELABOR|^\(?NIT", norm):
                    continue
                if meta["empresa"] and norm.strip() == _sin_tildes(meta["empresa"]).strip():
                    continue
                palabras = sorted(renglon, key=lambda w: w.x0)
                cifras = [p for p in palabras if _es_cifra(p.texto)]
                textos = [p for p in palabras if not _es_cifra(p.texto)]
                if textos and _CODIGO.match(textos[0].texto) and textos[0].x0 < (min(c[0] for c in columnas.values()) if columnas else 1e9):
                    reg = Registro(codigo=textos[0].texto, pagina=n_pag)
                    reg.nombre_partes = [t.texto for t in textos[1:]]
                    registros.append(reg)
                elif registros:
                    reg = registros[-1]  # continuación del renglón anterior
                    reg.nombre_partes += [t.texto for t in textos]
                else:
                    continue
                if columnas:
                    for c in cifras:
                        col = _columna_mas_cercana(c, columnas)
                        if col in reg.valores:
                            meta["advertencias"].append(f"Cuenta {reg.codigo}: dos cifras en la columna {col}")
                        reg.valores[col] = a_float(c.texto)
                else:
                    for col, c in zip(COLUMNAS, cifras[-4:]):
                        reg.valores[col] = a_float(c.texto)

    filas = []
    for r in registros:
        nombre = re.sub(r"\s+", " ", " ".join(r.nombre_partes)).strip()
        faltan = [c for c in COLUMNAS if c not in r.valores]
        if faltan:
            meta["advertencias"].append(f"Cuenta {r.codigo} (pág. {r.pagina}): faltan columnas {faltan}; se asumen en 0")
        filas.append({"codigo": r.codigo, "nombre": nombre, **{c: r.valores.get(c, 0.0) for c in COLUMNAS}})
    return filas, meta
