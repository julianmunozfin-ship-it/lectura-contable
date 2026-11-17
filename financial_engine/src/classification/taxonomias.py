"""Carga de taxonomías (PUC 2650, CUIF solidario, CUIF financiero) y detección del catálogo del balance."""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[2]
DIR_TAX = RAIZ / "data" / "taxonomies"

CATALOGOS = {
    "puc_2650": ("puc_2650.csv", "PUC comercial (Decreto 2650 de 1993)"),
    "cuif_solidario": ("cuif_solidario.csv", "CUIF Superintendencia de la Economía Solidaria"),
    "cuif_financiero": ("cuif_financiero_grupos.csv", "CUIF Superintendencia Financiera (grupos)"),
}
_VACIAS = {"DE", "LA", "EL", "Y", "LOS", "LAS", "DEL", "EN", "A", "POR", "O", "CON", "PARA", "Y/O"}


def normalizar(t: str) -> str:
    t = "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9/ ]", " ", t.upper())).strip()


def _tokens(t: str) -> set[str]:
    return {w for w in normalizar(t).split() if w not in _VACIAS and len(w) > 1}


def cargar_catalogos() -> dict[str, dict[str, str]]:
    out = {}
    for clave, (archivo, _) in CATALOGOS.items():
        ruta = DIR_TAX / archivo
        if ruta.exists():
            df = pd.read_csv(ruta, dtype=str, encoding="utf-8")
            out[clave] = dict(zip(df["codigo"], df["nombre"]))
    # Catálogos adicionales que el usuario agregue: data/taxonomies/catalogo_*.csv (codigo,nombre)
    for ruta in DIR_TAX.glob("catalogo_*.csv"):
        df = pd.read_csv(ruta, dtype=str, encoding="utf-8")
        out[ruta.stem.replace("catalogo_", "")] = dict(zip(df["codigo"], df["nombre"]))
    return out


def cargar_rubros() -> pd.DataFrame:
    return pd.read_csv(DIR_TAX / "rubros.csv", dtype={"familia": int, "signo": int}, encoding="utf-8").fillna("")


def detectar_catalogo(df: pd.DataFrame, catalogos: dict[str, dict[str, str]]) -> tuple[str, dict[str, float]]:
    """Compara los nombres del balance con los de cada catálogo, código por código.

    Usa los grupos (2 dígitos) si el balance los trae; si no, las cuentas (4 dígitos).
    Devuelve el catálogo con mayor similitud promedio, o 'desconocido' si ninguno pasa de 0.35.
    """
    puntajes = {}
    for clave, cat in catalogos.items():
        sims = []
        # Un libro auxiliar solo trae subcuentas (6 dígitos): en ese caso también se comparan.
        for nivel in (2, 4, 6):
            filas = df[df["codigo"].str.len() == nivel]
            for _, r in filas.iterrows():
                if r["codigo"] in cat:
                    a, b = _tokens(r["nombre"]), _tokens(cat[r["codigo"]])
                    if a and b:
                        sims.append(len(a & b) / len(a | b))
            if len(sims) >= 3:
                break
        cobertura = len(sims) / max(1, len(df[df["codigo"].str.len().isin([2, 4, 6])]))
        puntajes[clave] = round((sum(sims) / len(sims)) if sims else 0.0, 3) if cobertura > 0 else 0.0
    mejor = max(puntajes, key=puntajes.get) if puntajes else "desconocido"
    if puntajes and puntajes[mejor] >= 0.35:
        return mejor, puntajes
    # Plan B: los nombres no coinciden (p. ej. un auxiliar con descripciones propias del ERP).
    # Se decide por cuántos códigos de cuenta (4 dígitos) existen en cada catálogo.
    cuentas = {c[:4] for c in df["codigo"] if len(c) >= 4}
    if cuentas:
        cobertura = {k: sum(c in cat for c in cuentas) / len(cuentas) for k, cat in catalogos.items()}
        orden = sorted(cobertura, key=cobertura.get, reverse=True)
        segundo = cobertura[orden[1]] if len(orden) > 1 else 0.0
        if cobertura[orden[0]] >= 0.8 and cobertura[orden[0]] - segundo >= 0.2:
            puntajes["cobertura_codigos"] = round(cobertura[orden[0]], 3)
            return orden[0], puntajes
    return "desconocido", puntajes
