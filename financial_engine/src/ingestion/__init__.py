"""Ingesta: convierte cualquier balance de prueba (PDF, Excel o CSV) en una matriz contable homogénea."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from ..accounting_engine.esquemas import CuentaBalance
from .lector_pdf import leer_pdf
from .lector_tabular import leer_tabular

NIVELES = {1: "Clase", 2: "Grupo", 4: "Cuenta", 6: "Subcuenta"}
CLASES_CREDITO = {2, 3, 4, 9}


def _normalizar_codigos(df: pd.DataFrame) -> pd.DataFrame:
    """Catálogos que escriben la clase como 1000 y el grupo como 1100 se llevan a 1 y 11."""
    codigos = set(df["codigo"])
    if any(len(c) == 4 and c.endswith("000") for c in codigos):
        def corto(c: str) -> str:
            if len(c) == 4 and c.endswith("000"):
                return c[0]
            if len(c) == 4 and c.endswith("00"):
                return c[:2]
            return c
        df["codigo_original"] = df["codigo"]
        df["codigo"] = df["codigo"].map(corto)
    return df


def _detectar_convencion(df: pd.DataFrame) -> str:
    """'natural': saldos positivos según la naturaleza de la cuenta.
    'con_signo': saldos crédito con signo negativo (débito − crédito)."""
    hojas = df[df["es_detalle"] & df["clase"].isin([2, 3, 4]) & ((df["debitos"] - df["creditos"]).abs() > 0.005)]
    if hojas.empty:
        hojas = df[df["es_detalle"] & df["clase"].isin([2, 3, 4])]
        neg = (hojas["saldo_final"] < 0).sum()
        return "con_signo" if neg > len(hojas) / 2 else "natural"
    natural = ((hojas["saldo_inicial"] - hojas["debitos"] + hojas["creditos"] - hojas["saldo_final"]).abs() < 0.01).sum()
    con_signo = ((hojas["saldo_inicial"] + hojas["debitos"] - hojas["creditos"] - hojas["saldo_final"]).abs() < 0.01).sum()
    return "con_signo" if con_signo > natural else "natural"


def leer_balance(ruta: str | Path) -> tuple[pd.DataFrame, dict]:
    ruta = Path(str(ruta).strip().strip('"').strip("'"))
    if not ruta.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta}")
    ext = ruta.suffix.lower()
    if ext == ".pdf":
        filas, meta = leer_pdf(ruta)
    elif ext in (".xlsx", ".xlsm", ".xls", ".csv"):
        filas, meta = leer_tabular(ruta)
    else:
        raise ValueError(f"Formato no soportado: {ext}. Use PDF, Excel o CSV.")
    if not filas:
        raise ValueError("No se encontraron cuentas en el archivo.")

    validas = [CuentaBalance(**f).model_dump() for f in filas]
    df = pd.DataFrame(validas)
    df = _normalizar_codigos(df)
    dup = df[df.duplicated("codigo", keep="first")]
    if not dup.empty:
        meta["advertencias"].append(f"Códigos repetidos (se conserva el primero): {', '.join(dup['codigo'].unique())}")
        df = df.drop_duplicates("codigo", keep="first")
    df["nivel"] = df["codigo"].str.len().map(lambda n: NIVELES.get(n, "Auxiliar" if n > 6 else "Otro"))
    df["clase"] = df["codigo"].str[0].astype(int)
    codigos = df["codigo"].tolist()
    df["es_detalle"] = [not any(o != c and o.startswith(c) for o in codigos) for c in codigos]

    convencion = _detectar_convencion(df)
    meta["convencion_signos"] = convencion
    if convencion == "con_signo":
        mask = df["clase"].isin(CLASES_CREDITO)
        df.loc[mask, ["saldo_inicial", "saldo_final"]] *= -1
        meta["advertencias"].append("Los saldos crédito venían con signo negativo; se pasaron a saldos por naturaleza.")
    df["naturaleza"] = df["clase"].map(lambda k: "C" if k in CLASES_CREDITO else "D")
    meta["archivo"] = ruta.name
    meta.setdefault("empresa", "")
    return df.reset_index(drop=True), meta
