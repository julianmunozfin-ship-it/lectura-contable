"""Lector de balances de prueba en Excel (.xlsx/.xls) o CSV.

Busca automáticamente la fila de encabezados y las columnas de código, nombre,
saldo inicial, débitos, créditos y saldo final, aunque tengan nombres distintos.

También acepta un libro auxiliar (un movimiento por línea, como lo exportan SAP,
Siigo, World Office o Helisa): si el archivo trae código, débito y crédito pero
no trae saldos, suma los movimientos por cuenta y arma el balance de prueba.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd

from .numeros import a_float

SINONIMOS = {
    "codigo": ["CODIGO", "CUENTA", "COD", "CODIGO CUENTA", "NO. CUENTA", "NUMERO DE CUENTA", "PUC", "CODIGO PUC"],
    "nombre": ["NOMBRE", "NOMBRE DE LA CUENTA", "DESCRIPCION", "NOMBRE CUENTA", "DETALLE", "CONCEPTO"],
    "saldo_inicial": ["SALDO INICIAL", "SALDO ANTERIOR", "INICIAL", "ANTERIOR"],
    "debitos": ["DEBITOS", "DEBITO", "DEBE", "MOVIMIENTO DEBITO", "MOV DEBITO", "DEBITO COP", "VALOR DEBITO"],
    "creditos": ["CREDITOS", "CREDITO", "HABER", "MOVIMIENTO CREDITO", "MOV CREDITO", "CREDITO COP", "VALOR CREDITO"],
    "saldo_final": ["SALDO FINAL", "NUEVO SALDO", "SALDO ACTUAL", "SALDO", "FINAL"],
}


def _norm(t) -> str:
    t = "" if t is None or (isinstance(t, float) and pd.isna(t)) else str(t)
    t = "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")
    t = t.replace("_", " ")
    return re.sub(r"\s+", " ", t.upper().strip())


def _mapear_encabezados(fila: list) -> dict[str, int]:
    textos = [_norm(x) for x in fila]
    mapa: dict[str, int] = {}
    # primero coincidencias exactas, luego parciales
    for exacta in (True, False):
        for campo, sinonimos in SINONIMOS.items():
            if campo in mapa:
                continue
            for i, t in enumerate(textos):
                if i in mapa.values() or not t:
                    continue
                if campo == "codigo" and not exacta:
                    coincide = any(re.search(rf"(^|\s){re.escape(s)}(\s|$)", t) for s in sinonimos) and "DOCUMENTO" not in t
                elif campo in ("debitos", "creditos") and not exacta:
                    coincide = any(re.search(rf"(^|\s){re.escape(s)}(\s|$)", t) for s in sinonimos) and "NATURALEZA" not in t
                else:
                    coincide = any((t == s) if exacta else (s in t) for s in sinonimos)
                if coincide:
                    mapa[campo] = i
                    break
    # "CUENTA" puede ser el nombre si hay otra columna de código
    return mapa


def _leer_crudo(ruta: Path) -> pd.DataFrame:
    if ruta.suffix.lower() == ".csv":
        for sep in (";", ",", "\t", "|"):
            for enc in ("utf-8-sig", "latin-1"):
                try:
                    df = pd.read_csv(ruta, sep=sep, header=None, dtype=str, encoding=enc)
                    if df.shape[1] >= 4:
                        return df
                except Exception:
                    continue
        raise ValueError("No se pudo leer el CSV (separador o codificación no reconocidos).")
    hojas = pd.read_excel(ruta, sheet_name=None, header=None, dtype=object)
    # la hoja con más filas que tengan un código numérico al inicio
    def puntaje(df):
        return sum(bool(re.match(r"^\d{1,10}$", re.sub(r"[.\s]", "", str(v)))) for v in df.iloc[:, :3].values.ravel())
    return max(hojas.values(), key=puntaje)


def leer_tabular(ruta: str | Path) -> tuple[list[dict], dict]:
    ruta = Path(ruta)
    df = _leer_crudo(ruta)
    meta: dict = {"empresa": "", "nit": "", "periodo": "", "advertencias": []}
    fila_enc, mapa = None, {}
    for i in range(min(40, len(df))):
        m = _mapear_encabezados(list(df.iloc[i].values))
        if len(m) >= 5 and {"codigo", "saldo_final"} <= set(m):
            fila_enc, mapa = i, m
            break
        texto = " ".join(_norm(x) for x in df.iloc[i].values if _norm(x))
        if texto and not meta["empresa"] and not texto.startswith("BALANCE"):
            meta["empresa"] = texto.title()
        n = re.search(r"NIT[:\s.]*([\d.\-]+)", texto)
        if n:
            meta["nit"] = n.group(1)
    if fila_enc is None:
        aux = _buscar_auxiliar(df)
        if aux is not None:
            return _leer_auxiliar(df, *aux, meta)
        primeras = [" | ".join(_norm(x) for x in df.iloc[i].values if _norm(x))[:150] for i in range(min(3, len(df)))]
        raise ValueError(
            "No se reconoció el formato. El archivo debe ser un balance de prueba (código, nombre, saldo inicial, "
            "débitos, créditos y saldo final) o un libro auxiliar (código de cuenta, débito y crédito por movimiento). "
            "Primeras filas leídas: " + " // ".join(primeras))
    if "nombre" not in mapa:
        # algunos sistemas llaman "CUENTA" al nombre y "CODIGO" al código
        raise ValueError("No se encontró la columna con el nombre de la cuenta.")
    for campo in ("saldo_inicial", "debitos", "creditos"):
        if campo not in mapa:
            meta["advertencias"].append(f"No se encontró la columna {campo}; se asume en 0")

    filas = []
    for _, r in df.iloc[fila_enc + 1:].iterrows():
        bruto = r.iloc[mapa["codigo"]]
        if isinstance(bruto, float) and pd.isna(bruto):
            continue
        if isinstance(bruto, (int, float)) and float(bruto).is_integer():
            codigo = str(int(bruto))
        else:
            codigo = re.sub(r"[.\s\-]", "", re.sub(r"\.0$", "", str(bruto).strip()))
        if not re.match(r"^\d{1,10}$", codigo):
            continue
        fila = {"codigo": codigo, "nombre": str(r.iloc[mapa["nombre"]]).strip()}
        for campo in ("saldo_inicial", "debitos", "creditos", "saldo_final"):
            v = r.iloc[mapa[campo]] if campo in mapa else 0
            fila[campo] = 0.0 if v is None or (isinstance(v, float) and pd.isna(v)) else a_float(v)
        filas.append(fila)
    return filas, meta


FECHA = ["FECHA", "FECHA CONTABILIZACION", "FECHA DOCUMENTO", "FECHA MOVIMIENTO"]
CLASES_CREDITO = {2, 3, 4, 9}


def _buscar_auxiliar(df: pd.DataFrame):
    """Fila de encabezados de un libro auxiliar: código + débito + crédito, sin saldos."""
    for i in range(min(40, len(df))):
        m = _mapear_encabezados(list(df.iloc[i].values))
        if {"codigo", "debitos", "creditos"} <= set(m):
            textos = [_norm(x) for x in df.iloc[i].values]
            fecha = next((j for j, t in enumerate(textos) if any(t == f or t.startswith(f) for f in FECHA)), None)
            return i, m, fecha
    return None


def _leer_auxiliar(df: pd.DataFrame, fila_enc: int, mapa: dict, col_fecha, meta: dict):
    # la empresa y el NIT solo pueden venir en las filas de título, antes de los encabezados
    meta["empresa"], meta["nit"] = "", ""
    for i in range(fila_enc):
        texto = " ".join(_norm(x) for x in df.iloc[i].values if _norm(x))
        if texto and not meta["empresa"] and not texto.startswith(("BALANCE", "LIBRO", "AUXILIAR")):
            meta["empresa"] = texto.title()
        n = re.search(r"NIT[:\s.]*([\d.\-]+)", texto)
        if n:
            meta["nit"] = n.group(1)
    datos = df.iloc[fila_enc + 1:]
    cuentas: dict[str, dict] = {}
    fechas = []
    n_mov = 0
    for _, r in datos.iterrows():
        bruto = r.iloc[mapa["codigo"]]
        if bruto is None or (isinstance(bruto, float) and pd.isna(bruto)):
            continue
        if isinstance(bruto, (int, float)) and float(bruto).is_integer():
            codigo = str(int(bruto))
        else:
            codigo = re.sub(r"[.\s\-]", "", re.sub(r"\.0$", "", str(bruto).strip()))
        if not re.match(r"^\d{1,10}$", codigo):
            continue
        deb = r.iloc[mapa["debitos"]]
        cre = r.iloc[mapa["creditos"]]
        deb = 0.0 if deb is None or (isinstance(deb, float) and pd.isna(deb)) or str(deb).strip() == "" else a_float(deb)
        cre = 0.0 if cre is None or (isinstance(cre, float) and pd.isna(cre)) or str(cre).strip() == "" else a_float(cre)
        nombre = str(r.iloc[mapa["nombre"]]).strip() if "nombre" in mapa else ""
        c = cuentas.setdefault(codigo, {"codigo": codigo, "nombre": nombre or f"Cuenta {codigo}",
                                        "saldo_inicial": 0.0, "debitos": 0.0, "creditos": 0.0})
        c["debitos"] += deb
        c["creditos"] += cre
        n_mov += 1
        if col_fecha is not None:
            f = pd.to_datetime(r.iloc[col_fecha], errors="coerce", dayfirst=not re.match(r"^\d{4}-", str(r.iloc[col_fecha])))
            if pd.notna(f):
                fechas.append(f)
    if not cuentas:
        raise ValueError("El libro auxiliar no trae movimientos con código de cuenta válido.")

    filas = []
    for codigo in sorted(cuentas):
        c = cuentas[codigo]
        c["debitos"], c["creditos"] = round(c["debitos"], 2), round(c["creditos"], 2)
        clase = int(codigo[0])
        c["saldo_final"] = round(c["creditos"] - c["debitos"] if clase in CLASES_CREDITO else c["debitos"] - c["creditos"], 2)
        filas.append(c)

    if fechas:
        meta["periodo"] = f"{min(fechas):%d/%m/%Y} a {max(fechas):%d/%m/%Y}"
    meta["tipo_archivo"] = "libro auxiliar"
    meta["advertencias"].append(
        f"Es un libro auxiliar: se sumaron {n_mov} movimientos en {len(filas)} cuentas. "
        "No trae saldos iniciales, así que se toman en cero y los estados muestran solo los movimientos del periodo.")
    td = sum(f["debitos"] for f in filas)
    tc = sum(f["creditos"] for f in filas)
    if round(td - tc, 2) != 0:
        meta["advertencias"].append(
            f"El auxiliar no cumple la partida doble: débitos {td:,.2f} y créditos {tc:,.2f} "
            f"(diferencia {td - tc:,.2f}). Los estados van a salir descuadrados hasta corregir el archivo.")
    return filas, meta
