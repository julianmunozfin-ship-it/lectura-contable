"""Capa determinista: reglas por código (según el catálogo) y por palabras clave del nombre."""
from __future__ import annotations

import re

import pandas as pd

from .taxonomias import DIR_TAX, normalizar

CONF_CODIGO_ESPECIFICO = 0.99  # regla para la cuenta exacta (4 o 6 dígitos)
CONF_NOMBRE = 0.97             # palabra clave en el nombre
CONF_GRUPO = 0.90              # regla por grupo (2 dígitos) o clase


class MotorReglas:
    def __init__(self):
        rc = pd.read_csv(DIR_TAX / "reglas_codigo.csv", dtype=str, encoding="utf-8")
        self.reglas_codigo: dict[str, dict[str, str]] = {}
        for _, r in rc.iterrows():
            self.reglas_codigo.setdefault(r["catalogo"], {})[r["prefijo"]] = r["rubro"]
        rn = pd.read_csv(DIR_TAX / "reglas_nombre.csv", dtype=str, encoding="utf-8")
        self.reglas_nombre = [
            ({int(c) for c in str(r["clases"]).split("|")}, re.compile(r["patron"]), r["rubro"])
            for _, r in rn.iterrows()
        ]

    def por_nombre(self, nombre: str, clase: int):
        n = normalizar(nombre)
        for clases, patron, rubro in self.reglas_nombre:
            if clase in clases and patron.search(n):
                return rubro, CONF_NOMBRE, f"palabra clave «{patron.pattern}»"
        return None

    def por_codigo(self, codigo: str, catalogo: str):
        reglas = self.reglas_codigo.get(catalogo, {})
        for largo in range(len(codigo), 0, -1):
            pref = codigo[:largo]
            if pref in reglas:
                conf = CONF_CODIGO_ESPECIFICO if largo >= 4 else CONF_GRUPO
                return reglas[pref], conf, f"código {pref} en {catalogo}"
        return None

    def ejemplos_entrenamiento(self, catalogos: dict[str, dict[str, str]]) -> list[tuple[str, int, str]]:
        """Etiqueta las cuentas de los catálogos con las reglas para entrenar el modelo de ML."""
        ejemplos = []
        for clave, cat in catalogos.items():
            for codigo, nombre in cat.items():
                clase = int(codigo[0])
                if clase not in (1, 2, 3, 4, 5, 6, 7):
                    continue
                r = self.por_nombre(nombre, clase) or self.por_codigo(codigo, clave)
                if r:
                    ejemplos.append((nombre, clase, r[0]))
        return ejemplos
