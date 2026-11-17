"""Clasificador híbrido de cuentas (reglas deterministas + Machine Learning + validación humana).

Orden de decisión para cada cuenta de detalle:
1. Memoria de validaciones humanas (mismo código y nombre ya validados antes).
2. Regla por código exacto de 4 o 6 dígitos del catálogo detectado.
3. Regla por palabra clave en el nombre.
4. Regla por grupo (2 dígitos) o clase del catálogo detectado.
5. Modelo de ML. Si la confianza es menor al umbral (0.85 por defecto), se pide
   validación humana en consola y la respuesta se guarda para reentrenar.
"""
from __future__ import annotations

import csv
import re
from datetime import date

import pandas as pd

from .modelo_ml import ETIQUETAS_HUMANAS, ModeloCuentas
from .reglas import CONF_CODIGO_ESPECIFICO, MotorReglas
from .taxonomias import cargar_catalogos, cargar_rubros, detectar_catalogo, normalizar


class ClasificadorHibrido:
    def __init__(self, umbral: float = 0.85, interactivo: bool = True):
        self.umbral = umbral
        self.interactivo = interactivo
        self.rubros = cargar_rubros()
        self.familia = dict(zip(self.rubros["rubro"], self.rubros["familia"]))
        self.descripcion = dict(zip(self.rubros["rubro"], self.rubros["descripcion"]))
        self.corriente_def = dict(zip(self.rubros["rubro"], self.rubros["corriente"]))
        self.catalogos = cargar_catalogos()
        self.reglas = MotorReglas()
        self.modelo = ModeloCuentas(self.rubros)
        self.modelo.entrenar(self.reglas.ejemplos_entrenamiento(self.catalogos))
        self.memoria = self._cargar_memoria()

    # ------------------------------------------------------------------ memoria humana
    @staticmethod
    def _llave(codigo: str, nombre: str) -> str:
        return f"{codigo}|{normalizar(nombre)}"

    def _cargar_memoria(self) -> dict[str, str]:
        if not ETIQUETAS_HUMANAS.exists():
            return {}
        h = pd.read_csv(ETIQUETAS_HUMANAS, dtype=str, encoding="utf-8")
        return {self._llave(r["codigo"], r["nombre"]): r["rubro"] for _, r in h.iterrows()}

    def _guardar_humana(self, codigo: str, nombre: str, rubro: str) -> None:
        nuevo = not ETIQUETAS_HUMANAS.exists()
        ETIQUETAS_HUMANAS.parent.mkdir(parents=True, exist_ok=True)
        with open(ETIQUETAS_HUMANAS, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if nuevo:
                w.writerow(["codigo", "nombre", "rubro", "fecha"])
            w.writerow([codigo, nombre, rubro, date.today().isoformat()])
        self.memoria[self._llave(codigo, nombre)] = rubro

    # ------------------------------------------------------------------ utilidades
    def _compatible(self, rubro: str, clase: int) -> bool:
        fam = 5 if clase in (6, 7) else clase
        return self.familia.get(rubro) == fam

    def _preguntar(self, fila, sugerencias: list[tuple[str, float]]) -> str:
        print("\n" + "─" * 78)
        print(f"Cuenta {fila['codigo']} – {fila['nombre']}")
        print(f"Saldo final: {fila['saldo_final']:,.2f}   (clase {fila['clase']})")
        print("El modelo no está seguro. Sugerencias:")
        for i, (r, p) in enumerate(sugerencias, 1):
            print(f"  {i}. {r:<20} {self.descripcion[r]}  ({p:.0%})")
        opciones = [r for r in self.rubros["rubro"] if self._compatible(r, fila["clase"])]
        while True:
            resp = input("Enter = opción 1 · número de la opción · código de rubro · L = listar rubros: ").strip()
            if resp == "" and sugerencias:
                return sugerencias[0][0]
            if resp.isdigit() and 1 <= int(resp) <= len(sugerencias):
                return sugerencias[int(resp) - 1][0]
            if resp.upper() == "L":
                for r in opciones:
                    print(f"     {r:<20} {self.descripcion[r]}")
                continue
            if resp.upper() in opciones:
                return resp.upper()
            print("  Respuesta no válida.")

    # ------------------------------------------------------------------ clasificación
    def clasificar(self, df: pd.DataFrame, catalogo: str | None = None) -> tuple[pd.DataFrame, str, dict]:
        if not catalogo or catalogo == "auto":
            catalogo, puntajes = detectar_catalogo(df, self.catalogos)
        else:
            puntajes = {}
        res = []
        for _, f in df.iterrows():
            if not f["es_detalle"]:
                res.append(("", "", None, "cuenta mayor (no se clasifica)", ""))
                continue
            if f["clase"] in (8, 9) or f["clase"] == 0:
                res.append(("ORDEN", "exclusión", 1.0, "cuenta de orden: no entra a los estados", ""))
                continue
            codigo, nombre, clase = f["codigo"], f["nombre"], int(f["clase"])
            elegido = None
            llave = self._llave(codigo, nombre)
            if llave in self.memoria and self._compatible(self.memoria[llave], clase):
                elegido = (self.memoria[llave], "humano (memoria)", 1.0, "validado antes por el usuario")
            if not elegido:
                r = self.reglas.por_codigo(codigo, catalogo)
                if r and r[1] >= CONF_CODIGO_ESPECIFICO and self._compatible(r[0], clase):
                    elegido = (r[0], "regla_codigo", r[1], r[2])
            if not elegido:
                r = self.reglas.por_nombre(nombre, clase)
                if r and self._compatible(r[0], clase):
                    elegido = (r[0], "regla_nombre", r[1], r[2])
            if not elegido:
                r = self.reglas.por_codigo(codigo, catalogo)
                if r and self._compatible(r[0], clase):
                    elegido = (r[0], "regla_codigo", r[1], r[2])
            alternativas = ""
            if not elegido:
                sug = self.modelo.predecir(nombre, clase)
                alternativas = "; ".join(f"{r} {p:.0%}" for r, p in sug)
                if sug and sug[0][1] >= self.umbral:
                    elegido = (sug[0][0], "ml", sug[0][1], "modelo TF-IDF + regresión logística")
                elif self.interactivo and sug:
                    rub = self._preguntar(f, sug)
                    self._guardar_humana(codigo, nombre, rub)
                    elegido = (rub, "humano", 1.0, "validado en esta corrida")
                elif sug:
                    elegido = (sug[0][0], "ml (REVISAR)", sug[0][1], "confianza menor al umbral")
                else:
                    fam_def = {1: "A_OTROS", 2: "P_OTROS", 3: "PT_RESERVAS", 4: "I_OTROS", 5: "G_OTROS_OPER"}
                    elegido = (fam_def[5 if clase in (6, 7) else clase], "por defecto (REVISAR)", 0.0, "sin regla ni modelo")
            res.append((*elegido, alternativas))
        df = df.copy()
        df["rubro"] = [r[0] for r in res]
        df["fuente"] = [r[1] for r in res]
        df["confianza"] = [r[2] for r in res]
        df["detalle_clasificacion"] = [r[3] for r in res]
        df["alternativas_ml"] = [r[4] for r in res]
        df["revisar"] = df["fuente"].str.contains("REVISAR", na=False)
        # corriente / no corriente: por defecto del rubro, ajustado por palabras del nombre
        def corriente(f):
            if f["rubro"] not in self.corriente_def or f["clase"] not in (1, 2):
                return ""
            n = normalizar(f["nombre"])
            if re.search(r"LARGO PLAZO|NO CORRIENTE", n):
                return "NC"
            if re.search(r"CORTO PLAZO|\bCORRIENTES?\b", n):
                return "C"
            return self.corriente_def[f["rubro"]] or "C"
        df["corriente"] = df.apply(corriente, axis=1)
        df["funcion"] = df.apply(
            lambda f: ("Ventas" if catalogo == "puc_2650" and f["codigo"].startswith("52") else "Administración")
            if f["clase"] in (5, 6, 7) and f["es_detalle"] else "", axis=1)
        return df, catalogo, puntajes
