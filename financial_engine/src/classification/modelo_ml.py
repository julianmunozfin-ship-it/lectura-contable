"""Capa de Machine Learning: clasifica cuentas atípicas por su nombre.

Modelo: TF-IDF (palabras + n-gramas de caracteres) + regresión logística, que entrega
probabilidades. Se entrena con las cuentas de los catálogos etiquetadas por las reglas
y con las validaciones humanas guardadas en data/training/etiquetas_humanas.csv.
Las probabilidades se restringen a los rubros compatibles con la clase de la cuenta.
"""
from __future__ import annotations

import hashlib
import json

import joblib
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import FeatureUnion, Pipeline

from .taxonomias import RAIZ, normalizar

DIR_MODELOS = RAIZ / "data" / "models"
ETIQUETAS_HUMANAS = RAIZ / "data" / "training" / "etiquetas_humanas.csv"


def _texto(nombre: str, clase: int) -> str:
    return f"CLASE{clase} {normalizar(nombre)}"


class ModeloCuentas:
    def __init__(self, rubros: pd.DataFrame):
        self.familia = dict(zip(rubros["rubro"], rubros["familia"]))
        self.pipeline: Pipeline | None = None

    def _datos(self, ejemplos: list[tuple[str, int, str]]):
        X = [_texto(n, c) for n, c, _ in ejemplos]
        y = [r for _, _, r in ejemplos]
        if ETIQUETAS_HUMANAS.exists():
            h = pd.read_csv(ETIQUETAS_HUMANAS, dtype=str, encoding="utf-8")
            for _, r in h.iterrows():
                if r["rubro"] in self.familia:
                    # las validaciones humanas pesan más: se repiten
                    X += [_texto(r["nombre"], int(r["codigo"][0]))] * 5
                    y += [r["rubro"]] * 5
        return X, y

    def entrenar(self, ejemplos: list[tuple[str, int, str]]) -> None:
        X, y = self._datos(ejemplos)
        huella = hashlib.md5(json.dumps([X, y]).encode()).hexdigest()[:12]
        ruta = DIR_MODELOS / f"modelo_cuentas_{huella}.joblib"
        if ruta.exists():
            self.pipeline = joblib.load(ruta)
            return
        self.pipeline = Pipeline([
            ("rasgos", FeatureUnion([
                ("palabras", TfidfVectorizer(ngram_range=(1, 2), sublinear_tf=True)),
                ("caracteres", TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), sublinear_tf=True)),
            ])),
            ("clf", LogisticRegression(C=20, max_iter=3000)),
        ])
        self.pipeline.fit(X, y)
        DIR_MODELOS.mkdir(parents=True, exist_ok=True)
        for viejo in DIR_MODELOS.glob("modelo_cuentas_*.joblib"):
            viejo.unlink()
        joblib.dump(self.pipeline, ruta)

    def predecir(self, nombre: str, clase: int, top: int = 3) -> list[tuple[str, float]]:
        """Devuelve los `top` rubros más probables, con probabilidad renormalizada dentro de la clase."""
        if self.pipeline is None:
            return []
        fam = 5 if clase in (6, 7) else clase
        proba = self.pipeline.predict_proba([_texto(nombre, clase)])[0]
        clases = self.pipeline.classes_
        mask = np.array([self.familia.get(c) == fam for c in clases])
        if not mask.any():
            return []
        p = np.where(mask, proba, 0.0)
        p = p / p.sum() if p.sum() > 0 else p
        orden = np.argsort(p)[::-1][:top]
        return [(clases[i], float(p[i])) for i in orden if p[i] > 0]
