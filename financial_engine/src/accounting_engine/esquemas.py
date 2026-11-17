"""Esquemas de datos validados con pydantic."""
from __future__ import annotations

from pydantic import BaseModel, Field, field_validator


class CuentaBalance(BaseModel):
    codigo: str = Field(pattern=r"^\d{1,10}$")
    nombre: str
    saldo_inicial: float = 0.0
    debitos: float = 0.0
    creditos: float = 0.0
    saldo_final: float = 0.0

    @field_validator("nombre")
    @classmethod
    def limpiar_nombre(cls, v: str) -> str:
        return " ".join(str(v).split())


class Clasificacion(BaseModel):
    rubro: str
    fuente: str  # memoria | regla_codigo | regla_nombre | ml | humano
    confianza: float = Field(ge=0, le=1)
    detalle: str = ""
