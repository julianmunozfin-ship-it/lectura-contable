"""Motor contable: arma el Estado de Situación Financiera, el Estado de Resultados
(NIIF 18 para Grupo 1 o NIIF para las Pymes para Grupo 2) y el Estado de Flujos de
Efectivo por el método indirecto, conectados entre sí.

Todos los saldos se trabajan "por naturaleza": activos, gastos y costos positivos
cuando son deudores; pasivos, patrimonio e ingresos positivos cuando son acreedores.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

OP, INV, FIN, IMP, DISC = "Operación", "Inversión", "Financiación", "Impuesto a las ganancias", "Operaciones discontinuadas"
SECCIONES_ESF = ["Activo corriente", "Activo no corriente", "Pasivo corriente", "Pasivo no corriente", "Patrimonio"]
ORDEN_PYMES = [
    "Ingresos de actividades ordinarias", "Costo de ventas", "Otros ingresos", "Gastos de distribución",
    "Gastos de administración", "Otros gastos", "Ingresos financieros", "Costos financieros",
    "Participación en resultados de subsidiarias y asociadas", "Gasto por impuesto a las ganancias",
    "Resultado de operaciones discontinuadas",
]
# NIIF 18, párrafos 55-58 y 65: reclasificaciones para entidades con actividad principal especificada
RECLASIF_ACTIVIDAD = {
    "inversion": ["I_FINANCIEROS", "I_VALORACION_VR", "I_GAN_INVERSIONES", "I_DIVIDENDOS", "I_ARRENDAMIENTOS", "G_VALORACION_VR"],
    "financiacion": ["I_FINANCIEROS", "G_FINANCIEROS"],
}


@dataclass
class Estados:
    datos: pd.DataFrame
    rubros: pd.DataFrame
    grupo: int
    meta: dict
    esf: dict = field(default_factory=dict)          # seccion -> [(linea, valor)]
    er: list = field(default_factory=list)           # [(tipo, etiqueta, acum, mov)] tipo: titulo|linea|subtotal|total
    efe: list = field(default_factory=list)          # [(tipo, etiqueta, valor)]
    controles: list = field(default_factory=list)    # [(control, a, b)]
    cifras: dict = field(default_factory=dict)


def preparar_rubros(rubros: pd.DataFrame, actividad_principal: str) -> pd.DataFrame:
    r = rubros.copy()
    for rub in RECLASIF_ACTIVIDAD.get(actividad_principal, []):
        r.loc[r["rubro"] == rub, "categoria_niif18"] = OP
    orden = pd.DataFrame([{"rubro": "ORDEN", "descripcion": "Cuenta de orden (excluida de los estados)", "estado": "ORDEN",
                           "familia": 8, "signo": 0}])
    return pd.concat([r, orden], ignore_index=True).fillna("")


def preparar_datos(df: pd.DataFrame, rubros: pd.DataFrame, grupo: int) -> pd.DataFrame:
    d = df.copy()
    attrs = rubros.set_index("rubro")
    def a(col):
        return d.apply(lambda f: attrs.at[f["rubro"], col] if f["es_detalle"] and f["rubro"] in attrs.index else "", axis=1)
    for col in ("estado", "linea_esf", "categoria_niif18", "linea_er_g1", "linea_er_pymes", "linea_flujo",
                "categoria_flujo", "linea_flujo_contra", "no_niif"):
        d[col] = a(col)
    d["signo"] = d.apply(lambda f: int(attrs.at[f["rubro"], "signo"]) if f["es_detalle"] and f["rubro"] in attrs.index else 0, axis=1)
    d["movimiento"] = d.apply(lambda f: f["debitos"] - f["creditos"] if f["naturaleza"] == "D" else f["creditos"] - f["debitos"], axis=1)
    d["control_fila"] = (d["saldo_inicial"] + d["movimiento"] - d["saldo_final"]).round(2)
    d["linea_er_pymes"] = d.apply(
        lambda f: "Gastos de distribución" if f["linea_er_pymes"] == "Gastos de administración" and f.get("funcion") == "Ventas"
        else f["linea_er_pymes"], axis=1)
    d["linea_er"] = d["linea_er_g1"] if grupo == 1 else d["linea_er_pymes"]
    def seccion(f):
        if f["estado"] == "ESF":
            if f["clase"] == 1:
                return "Activo no corriente" if f["corriente"] == "NC" else "Activo corriente"
            if f["clase"] == 2:
                return "Pasivo no corriente" if f["corriente"] == "NC" else "Pasivo corriente"
            return "Patrimonio"
        return "Patrimonio" if f["estado"] == "ORI" else ""
    d["seccion_esf"] = d.apply(seccion, axis=1)
    es_res = d["es_detalle"] & d["estado"].isin(["ER", "ORI"])
    d["valor_er_acum"] = (d["signo"] * d["saldo_final"]).where(es_res, 0.0)
    d["valor_er_mov"] = (d["signo"] * d["movimiento"]).where(es_res, 0.0)
    def efecto(f):
        if not f["es_detalle"]:
            return 0.0
        if f["estado"] == "ESF":
            if f["rubro"] == "A_EFECTIVO":
                return 0.0
            return f["saldo_inicial"] - f["saldo_final"] if f["clase"] == 1 else f["saldo_final"] - f["saldo_inicial"]
        if f["estado"] == "ORI":
            return f["valor_er_mov"]
        if str(f["linea_flujo"]).startswith("Ajuste"):
            return -f["valor_er_mov"]
        return 0.0
    d["efecto_flujo"] = d.apply(efecto, axis=1)
    d["efecto_contra"] = d.apply(lambda f: -f["efecto_flujo"] if f["linea_flujo_contra"] else 0.0, axis=1)
    return d


def construir(df_clas: pd.DataFrame, rubros_base: pd.DataFrame, grupo: int, meta: dict,
              actividad_principal: str = "ninguna") -> Estados:
    rubros = preparar_rubros(rubros_base, actividad_principal)
    d = preparar_datos(df_clas, rubros, grupo)
    e = Estados(datos=d, rubros=rubros, grupo=grupo, meta=meta)
    hojas = d[d["es_detalle"]]
    orden_esf = list(dict.fromkeys(rubros["linea_esf"]))

    # ---------------- Estado de resultados
    er_rows = hojas[hojas["estado"] == "ER"]
    def suma(filtro):
        sub = er_rows[filtro]
        return float(sub["valor_er_acum"].sum()), float(sub["valor_er_mov"].sum())
    if grupo == 1:
        orden_g1 = list(dict.fromkeys(rubros["linea_er_g1"]))
        acum_tot = [0.0, 0.0]
        def bloque(cat, titulo):
            lineas = [l for l in orden_g1 if ((er_rows["categoria_niif18"] == cat) & (er_rows["linea_er"] == l)).any()]
            if not lineas:
                return 0.0, 0.0
            e.er.append(("titulo", titulo, None, None))
            ta = tm = 0.0
            for l in lineas:
                a_, m_ = suma((er_rows["categoria_niif18"] == cat) & (er_rows["linea_er"] == l))
                e.er.append(("linea", l, a_, m_, cat))
                ta += a_; tm += m_
            return ta, tm
        op = bloque(OP, "CATEGORÍA DE OPERACIÓN")
        acum_tot = [op[0], op[1]]
        e.er.append(("subtotal", "Resultado de operación", *acum_tot))
        inv = bloque(INV, "CATEGORÍA DE INVERSIÓN")
        acum_tot = [acum_tot[0] + inv[0], acum_tot[1] + inv[1]]
        e.er.append(("subtotal", "Resultado antes de financiación e impuestos a las ganancias", *acum_tot))
        fin = bloque(FIN, "CATEGORÍA DE FINANCIACIÓN")
        acum_tot = [acum_tot[0] + fin[0], acum_tot[1] + fin[1]]
        e.er.append(("subtotal", "Resultado antes de impuestos a las ganancias", *acum_tot))
        imp = bloque(IMP, "CATEGORÍA DE IMPUESTOS A LAS GANANCIAS")
        acum_tot = [acum_tot[0] + imp[0], acum_tot[1] + imp[1]]
        if (er_rows["categoria_niif18"] == DISC).any():
            e.er.append(("subtotal", "Resultado de operaciones continuadas", *acum_tot))
            disc = bloque(DISC, "CATEGORÍA DE OPERACIONES DISCONTINUADAS")
            acum_tot = [acum_tot[0] + disc[0], acum_tot[1] + disc[1]]
        e.er.append(("total", "RESULTADO DEL PERIODO", *acum_tot))
    else:
        tot = [0.0, 0.0]
        e.er.append(("titulo", "RESULTADOS DEL PERIODO", None, None))
        presentes = [l for l in ORDEN_PYMES if (er_rows["linea_er"] == l).any()]
        def lin(nombre):
            if nombre in presentes:
                a_, m_ = suma(er_rows["linea_er"] == nombre)
                e.er.append(("linea", nombre, a_, m_, ""))
                tot[0] += a_; tot[1] += m_
        for l in ORDEN_PYMES[:2]:
            lin(l)
        e.er.append(("subtotal", "Ganancia bruta", *tot))
        for l in ORDEN_PYMES[2:6]:
            lin(l)
        e.er.append(("subtotal", "Resultado de actividades de operación", *tot))
        for l in ORDEN_PYMES[6:9]:
            lin(l)
        e.er.append(("subtotal", "Ganancia (pérdida) antes de impuestos", *tot))
        for l in ORDEN_PYMES[9:]:
            lin(l)
        e.er.append(("total", "RESULTADO DEL PERIODO", *tot))
    resultado_acum = e.er[-1][2]
    resultado_mov = e.er[-1][3]
    ori = hojas[hojas["estado"] == "ORI"]
    if not ori.empty:
        e.er.append(("titulo", "OTRO RESULTADO INTEGRAL", None, None))
        for l in dict.fromkeys(ori["linea_er"]):
            sub = ori[ori["linea_er"] == l]
            e.er.append(("linea_ori", l, float(sub["valor_er_acum"].sum()), float(sub["valor_er_mov"].sum()), ""))
        e.er.append(("total", "RESULTADO INTEGRAL TOTAL", resultado_acum + float(ori["valor_er_acum"].sum()),
                     resultado_mov + float(ori["valor_er_mov"].sum())))

    # ---------------- Estado de situación financiera
    esf_rows = hojas[hojas["seccion_esf"] != ""]
    for s in SECCIONES_ESF:
        sub = esf_rows[esf_rows["seccion_esf"] == s]
        lineas = [l for l in orden_esf if l and (sub["linea_esf"] == l).any()]
        e.esf[s] = [(l, float(sub.loc[sub["linea_esf"] == l, "saldo_final"].sum())) for l in lineas]
    e.esf["Patrimonio"].append(("Resultado del periodo", resultado_acum))

    # ---------------- Estado de flujos de efectivo (indirecto)
    cat_linea = {}
    for _, r in rubros[rubros["estado"].isin(["ESF", "ORI"])].iterrows():
        if r["linea_flujo"] and r["categoria_flujo"] != "Efectivo":
            cat_linea.setdefault(r["linea_flujo"], r["categoria_flujo"])
    orden_flujo = list(dict.fromkeys([l for l in rubros["linea_flujo"] if l]))
    usados = set(hojas.loc[hojas["efecto_flujo"].abs() >= 0, "linea_flujo"]) | set(hojas["linea_flujo_contra"])
    usados = {l for l in usados if l and l != "Efectivo"}
    def valor_flujo(l):
        return float(hojas.loc[hojas["linea_flujo"] == l, "efecto_flujo"].sum()
                     + hojas.loc[hojas["linea_flujo_contra"] == l, "efecto_contra"].sum())
    ajustes = [l for l in orden_flujo if l in usados and l.startswith("Ajuste")]
    por_cat = {c: [l for l in orden_flujo if l in usados and cat_linea.get(l) == c] for c in (OP, INV, FIN)}
    e.efe.append(("titulo", "FLUJOS DE EFECTIVO DE ACTIVIDADES DE OPERACIÓN", None))
    e.efe.append(("linea", "Resultado del periodo", resultado_mov))
    tot = {OP: resultado_mov, INV: 0.0, FIN: 0.0}
    if ajustes:
        e.efe.append(("subtitulo", "Ajustes por partidas que no afectan el efectivo:", None))
        for l in ajustes:
            v = valor_flujo(l); tot[OP] += v
            e.efe.append(("linea", l.replace("Ajuste: ", "").capitalize(), v, l))
    if por_cat[OP]:
        e.efe.append(("subtitulo", "Cambios en el capital de trabajo:", None))
        for l in por_cat[OP]:
            v = valor_flujo(l); tot[OP] += v
            e.efe.append(("linea", l, v, l))
    e.efe.append(("subtotal", "Efectivo neto de actividades de operación", tot[OP]))
    for c, titulo, sub in ((INV, "FLUJOS DE EFECTIVO DE ACTIVIDADES DE INVERSIÓN", "Efectivo neto de actividades de inversión"),
                           (FIN, "FLUJOS DE EFECTIVO DE ACTIVIDADES DE FINANCIACIÓN", "Efectivo neto de actividades de financiación")):
        e.efe.append(("titulo", titulo, None))
        for l in por_cat[c]:
            v = valor_flujo(l); tot[c] += v
            e.efe.append(("linea", l, v, l))
        e.efe.append(("subtotal", sub, tot[c]))
    variacion = tot[OP] + tot[INV] + tot[FIN]
    ef = hojas[hojas["rubro"] == "A_EFECTIVO"]
    ef_ini, ef_fin = float(ef["saldo_inicial"].sum()), float(ef["saldo_final"].sum())
    e.efe.append(("total", "AUMENTO (DISMINUCIÓN) NETO DEL EFECTIVO", variacion))
    e.efe.append(("linea_fija", "Efectivo y equivalentes al inicio del periodo", ef_ini))
    e.efe.append(("total", "EFECTIVO Y EQUIVALENTES AL FINAL DEL PERIODO", ef_ini + variacion))

    # ---------------- Controles
    tot_a = sum(v for s in SECCIONES_ESF[:2] for _, v in e.esf[s])
    tot_p = sum(v for s in SECCIONES_ESF[2:4] for _, v in e.esf[s])
    tot_pt = sum(v for _, v in e.esf["Patrimonio"])
    det = hojas[hojas["clase"] <= 7]
    nat = lambda col, k: float(hojas.loc[hojas["clase"] == k, col].sum())
    ecuacion = lambda col: (nat(col, 1), nat(col, 2) + nat(col, 3) + nat(col, 4) - nat(col, 5) - nat(col, 6) - nat(col, 7))
    e.controles = [
        ("Partida doble: Σ débitos = Σ créditos (cuentas de detalle)", float(det["debitos"].sum()), float(det["creditos"].sum())),
        ("Filas donde inicial ± movimientos ≠ saldo final", float((d["control_fila"] != 0).sum()), 0.0),
        ("Ecuación contable al inicio del periodo", *ecuacion("saldo_inicial")),
        ("Ecuación contable al cierre del periodo", *ecuacion("saldo_final")),
        ("Estado de situación financiera: Activo = Pasivo + Patrimonio", tot_a, tot_p + tot_pt),
        ("Flujo de efectivo: variación neta = Δ efectivo del balance", variacion, ef_fin - ef_ini),
    ]
    e.cifras = {"resultado_acum": resultado_acum, "resultado_mov": resultado_mov, "total_activo": tot_a,
                "total_pasivo": tot_p, "total_patrimonio": tot_pt, "variacion_efectivo": variacion,
                "efectivo_inicial": ef_ini, "efectivo_final": ef_fin}
    return e
