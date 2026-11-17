"""Generador del libro de Excel con los estados financieros conectados por fórmulas nativas.

Pestañas:
1. Balance_Prueba_Normalizado  matriz base limpia, clasificada y con fórmulas de control
2. Estado_Resultados           NIIF 18 (Grupo 1) o NIIF para las Pymes (Grupo 2)
3. Balance_General             Estado de situación financiera, conectado al resultado
4. Flujo_Efectivo              Método indirecto, conectado al resultado y al balance
5. Verificaciones              Semáforos de cuadre
6. Alertas                     Cuentas por revisar, partidas no NIIF y advertencias de lectura
7. Catalogo_Rubros             Tabla editable que conecta cada rubro con las líneas de los estados

Cada cifra de los estados es una fórmula (SUMIFS) sobre la pestaña 1. Si el usuario cambia
el rubro de una cuenta en la columna M de la pestaña 1, los estados se recalculan solos.
Las fórmulas se escriben con su valor ya calculado para que se vean aun sin recalcular.
"""
from __future__ import annotations

from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from ..accounting_engine.motor import SECCIONES_ESF, Estados

N = "Balance_Prueba_Normalizado"
COLS = [  # (encabezado, ancho, clave en el DataFrame o None si es fórmula)
    ("codigo_cuenta", 11, "codigo"), ("nombre_cuenta", 46, "nombre"), ("nivel", 9, "nivel"), ("clase", 6, "clase"),
    ("es_detalle", 8, "es_detalle"), ("naturaleza", 9, "naturaleza"), ("saldo_inicial", 19, "saldo_inicial"),
    ("debitos", 17, "debitos"), ("creditos", 17, "creditos"), ("saldo_final", 19, "saldo_final"),
    ("movimiento_periodo", 18, None), ("control_fila", 11, None), ("rubro", 18, "rubro"), ("corriente", 9, "corriente"),
    ("funcion", 13, "funcion"), ("estado", 8, None), ("seccion_esf", 18, None), ("linea_esf", 34, None),
    ("categoria_niif18", 16, None), ("linea_er", 34, None), ("signo_er", 7, None), ("valor_er_acumulado", 18, None),
    ("valor_er_periodo", 17, None), ("linea_flujo", 36, None), ("efecto_flujo", 16, None),
    ("linea_flujo_contra", 30, None), ("efecto_contra", 15, None), ("fuente_clasificacion", 17, "fuente"),
    ("confianza", 9, "confianza"), ("alternativas_ml", 28, "alternativas_ml"), ("detalle_clasificacion", 34, "detalle_clasificacion"),
]
L = {c[0]: xl_col_to_name(i) for i, c in enumerate(COLS)}
CAT_COLS = ["rubro", "descripcion", "estado", "familia", "linea_esf", "corriente", "categoria_niif18", "linea_er_g1",
            "linea_er_pymes", "signo", "linea_flujo", "categoria_flujo", "linea_flujo_contra", "no_niif"]
CL = {c: xl_col_to_name(i) for i, c in enumerate(CAT_COLS)}


class _Formatos:
    def __init__(self, wb):
        base = {"font_name": "Arial", "font_size": 10}
        num = '#,##0.00;[Red](#,##0.00);"-"'
        f = lambda **k: wb.add_format({**base, **k})
        self.titulo = f(bold=True, font_size=14, font_color="#1F3864")
        self.sub = f(italic=True, font_color="#595959")
        self.enc = f(bold=True, font_color="#FFFFFF", bg_color="#1F3864", border=1, text_wrap=True, valign="vcenter")
        self.txt = f()
        self.txt_b = f(bold=True)
        self.seccion = f(bold=True, font_color="#1F3864", top=1)
        self.num = f(num_format=num)
        self.num_in = f(num_format=num, font_color="#0000FF")
        self.num_b = f(num_format=num, bold=True, top=1)
        self.num_tot = f(num_format=num, bold=True, top=1, bottom=6)
        self.txt_tot = f(bold=True, top=1, bottom=6)
        self.txt_sub = f(bold=True, top=1)
        self.editable = f(bg_color="#FFF2CC", font_color="#0000FF")
        self.pct = f(num_format="0%")
        self.ok = f(bold=True, bg_color="#C6EFCE", font_color="#006100", align="center")
        self.mal = f(bold=True, bg_color="#FFC7CE", font_color="#9C0006", align="center")
        self.nota = f(italic=True, font_size=9, font_color="#595959", text_wrap=True, valign="top")
        self.sangria = f(indent=1)
        self.sangria2 = f(indent=2)


def _encabezado(ws, F, e: Estados, titulo: str, fila: int = 0) -> int:
    m = e.meta
    ws.write(fila, 0, m.get("empresa") or "Empresa", F.titulo)
    linea = "  ·  ".join(x for x in [f"NIT {m['nit']}" if m.get("nit") else "", f"Periodo: {m['periodo']}" if m.get("periodo") else ""] if x)
    ws.write(fila + 1, 0, titulo, F.txt_b)
    ws.write(fila + 2, 0, (linea + "  ·  " if linea else "") + "Cifras en pesos colombianos", F.sub)
    return fila + 4


def exportar(e: Estados, ruta_salida: str | Path) -> Path:
    ruta_salida = Path(ruta_salida)
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(ruta_salida))
    F = _Formatos(wb)
    hojas = {nombre: wb.add_worksheet(nombre) for nombre in
             [N, "Estado_Resultados", "Balance_General", "Flujo_Efectivo", "Verificaciones", "Alertas", "Catalogo_Rubros"]}
    d = e.datos.reset_index(drop=True)
    n = len(d)
    r0, r1 = 5, 5 + n - 1          # filas de datos en Excel (1-based): 6 .. 5+n
    rng = lambda col: f"{N}!${L[col]}${r0 + 1}:${L[col]}${r1 + 1}"
    ncat = len(e.rubros)
    crng = lambda col: f"Catalogo_Rubros!${CL[col]}$5:${CL[col]}${4 + ncat}"

    # ------------------------------------------------------------ 1. Balance de prueba normalizado
    ws = hojas[N]
    _encabezado(ws, F, e, "Balance de prueba normalizado y clasificado")
    ws.write(3, 0, "En azul: valores leídos del archivo. Amarillo: columnas editables (rubro, corriente, función). "
                   "Si cambia un rubro, los estados se recalculan. Clases 8 y 9 (cuentas de orden) no entran a los estados.", F.nota)
    for j, (h, w, _) in enumerate(COLS):
        ws.write(4, j, h, F.enc)
        ws.set_column(j, j, w)
    ws.set_row(4, 30)
    linea_er_col = "linea_er_g1" if e.grupo == 1 else "linea_er_pymes"
    for i, f in d.iterrows():
        r = r0 + i            # índice 0-based
        x = r + 1             # número de fila Excel
        det = bool(f["es_detalle"])
        vals = {
            "codigo": str(f["codigo"]), "nombre": f["nombre"], "nivel": f["nivel"], "clase": int(f["clase"]),
            "es_detalle": "Sí" if det else "No", "naturaleza": f["naturaleza"],
        }
        for j, (h, _, k) in enumerate(COLS):
            c = L[h]
            if k in vals:
                ws.write_string(r, j, vals[k]) if isinstance(vals[k], str) else ws.write_number(r, j, vals[k])
            elif k in ("saldo_inicial", "debitos", "creditos", "saldo_final"):
                ws.write_number(r, j, float(f[k]), F.num_in)
            elif k in ("rubro", "corriente", "funcion"):
                ws.write_string(r, j, str(f[k] or ""), F.editable if det else F.txt)
            elif k == "confianza":
                if f[k] is not None and f[k] == f[k]:
                    ws.write_number(r, j, float(f[k]), F.pct)
            elif k:
                ws.write_string(r, j, str(f[k] or ""))
        R = lambda col: f"{L[col]}{x}"
        look = lambda col, defecto='""': (f'=IF(OR({R("es_detalle")}<>"Sí",{R("rubro")}=""),{defecto},'
                                          f'IFERROR(INDEX({crng(col)},MATCH({R("rubro")},{crng("rubro")},0)),"RUBRO NO VÁLIDO"))')
        fm = {
            "movimiento_periodo": (f'=IF({R("naturaleza")}="D",{R("debitos")}-{R("creditos")},{R("creditos")}-{R("debitos")})', f["movimiento"], F.num),
            "control_fila": (f'=ROUND({R("saldo_inicial")}+{R("movimiento_periodo")}-{R("saldo_final")},2)', f["control_fila"], F.num),
            "estado": (look("estado"), f["estado"], None),
            "seccion_esf": (f'=IF({R("estado")}="ESF",IF({R("clase")}=1,IF({R("corriente")}="NC","Activo no corriente","Activo corriente"),'
                            f'IF({R("clase")}=2,IF({R("corriente")}="NC","Pasivo no corriente","Pasivo corriente"),"Patrimonio")),'
                            f'IF({R("estado")}="ORI","Patrimonio",""))', f["seccion_esf"], None),
            "linea_esf": (look("linea_esf"), f["linea_esf"], None),
            "categoria_niif18": (look("categoria_niif18"), f["categoria_niif18"], None),
            "linea_er": ((f'=IF(AND({look(linea_er_col)[1:]}="Gastos de administración",{R("funcion")}="Ventas"),"Gastos de distribución",{look(linea_er_col)[1:]})'
                          if e.grupo == 2 else look(linea_er_col)), f["linea_er"], None),
            "signo_er": (look("signo", "0"), f["signo"], None),
            "valor_er_acumulado": (f'=IF(AND({R("es_detalle")}="Sí",OR({R("estado")}="ER",{R("estado")}="ORI")),{R("signo_er")}*{R("saldo_final")},0)', f["valor_er_acum"], F.num),
            "valor_er_periodo": (f'=IF(AND({R("es_detalle")}="Sí",OR({R("estado")}="ER",{R("estado")}="ORI")),{R("signo_er")}*{R("movimiento_periodo")},0)', f["valor_er_mov"], F.num),
            "linea_flujo": (look("linea_flujo"), f["linea_flujo"], None),
            "efecto_flujo": (f'=IF({R("es_detalle")}<>"Sí",0,IF({R("estado")}="ESF",IF({R("rubro")}="A_EFECTIVO",0,'
                             f'IF({R("clase")}=1,{R("saldo_inicial")}-{R("saldo_final")},{R("saldo_final")}-{R("saldo_inicial")})),'
                             f'IF({R("estado")}="ORI",{R("valor_er_periodo")},IF(LEFT({R("linea_flujo")},6)="Ajuste",-{R("valor_er_periodo")},0))))',
                             f["efecto_flujo"], F.num),
            "linea_flujo_contra": (look("linea_flujo_contra"), f["linea_flujo_contra"], None),
            "efecto_contra": (f'=IF({R("linea_flujo_contra")}="",0,-{R("efecto_flujo")})', f["efecto_contra"], F.num),
        }
        for h, (formula, valor, fmt) in fm.items():
            j = [c[0] for c in COLS].index(h)
            v = valor if isinstance(valor, (int, float)) else str(valor or "")
            ws.write_formula(r, j, formula, fmt, v)
    idx = {c[0]: j for j, c in enumerate(COLS)}
    ws.data_validation(r0, idx["rubro"], r1, idx["rubro"], {"validate": "list", "source": f"=Catalogo_Rubros!$A$5:$A${4 + ncat}"})
    ws.data_validation(r0, idx["corriente"], r1, idx["corriente"], {"validate": "list", "source": ["C", "NC"]})
    ws.data_validation(r0, idx["funcion"], r1, idx["funcion"], {"validate": "list", "source": ["Administración", "Ventas"]})
    ws.freeze_panes(5, 2)
    ws.autofilter(4, 0, r1, len(COLS) - 1)

    # ------------------------------------------------------------ 7. Catálogo de rubros
    wc = hojas["Catalogo_Rubros"]
    wc.write(0, 0, "Catálogo de rubros: conecta cada tipo de cuenta con su línea en cada estado", F.titulo)
    wc.write(1, 0, "Puede editar estas columnas para cambiar la presentación. La categoría NIIF 18 ya incluye el ajuste por actividad "
                   "principal elegido en config.yaml.", F.nota)
    for j, h in enumerate(CAT_COLS):
        wc.write(3, j, h, F.enc)
        wc.set_column(j, j, 30 if "linea" in h or h == "descripcion" else 14)
    for i, (_, rr) in enumerate(e.rubros.iterrows()):
        for j, h in enumerate(CAT_COLS):
            v = rr.get(h, "")
            if h in ("familia", "signo") and v != "":
                wc.write_number(4 + i, j, int(v))
            else:
                wc.write_string(4 + i, j, str(v))
    wc.freeze_panes(4, 1)

    # ------------------------------------------------------------ 2. Estado de resultados
    we = hojas["Estado_Resultados"]
    titulo_er = ("Estado del resultado del periodo (NIIF 18 – Grupo 1)" if e.grupo == 1
                 else "Estado de resultados por función del gasto (NIIF para las Pymes, Sección 5 – Grupo 2)")
    fila = _encabezado(we, F, e, titulo_er)
    we.set_column(0, 0, 62); we.set_column(1, 2, 22)
    we.write(fila, 0, "Concepto", F.enc); we.write(fila, 1, "Acumulado al corte", F.enc); we.write(fila, 2, "Movimiento del periodo", F.enc)
    fila += 1
    inicio_bloque, filas_lineas, filas_sub = None, [], []
    fila_resultado = fila_resultado_integral = None
    for item in e.er:
        tipo, etiqueta, a, m = item[:4]
        x = fila + 1
        if tipo == "titulo":
            we.write(fila, 0, etiqueta, F.seccion)
        elif tipo in ("linea", "linea_ori"):
            cat = item[4] if len(item) > 4 else ""
            crit_cat = f',{rng("categoria_niif18")},"{cat}"' if (e.grupo == 1 and cat) else ""
            crit_est = f',{rng("estado")},"{"ORI" if tipo == "linea_ori" else "ER"}"'
            we.write(fila, 0, etiqueta, F.sangria)
            we.write_formula(fila, 1, f'=SUMIFS({rng("valor_er_acumulado")},{rng("linea_er")},"{etiqueta}"{crit_cat}{crit_est})', F.num, a)
            we.write_formula(fila, 2, f'=SUMIFS({rng("valor_er_periodo")},{rng("linea_er")},"{etiqueta}"{crit_cat}{crit_est})', F.num, m)
            filas_lineas.append((x, tipo))
        elif tipo in ("subtotal", "total"):
            es_integral = etiqueta == "RESULTADO INTEGRAL TOTAL"
            if es_integral:
                refs = [str(fila_resultado)] + [str(r) for r, t in filas_lineas if t == "linea_ori"]
            else:
                refs = [str(r) for r, t in filas_lineas if t == "linea"]
            fa = "=" + ("+".join(f"B{r}" for r in refs) if refs else "0")
            fmv = "=" + ("+".join(f"C{r}" for r in refs) if refs else "0")
            ft, fn = (F.txt_tot, F.num_tot) if tipo == "total" else (F.txt_sub, F.num_b)
            we.write(fila, 0, etiqueta, ft)
            we.write_formula(fila, 1, fa, fn, a)
            we.write_formula(fila, 2, fmv, fn, m)
            if etiqueta == "RESULTADO DEL PERIODO":
                fila_resultado = x
            if es_integral:
                fila_resultado_integral = x
        fila += 1
    fila += 1
    nota_er = ("Presentación según NIIF 18 (vigente para periodos que comiencen desde el 1 de enero de 2027; aplicación anticipada). "
               "Los ingresos y gastos se clasifican en las categorías de operación, inversión, financiación, impuestos a las ganancias "
               "y operaciones discontinuadas, con los subtotales obligatorios del párrafo 69."
               if e.grupo == 1 else
               "Presentación según la NIIF para las Pymes (Sección 5). El resultado del periodo pasa al patrimonio en el balance general.")
    we.merge_range(fila, 0, fila + 2, 2, nota_er + " «Acumulado al corte» usa los saldos finales de ingresos y gastos; "
                   "«Movimiento del periodo» usa solo los débitos y créditos del periodo del balance de prueba y es la base del flujo de efectivo.", F.nota)

    # ------------------------------------------------------------ 3. Balance general
    wg = hojas["Balance_General"]
    fila = _encabezado(wg, F, e, "Estado de situación financiera (balance general clasificado)")
    wg.set_column(0, 0, 62); wg.set_column(1, 1, 24)
    wg.write(fila, 0, "Concepto", F.enc); wg.write(fila, 1, "Saldo al corte", F.enc)
    fila += 1
    totales = {}
    titulos = {"Activo corriente": "ACTIVO", "Pasivo corriente": "PASIVO", "Patrimonio": "PATRIMONIO"}
    for s in SECCIONES_ESF:
        if s in titulos:
            wg.write(fila, 0, titulos[s], F.seccion); fila += 1
        if s != "Patrimonio":
            wg.write(fila, 0, s, F.txt_b); fila += 1
        primera = fila + 1
        for etiqueta, v in e.esf[s]:
            wg.write(fila, 0, etiqueta, F.sangria2 if s != "Patrimonio" else F.sangria)
            if etiqueta == "Resultado del periodo":
                wg.write_formula(fila, 1, f"=Estado_Resultados!B{fila_resultado}", F.num, v)
            else:
                wg.write_formula(fila, 1, f'=SUMIFS({rng("saldo_final")},{rng("seccion_esf")},"{s}",{rng("linea_esf")},"{etiqueta}",{rng("es_detalle")},"Sí")', F.num, v)
            fila += 1
        ultima = fila
        val = sum(v for _, v in e.esf[s])
        wg.write(fila, 0, f"Total {s.lower()}", F.txt_sub)
        wg.write_formula(fila, 1, f"=SUM(B{primera}:B{ultima})" if e.esf[s] else "=0", F.num_b, val)
        totales[s] = (fila + 1, val); fila += 1
        if s == "Activo no corriente":
            wg.write(fila, 0, "TOTAL ACTIVO", F.txt_tot)
            wg.write_formula(fila, 1, f"=B{totales['Activo corriente'][0]}+B{totales['Activo no corriente'][0]}", F.num_tot, e.cifras["total_activo"])
            totales["TA"] = fila + 1; fila += 2
        if s == "Pasivo no corriente":
            wg.write(fila, 0, "TOTAL PASIVO", F.txt_tot)
            wg.write_formula(fila, 1, f"=B{totales['Pasivo corriente'][0]}+B{totales['Pasivo no corriente'][0]}", F.num_tot, e.cifras["total_pasivo"])
            totales["TP"] = fila + 1; fila += 2
    wg.write(fila, 0, "TOTAL PASIVO Y PATRIMONIO", F.txt_tot)
    wg.write_formula(fila, 1, f"=B{totales['TP']}+B{totales['Patrimonio'][0]}", F.num_tot, e.cifras["total_pasivo"] + e.cifras["total_patrimonio"])
    totales["TPP"] = fila + 1; fila += 2
    wg.write(fila, 0, "Control: Activo − (Pasivo + Patrimonio)", F.txt_b)
    wg.write_formula(fila, 1, f"=ROUND(B{totales['TA']}-B{totales['TPP']},2)", F.num, 0)
    wg.write_formula(fila, 2, f'=IF(B{fila + 1}=0,"OK","DESCUADRADO")', F.txt_b, "OK")
    wg.conditional_format(fila, 2, fila, 2, {"type": "cell", "criteria": "==", "value": '"OK"', "format": F.ok})
    wg.conditional_format(fila, 2, fila, 2, {"type": "cell", "criteria": "!=", "value": '"OK"', "format": F.mal})
    wg.merge_range(fila + 2, 0, fila + 3, 1, "El «Resultado del periodo» está enlazado a la pestaña Estado_Resultados. "
                   "Corriente/no corriente según el rubro y el nombre de la cuenta (columna N de la pestaña 1, editable).", F.nota)

    # ------------------------------------------------------------ 4. Flujo de efectivo
    wf = hojas["Flujo_Efectivo"]
    fila = _encabezado(wf, F, e, "Estado de flujos de efectivo – método indirecto (NIC 7 / Sección 7 NIIF para las Pymes)")
    wf.set_column(0, 0, 66); wf.set_column(1, 1, 24)
    wf.write(fila, 0, "Concepto", F.enc); wf.write(fila, 1, "Periodo", F.enc)
    fila += 1
    bloque, subt = [], {}
    fila_var = fila_ini = None
    for item in e.efe:
        tipo, etiqueta, v = item[:3]
        x = fila + 1
        if tipo == "titulo":
            wf.write(fila, 0, etiqueta, F.seccion); bloque = []
        elif tipo == "subtitulo":
            wf.write(fila, 0, etiqueta, F.txt_b)
        elif tipo == "linea" and etiqueta == "Resultado del periodo":
            wf.write(fila, 0, etiqueta, F.sangria)
            wf.write_formula(fila, 1, f"=Estado_Resultados!C{fila_resultado}", F.num, v); bloque.append(x)
        elif tipo == "linea":
            clave = item[3]
            wf.write(fila, 0, etiqueta, F.sangria2)
            wf.write_formula(fila, 1, f'=SUMIFS({rng("efecto_flujo")},{rng("linea_flujo")},"{clave}",{rng("es_detalle")},"Sí")'
                                      f'+SUMIFS({rng("efecto_contra")},{rng("linea_flujo_contra")},"{clave}",{rng("es_detalle")},"Sí")', F.num, v)
            bloque.append(x)
        elif tipo == "subtotal":
            wf.write(fila, 0, etiqueta, F.txt_sub)
            wf.write_formula(fila, 1, "=" + ("+".join(f"B{r}" for r in bloque) or "0"), F.num_b, v)
            subt[etiqueta] = x
        elif tipo == "total" and fila_var is None:
            wf.write(fila, 0, etiqueta, F.txt_tot)
            wf.write_formula(fila, 1, "=" + "+".join(f"B{r}" for r in subt.values()), F.num_tot, v)
            fila_var = x
        elif tipo == "linea_fija":
            wf.write(fila, 0, etiqueta, F.txt)
            wf.write_formula(fila, 1, f'=SUMIFS({rng("saldo_inicial")},{rng("rubro")},"A_EFECTIVO",{rng("es_detalle")},"Sí")', F.num, v)
            fila_ini = x
        elif tipo == "total":
            wf.write(fila, 0, etiqueta, F.txt_tot)
            wf.write_formula(fila, 1, f"=B{fila_ini}+B{fila_var}", F.num_tot, v)
            fila_fin = x
        fila += 1
    fila += 1
    wf.write(fila, 0, "Efectivo y equivalentes según el balance general", F.txt)
    wf.write_formula(fila, 1, f'=SUMIFS({rng("saldo_final")},{rng("rubro")},"A_EFECTIVO",{rng("es_detalle")},"Sí")', F.num, e.cifras["efectivo_final"])
    fila_bal = fila + 1; fila += 1
    wf.write(fila, 0, "Control: efectivo final del flujo − efectivo del balance", F.txt_b)
    wf.write_formula(fila, 1, f"=ROUND(B{fila_fin}-B{fila_bal},2)", F.num, 0)
    wf.write_formula(fila, 2, f'=IF(B{fila + 1}=0,"OK","DESCUADRADO")', F.txt_b, "OK")
    wf.conditional_format(fila, 2, fila, 2, {"type": "cell", "criteria": "==", "value": '"OK"', "format": F.ok})
    wf.conditional_format(fila, 2, fila, 2, {"type": "cell", "criteria": "!=", "value": '"OK"', "format": F.mal})
    wf.merge_range(fila + 2, 0, fila + 4, 1,
                   "Parte del resultado del movimiento del periodo (pestaña Estado_Resultados, columna C). Las partidas que no "
                   "mueven efectivo (depreciación, método de participación, valoración a valor razonable, deterioro) se reversan "
                   "en operación y se descuentan de la variación del activo relacionado en inversión, para no contarlas dos veces. "
                   "Las cuentas por cobrar y por pagar se presentan como cambios en el capital de trabajo.", F.nota)

    # ------------------------------------------------------------ 5. Verificaciones
    wv = hojas["Verificaciones"]
    fila = _encabezado(wv, F, e, "Verificaciones de cuadre contable")
    wv.set_column(0, 0, 64); wv.set_column(1, 3, 22); wv.set_column(4, 4, 16)
    for j, h in enumerate(["Control", "Valor A", "Valor B", "Diferencia", "Estado"]):
        wv.write(fila, j, h, F.enc)
    fila += 1
    S = lambda col, k: f'SUMIFS({rng(col)},{rng("clase")},{k},{rng("es_detalle")},"Sí")'
    ecu = lambda col: f"={S(col, 2)}+{S(col, 3)}+{S(col, 4)}-{S(col, 5)}-{S(col, 6)}-{S(col, 7)}"
    gtot = lambda clave: f"Balance_General!B{clave}"
    checks = [
        ("Partida doble: Σ débitos = Σ créditos (cuentas de detalle, clases 1 a 7)",
         f'=SUMIFS({rng("debitos")},{rng("es_detalle")},"Sí",{rng("clase")},"<=7")',
         f'=SUMIFS({rng("creditos")},{rng("es_detalle")},"Sí",{rng("clase")},"<=7")', e.controles[0][1:]),
        ("Filas donde saldo inicial ± movimientos ≠ saldo final", f'=COUNTIF({rng("control_fila")},"<>0")', "=0", e.controles[1][1:]),
        ("Ecuación contable al inicio: Activo = Pasivo + Patrimonio + (Ingresos − Gastos − Costos)", f"={S('saldo_inicial', 1)}", ecu("saldo_inicial"), e.controles[2][1:]),
        ("Ecuación contable al cierre: Activo = Pasivo + Patrimonio + (Ingresos − Gastos − Costos)", f"={S('saldo_final', 1)}", ecu("saldo_final"), e.controles[3][1:]),
        ("Balance general: Total activo = Total pasivo + patrimonio", f"={gtot(totales['TA'])}", f"={gtot(totales['TPP'])}", e.controles[4][1:]),
        ("Balance general: Total activo = Σ cuentas de detalle clase 1", f"={gtot(totales['TA'])}", f"={S('saldo_final', 1)}", (e.cifras["total_activo"],) * 2),
        ("Flujo de efectivo: variación neta = efectivo final − efectivo inicial (balance)", f"=Flujo_Efectivo!B{fila_var}",
         f'=SUMIFS({rng("saldo_final")},{rng("rubro")},"A_EFECTIVO",{rng("es_detalle")},"Sí")-SUMIFS({rng("saldo_inicial")},{rng("rubro")},"A_EFECTIVO",{rng("es_detalle")},"Sí")',
         e.controles[5][1:]),
        ("Resultado acumulado del ER = Ingresos − Gastos − Costos (saldos finales)", f"=Estado_Resultados!B{fila_resultado}",
         f"={S('saldo_final', 4)}-{S('saldo_final', 5)}-{S('saldo_final', 6)}-{S('saldo_final', 7)}"
         + (f'-SUMIFS({rng("valor_er_acumulado")},{rng("estado")},"ORI")' if fila_resultado_integral else ""),
         (e.cifras["resultado_acum"],) * 2),
        ("Cuentas con rubro inválido", f'=COUNTIF({rng("estado")},"RUBRO NO VÁLIDO")', "=0", (0, 0)),
    ]
    for etiqueta, fa, fb, (va, vb) in checks:
        x = fila + 1
        wv.write(fila, 0, etiqueta, F.txt)
        wv.write_formula(fila, 1, fa, F.num, va)
        wv.write_formula(fila, 2, fb, F.num, vb)
        wv.write_formula(fila, 3, f"=ROUND(B{x}-C{x},2)", F.num, round(va - vb, 2))
        wv.write_formula(fila, 4, f'=IF(ROUND(B{x}-C{x},2)=0,"OK","DESCUADRADO")', F.txt_b, "OK" if round(va - vb, 2) == 0 else "DESCUADRADO")
        fila += 1
    revisar = int(d["revisar"].sum())
    wv.write(fila, 0, "Cuentas clasificadas por ML con confianza baja (revisar en Alertas)", F.txt)
    wv.write_formula(fila, 1, f'=COUNTIF({rng("fuente_clasificacion")},"*REVISAR*")', F.num, revisar)
    wv.write_formula(fila, 4, f'=IF(B{fila + 1}=0,"OK","REVISAR")', F.txt_b, "OK" if revisar == 0 else "REVISAR")
    fila += 1
    wv.conditional_format(6, 4, fila, 4, {"type": "cell", "criteria": "==", "value": '"OK"', "format": F.ok})
    wv.conditional_format(6, 4, fila, 4, {"type": "cell", "criteria": "!=", "value": '"OK"', "format": F.mal})
    fila += 1
    info = [("Catálogo detectado", e.meta.get("catalogo_desc", "")), ("Grupo normativo", "Grupo 1 – NIIF plenas" if e.grupo == 1 else "Grupo 2 – NIIF para las Pymes"),
            ("Actividad principal (NIIF 18)", e.meta.get("actividad_principal", "ninguna")), ("Convención de signos del archivo", e.meta.get("convencion_signos", "")),
            ("Archivo de origen", e.meta.get("archivo", "")), ("Cuentas leídas", str(n)), ("Cuentas de detalle", str(int(d["es_detalle"].sum())))]
    for k, v in info:
        wv.write(fila, 0, k, F.txt_b); wv.write(fila, 1, v, F.txt); fila += 1

    # ------------------------------------------------------------ 6. Alertas
    wa = hojas["Alertas"]
    fila = _encabezado(wa, F, e, "Alertas de lectura y clasificación")
    wa.set_column(0, 0, 12); wa.set_column(1, 1, 50); wa.set_column(2, 2, 20); wa.set_column(3, 3, 70)
    for j, h in enumerate(["Código", "Cuenta", "Tipo", "Detalle"]):
        wa.write(fila, j, h, F.enc)
    fila += 1
    alertas = []
    for _, f in d[d["es_detalle"]].iterrows():
        if f["revisar"]:
            alertas.append((f["codigo"], f["nombre"], "Revisar clasificación", f"Rubro sugerido {f['rubro']} ({f['confianza']:.0%}). Alternativas: {f['alternativas_ml']}"))
        if f.get("no_niif") == "Sí":
            alertas.append((f["codigo"], f["nombre"], "Partida no NIIF", "Rubro de PCGA anteriores (valorizaciones, ajustes por inflación o revalorización). Debe eliminarse o reclasificarse bajo NIIF."))
        if f["rubro"] == "G_CIERRE":
            alertas.append((f["codigo"], f["nombre"], "Cuenta de cierre", "La cuenta de ganancias y pérdidas tiene saldo: verifique que el cierre no duplique el resultado."))
        if f["rubro"] == "P_OBLIG_FIN":
            alertas.append((f["codigo"], f["nombre"], "Supuesto", "Obligación financiera presentada como corriente. Cambie la columna N a NC si vence en más de 12 meses."))
    for a in e.meta.get("advertencias", []):
        alertas.append(("", "", "Lectura del archivo", a))
    if not alertas:
        alertas.append(("", "", "Sin alertas", "Todas las cuentas se clasificaron con reglas o con confianza suficiente."))
    for a in alertas:
        for j, v in enumerate(a):
            wa.write_string(fila, j, str(v), F.txt)
        fila += 1

    hojas["Balance_General"].activate()
    wb.close()
    return ruta_salida
