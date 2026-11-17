"""Motor de estados financieros NIIF a partir de un balance de prueba.

Uso:
    python main.py                       # modo interactivo: pide el archivo y el grupo
    python main.py --archivo balance.pdf --grupo 1
    python main.py --archivo balance.xlsx --grupo 2 --no-interactivo

Salida: un Excel con el Estado de Situación Financiera, el Estado de Resultados y el
Estado de Flujos de Efectivo, conectados por fórmulas, más verificaciones de cuadre.
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ))

from src.accounting_engine.motor import construir  # noqa: E402
from src.classification.clasificador import ClasificadorHibrido  # noqa: E402
from src.classification.taxonomias import CATALOGOS  # noqa: E402
from src.exporter.excel import exportar  # noqa: E402
from src.ingestion import leer_balance  # noqa: E402


def cargar_config() -> dict:
    ruta = RAIZ / "config.yaml"
    if ruta.exists():
        return leer_config_simple(ruta.read_text(encoding="utf-8"))
    return {}


def leer_config_simple(texto: str) -> dict:
    """Lee config.yaml (pares «clave: valor») sin depender de PyYAML."""
    cfg = {}
    for linea in texto.splitlines():
        linea = linea.split("#", 1)[0].rstrip()
        if not linea.strip() or ":" not in linea or linea[0].isspace():
            continue
        clave, valor = linea.split(":", 1)
        valor = valor.strip().strip('"').strip("'")
        if valor == "" or valor.lower() in ("null", "none", "~"):
            cfg[clave.strip()] = None
            continue
        for tipo in (int, float):
            try:
                valor = tipo(valor)
                break
            except ValueError:
                pass
        cfg[clave.strip()] = valor
    return cfg


def pedir_archivo() -> Path:
    print("\nSuba el balance de prueba (PDF, Excel o CSV).")
    print("Arrastre el archivo a esta ventana y presione Enter, o escriba la ruta.")
    print("Si presiona Enter sin escribir nada, se abre una ventana para buscarlo.")
    while True:
        resp = input("Archivo: ").strip().strip('"').strip("'").strip()
        if not resp:
            try:
                import tkinter as tk
                from tkinter import filedialog
                raiz = tk.Tk(); raiz.withdraw(); raiz.attributes("-topmost", True)
                resp = filedialog.askopenfilename(
                    title="Seleccione el balance de prueba",
                    filetypes=[("Balances", "*.pdf *.xlsx *.xls *.csv"), ("Todos", "*.*")])
                raiz.destroy()
            except Exception:
                print("  No se pudo abrir la ventana. Escriba la ruta del archivo.")
                continue
            if not resp:
                continue
        ruta = Path(resp)
        if ruta.exists():
            return ruta
        print(f"  No se encontró «{ruta}». Intente de nuevo.")


def pedir_grupo(sugerido: int) -> int:
    print("\n¿Qué marco normativo aplica la empresa?")
    print("  1 = Grupo 1 – NIIF plenas (Dec. 2784/2012). Estado de resultados según NIIF 18")
    print("  2 = Grupo 2 – NIIF para las Pymes (Dec. 3022/2013)")
    while True:
        resp = input(f"Grupo [Enter = {sugerido}]: ").strip()
        if resp == "":
            return sugerido
        if resp in ("1", "2"):
            return int(resp)
        print("  Escriba 1 o 2.")


def main(argv=None) -> int:
    cfg = cargar_config()
    ap = argparse.ArgumentParser(description="Genera los estados financieros NIIF desde un balance de prueba.")
    ap.add_argument("--archivo", help="Ruta del balance de prueba (PDF, Excel o CSV)")
    ap.add_argument("--grupo", type=int, choices=[1, 2], default=cfg.get("grupo"))
    ap.add_argument("--catalogo", default=cfg.get("catalogo", "auto"), choices=["auto", *CATALOGOS.keys()])
    ap.add_argument("--actividad-principal", default=cfg.get("actividad_principal", "ninguna"),
                    choices=["ninguna", "inversion", "financiacion"])
    ap.add_argument("--umbral", type=float, default=float(cfg.get("umbral_confianza", 0.85)))
    ap.add_argument("--salida", help="Ruta del Excel de salida")
    ap.add_argument("--no-interactivo", action="store_true", help="No hace preguntas; marca para revisión lo dudoso")
    args = ap.parse_args(argv)
    interactivo = not args.no_interactivo

    print("=" * 78)
    print(" MOTOR DE ESTADOS FINANCIEROS NIIF · balance de prueba → 3 estados conectados")
    print("=" * 78)
    ruta = Path(args.archivo) if args.archivo else (pedir_archivo() if interactivo else None)
    if ruta is None:
        print("Falta --archivo en modo no interactivo."); return 2

    while True:
        print(f"\n[1/4] Leyendo {ruta.name} ...")
        try:
            df, meta = leer_balance(ruta)
            break
        except (ValueError, FileNotFoundError) as err:
            print(f"\n      ✖ No se pudo leer el archivo.\n      {err}")
            if not interactivo:
                return 3
            print("\n      Pruebe con otro archivo (o cierre la ventana para salir).")
            ruta = pedir_archivo()
    print(f"      {meta.get('empresa') or 'Empresa sin nombre'}  {('NIT ' + meta['nit']) if meta.get('nit') else ''}")
    if meta.get("periodo"):
        print(f"      Periodo: {meta['periodo']}")
    print(f"      {len(df)} cuentas leídas, {int(df['es_detalle'].sum())} de detalle. Saldos: {meta['convencion_signos']}.")
    for a in meta.get("advertencias", []):
        print(f"      ⚠ {a}")

    print("\n[2/4] Clasificando cuentas (reglas + Machine Learning) ...")
    clasif = ClasificadorHibrido(umbral=args.umbral, interactivo=interactivo)
    dfc, catalogo, puntajes = clasif.clasificar(df, args.catalogo)
    meta["catalogo"] = catalogo
    meta["catalogo_desc"] = CATALOGOS.get(catalogo, (None, "Desconocido: clasificación por nombre y ML"))[1]
    print(f"      Catálogo detectado: {meta['catalogo_desc']}")
    hojas = dfc[dfc["es_detalle"]]
    print("      Fuentes: " + ", ".join(f"{k}: {v}" for k, v in hojas["fuente"].value_counts().items()))

    grupo = args.grupo
    if grupo is None:
        sugerido = 1 if catalogo == "cuif_financiero" else 2
        grupo = pedir_grupo(sugerido) if interactivo else sugerido
    meta["actividad_principal"] = args.actividad_principal

    print(f"\n[3/4] Construyendo estados financieros (Grupo {grupo}) ...")
    estados = construir(dfc, clasif.rubros, grupo, meta, args.actividad_principal)
    c = estados.cifras
    print(f"      Total activo:            {c['total_activo']:>24,.2f}")
    print(f"      Total pasivo:            {c['total_pasivo']:>24,.2f}")
    print(f"      Total patrimonio:        {c['total_patrimonio']:>24,.2f}")
    print(f"      Resultado acumulado:     {c['resultado_acum']:>24,.2f}")
    print(f"      Variación del efectivo:  {c['variacion_efectivo']:>24,.2f}")
    print("      Verificaciones:")
    todo_ok = True
    for nombre, a, b in estados.controles:
        ok = round(a - b, 2) == 0
        todo_ok &= ok
        print(f"        {'OK ' if ok else 'XX '} {nombre}")

    print("\n[4/4] Generando Excel ...")
    if args.salida:
        salida = Path(args.salida)
    else:
        nombre = re.sub(r"[^\w]+", "_", (meta.get("empresa") or ruta.stem)).strip("_")[:40]
        salida = RAIZ / cfg.get("carpeta_salida", "data/output") / f"EEFF_{nombre}_G{grupo}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    exportar(estados, salida)
    print(f"      Listo: {salida}")
    revisar = int(hojas["revisar"].sum())
    if revisar:
        print(f"      ⚠ {revisar} cuenta(s) quedaron para revisión: vea la pestaña Alertas.")
    if not todo_ok:
        print("      ⚠ Hay controles descuadrados: vea la pestaña Verificaciones.")
    return 0 if todo_ok else 1


if __name__ == "__main__":
    sys.exit(main())
