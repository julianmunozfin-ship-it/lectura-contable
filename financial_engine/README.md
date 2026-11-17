# Motor de estados financieros NIIF

Lee un balance de prueba (PDF, Excel o CSV), o un libro auxiliar exportado del ERP, y entrega un Excel con los tres estados financieros de la empresa, conectados entre sí por fórmulas:

1. Estado de Situación Financiera (balance general), clasificado en corriente y no corriente.
2. Estado de Resultados: por categorías de la NIIF 18 en Grupo 1, o según la NIIF para las Pymes en Grupo 2.
3. Estado de Flujos de Efectivo por el método indirecto.

El Excel también trae el balance normalizado, una hoja de verificaciones con semáforos de cuadre y una hoja de alertas con las cuentas que conviene revisar.

## Cómo usarlo en Windows

La forma más simple es hacer doble clic en `ejecutar.bat`. Si el computador no tiene Python, el lanzador ofrece instalarlo (escriba S y Enter); si no puede hacerlo solo, abre la página de descarga. La primera vez también crea el entorno virtual e instala las librerías, lo que toma unos minutos. Después el programa:

1. Pide el balance de prueba. Puede arrastrar el archivo a la ventana y presionar Enter. Si presiona Enter sin escribir nada, se abre una ventana para buscarlo.
2. Detecta el catálogo de cuentas (PUC 2650, CUIF de la Superintendencia Financiera o CUIF de la Supersolidaria).
3. Pregunta si la empresa es Grupo 1 (NIIF plenas) o Grupo 2 (NIIF para las Pymes) y sugiere uno según el catálogo.
4. Si el modelo de Machine Learning no está seguro de alguna cuenta (confianza menor a 0,85), pregunta en pantalla a qué rubro va. La respuesta queda guardada y el modelo la aprende para la próxima vez.
5. Guarda el Excel en `data/output`.

Requisito: Python 3.10 o superior. Si lo instala a mano desde python.org, marque la casilla «Add python.exe to PATH».

## Desde VS Code o la terminal

```bash
python -m venv .venv
.venv\Scripts\activate          # en Windows
pip install -r requirements.txt

python main.py                                         # modo interactivo
python main.py --archivo balance.pdf --grupo 1         # sin preguntar el grupo
python main.py --archivo balance.xlsx --grupo 2 --no-interactivo --salida resultado.xlsx
```

Otras opciones: `--catalogo` (auto, puc_2650, cuif_solidario, cuif_financiero), `--actividad-principal` (ninguna, inversion, financiacion) y `--umbral` (confianza mínima del modelo). Los valores por defecto están en `config.yaml`.

## Qué hace cada parte

```
financial_engine/
├── main.py                      programa principal: pregunta, procesa y exporta
├── config.yaml                  grupo, catálogo, actividad principal, umbral
├── ejecutar.bat                 lanzador para Windows
├── data/
│   ├── taxonomies/              PUC 2650, CUIF solidario, grupos CUIF financiero,
│   │                            catálogo de rubros NIIF y reglas por código y por nombre
│   ├── training/                validaciones humanas que reentrenan el modelo
│   ├── models/                  modelo ML entrenado (se crea solo)
│   └── output/                  Excel de salida
├── src/
│   ├── ingestion/               lectores de PDF, Excel y CSV; limpieza de números
│   ├── classification/          reglas deterministas + modelo TF-IDF y regresión logística
│   ├── accounting_engine/       Grupo 1 / Grupo 2, los tres estados y los controles
│   └── exporter/                Excel con fórmulas nativas (xlsxwriter)
└── tests/balance_prueba_simesa.xlsx   balance de Valores Simesa (sep-2020) para probar
```

**Lectura.** Acepta cifras en formato latino (1.224.609.811,68) y anglosajón (1,224,609,811.68), negativos con signo, paréntesis o guion al final, y letras pegadas al saldo («D», «C»). En PDF reúne los nombres partidos en varias líneas. Identifica clase, grupo, cuenta y subcuenta, trabaja solo con las cuentas de detalle para no duplicar saldos y separa las cuentas de orden (clases 8 y 9).

**Libro auxiliar.** Si el archivo trae un movimiento por línea (fecha, documento, tercero, código de cuenta, débito y crédito, como lo exportan SAP o Siigo) y no trae columnas de saldos, el programa suma los movimientos por cuenta y arma el balance de prueba. Como un auxiliar no trae saldos iniciales, los toma en cero: los estados muestran solo lo que se movió en el periodo. Si los débitos no suman lo mismo que los créditos, lo avisa en pantalla y en la hoja de verificaciones.

**Clasificación.** Primero aplica reglas por código y por nombre de cuenta. Lo que no cubren las reglas pasa al modelo de Machine Learning. Si la confianza es baja, pregunta (modo interactivo) o deja la cuenta marcada en la hoja «Alertas» (modo no interactivo).

**Estados.** En Grupo 1 el estado de resultados sigue las cinco categorías de la NIIF 18 (operación, inversión, financiación, impuesto a las ganancias y operaciones discontinuadas), con los subtotales «resultado de operación» y «resultado antes de financiación e impuestos». En Grupo 2 usa la presentación tradicional de la Sección 5 de la NIIF para las Pymes. El resultado del periodo pasa al patrimonio del balance, y el flujo de efectivo parte de ese resultado, reversa las partidas que no mueven caja (depreciaciones, método de participación, valoraciones a valor razonable) y suma las variaciones de capital de trabajo, inversión y financiación.

## Cómo leer el Excel de salida

| Hoja | Contenido |
|---|---|
| Balance_Prueba_Normalizado | Matriz base. Las columnas de rubro, corriente/no corriente y función tienen listas desplegables: si cambia la clasificación de una cuenta, los tres estados se recalculan solos. |
| Estado_Resultados | Por categorías NIIF 18 (Grupo 1) o NIIF para las Pymes (Grupo 2), con el periodo y el acumulado. |
| Balance_General | Activo, pasivo y patrimonio. El resultado del ejercicio viene por fórmula del estado de resultados. |
| Flujo_Efectivo | Método indirecto, conectado al balance y al resultado. |
| Verificaciones | Semáforos «OK» / «DESCUADRADO»: partida doble, ecuación contable al inicio y al cierre, cuadre del balance y del flujo contra la variación de la cuenta 11. |
| Alertas | Cuentas excluidas, de baja confianza o con advertencias. |
| Catalogo_Rubros | Rubros disponibles y cómo se presentan en cada estado. |

## Prueba realizada

Con el balance de Valores Simesa S.A. (septiembre de 2020, catálogo CUIF, 95 cuentas) en Grupo 1 y en Grupo 2:

- Total activo 268.589.641.352,51 = pasivo 14.183.618.402,96 + patrimonio 254.406.022.949,55.
- Resultado acumulado enero–septiembre: −10.866.591.821,68.
- Variación del efectivo en el mes: 111.763.029,80, igual a la que calcula el flujo.
- Todas las verificaciones en «OK». Al recalcular el Excel en LibreOffice, las 1.456 fórmulas dan los mismos valores que calcula Python.

## Limitaciones conocidas

- El catálogo CUIF de la carpeta es el de la Supersolidaria. Para la Superintendencia Financiera el motor usa reglas por grupo y por nombre de cuenta; con el catálogo oficial de la Superfinanciera se podría afinar.
- Un balance de un solo mes da el flujo de efectivo de ese mes. Para el flujo del año se necesita el balance con saldos al 1 de enero.
- Con un libro auxiliar, el balance general solo es completo si el auxiliar arranca desde la apertura de la empresa. Para un periodo intermedio conviene usar el balance de prueba, que sí trae saldos iniciales.
- Los PDF escaneados (imagen) no se pueden leer; hace falta un PDF con texto o el balance en Excel o CSV.
