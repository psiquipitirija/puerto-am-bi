"""
Genera Puerto_AM_Base_Datos_DEMO.xlsx: base 100% ficticia para práctica académica.

Uso:
    python scripts/generar_excel.py

El archivo se sobrescribe cada vez. Todos los importes están en MXN.
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "Puerto_AM_Base_Datos_DEMO.xlsx"

PETROLEUM = "0B3948"
PETROLEUM_DARK = "062733"
GRAPHITE = "24313B"
GOLD = "C9A227"
GOLD_SOFT = "F7EFD8"
WARM_WHITE = "FBF9F5"
WHITE = "FFFFFF"
INK = "26333D"
MUTED = "667782"
LINE = "DCE4E8"
GREEN = "1E7A5A"

MONEY_FMT = '"$"#,##0.00'
INT_FMT = "#,##0"
PCT_FMT = "0.0%"
DATE_FMT = "dd/mm/yyyy"

THIN = Side(style="thin", color=LINE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Catálogo ficticio. (SKU, producto, categoría, precio, costo)
PRODUCTS = [
    ("PAM-001", "Cuaderno ejecutivo A5", "Papelería", 189.0, 94.0),
    ("PAM-002", "Cuaderno universitario 100 hojas", "Papelería", 96.0, 48.0),
    ("PAM-003", "Bolígrafo metálico dorado", "Papelería", 74.0, 29.0),
    ("PAM-004", "Set de 6 plumones finos", "Papelería", 158.0, 71.0),
    ("PAM-005", "Block de notas adhesivas", "Papelería", 68.0, 31.0),
    ("PAM-006", "Carpeta con ganchos Oficina", "Oficina", 145.0, 66.0),
    ("PAM-007", "Organizador de escritorio 4 niveles", "Oficina", 489.0, 268.0),
    ("PAM-008", "Archivo portátil con llave", "Oficina", 329.0, 174.0),
    ("PAM-009", "Silla de oficina ergonómica", "Oficina", 2890.0, 1780.0),
    ("PAM-010", "Lámpara LED de escritorio", "Oficina", 699.0, 395.0),
    ("PAM-011", "Mouse inalámbrico silencioso", "Tecnología", 549.0, 312.0),
    ("PAM-012", "Teclado mecánico compacto", "Tecnología", 1190.0, 735.0),
    ("PAM-013", "Audífonos con cancelación", "Tecnología", 1590.0, 980.0),
    ("PAM-014", "Base de carga 3 en 1", "Tecnología", 849.0, 498.0),
    ("PAM-015", "Novela contemporánea edición tapa dura", "Libros", 349.0, 176.0),
    ("PAM-016", "Guía de productividad 2026", "Libros", 279.0, 132.0),
    ("PAM-017", "Mochila ejecutiva impermeable", "Mochilas", 1290.0, 745.0),
    ("PAM-018", "Mochila escolar reforzada", "Mochilas", 879.0, 492.0),
]

CHANNELS = ["Tienda física", "E-commerce", "Distribuidores", "WhatsApp"]
CHANNEL_WEIGHTS = [34, 30, 21, 15]
PAYMENTS = ["Tarjeta", "Efectivo", "Transferencia", "Mercado Pago"]
PAYMENT_WEIGHTS = [42, 20, 26, 12]
MONTHS = [(2026, 1, "Enero", 118000, 112, 41500),
          (2026, 2, "Febrero", 132000, 124, 46500),
          (2026, 3, "Marzo", 148000, 138, 53000)]

# Productos con mayor peso en la simulación (índice en PRODUCTS)
PRODUCT_WEIGHTS = [12, 10, 9, 8, 7, 8, 4, 6, 2, 4, 7, 5, 4, 5, 6, 7, 3, 5]


def style_title(ws, cell_range: str, text: str, size: int = 16, fill: str = PETROLEUM) -> None:
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = text
    cell.font = Font(name="Aptos Display", size=size, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    ws.row_dimensions[cell.row].height = 34


def style_header(ws, row: int, columns: int) -> None:
    for col in range(1, columns + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=PETROLEUM)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def style_body(ws, start_row: int, end_row: int, columns: int,
               money_cols: set[int] | None = None, int_cols: set[int] | None = None,
               date_cols: set[int] | None = None, pct_cols: set[int] | None = None) -> None:
    money_cols = money_cols or set()
    int_cols = int_cols or set()
    date_cols = date_cols or set()
    pct_cols = pct_cols or set()
    for row in range(start_row, end_row + 1):
        for col in range(1, columns + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = Font(name="Aptos", size=10, color=INK)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center",
                                       horizontal="right" if col in money_cols | int_cols | pct_cols else "left")
            if row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=WARM_WHITE)
            if col in money_cols:
                cell.number_format = MONEY_FMT
            elif col in int_cols:
                cell.number_format = INT_FMT
            elif col in pct_cols:
                cell.number_format = PCT_FMT
            elif col in date_cols:
                cell.number_format = DATE_FMT


def add_table(ws, ref: str, name: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(table)


def set_widths(ws, widths: list[float]) -> None:
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


def build_ventas(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "Ventas"
    ws.sheet_properties.tabColor = PETROLEUM

    headers = ["Fecha", "Pedido", "Producto", "Categoría", "Cantidad",
               "Precio unitario", "Costo unitario", "Ingreso", "Costo total",
               "Utilidad", "Canal", "Método de pago"]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    rng = random.Random(20260322)
    folio = 1
    for year, month, _label, _target, _target_orders, _target_profit in MONTHS:
        last_day = (date(year + (month == 12), 1 if month == 12 else month + 1, 1) - timedelta(days=1)).day
        for _ in range(120):
            day = rng.randint(1, last_day)
            producto = rng.choices(PRODUCTS, weights=PRODUCT_WEIGHTS, k=1)[0]
            sku, name, category, price, cost = producto
            qty = rng.choices([1, 2, 3, 4, 5, 6, 8, 10], weights=[18, 24, 21, 15, 10, 7, 3, 2], k=1)[0]
            # Marzo cierra con un repunte de volumen para que la tendencia sea positiva.
            if month == 3:
                qty = min(12, max(1, round(qty * 1.22)))
            channel = rng.choices(CHANNELS, weights=CHANNEL_WEIGHTS, k=1)[0]
            payment = rng.choices(PAYMENTS, weights=PAYMENT_WEIGHTS, k=1)[0]
            row = ws.max_row + 1
            ws.append([
                date(year, month, day),
                f"PED-{year}-{folio:04d}",
                name,
                category,
                qty,
                price,
                cost,
                f"=E{row}*F{row}",
                f"=E{row}*G{row}",
                f"=H{row}-I{row}",
                channel,
                payment,
            ])
            folio += 1

    last = ws.max_row
    style_body(ws, 2, last, len(headers), money_cols={6, 7, 8, 9, 10}, int_cols={5}, date_cols={1})
    # Las columnas calculadas también reciben formato moneda aunque sean fórmulas.
    for row in range(2, last + 1):
        for col in (8, 9, 10):
            ws.cell(row=row, column=col).number_format = MONEY_FMT

    add_table(ws, f"A1:L{last}", "VentasDemo")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:L{last}"
    set_widths(ws, [13, 17, 34, 15, 11, 16, 16, 15, 15, 15, 18, 18])
    ws.sheet_view.showGridLines = False
    ws.oddHeader.center.text = "PUERTO A&M · SIMULACIÓN ACADÉMICA · IMPORTES EN MXN"
    ws.oddHeader.center.size = 9
    ws.oddFooter.center.text = "Top Learning · datos ficticios · Página &P de &N"

    dv_qty = DataValidation(type="whole", operator="greaterThan", formula1="0", allow_blank=False)
    dv_qty.error = "La cantidad debe ser un entero mayor a cero."
    dv_qty.errorTitle = "Cantidad inválida"
    dv_qty.showErrorMessage = True
    ws.add_data_validation(dv_qty)
    dv_qty.add(f"E2:E{last}")

    dv_channel = DataValidation(type="list", formula1='"' + ",".join(CHANNELS) + '"', allow_blank=False)
    ws.add_data_validation(dv_channel)
    dv_channel.add(f"K2:K{last}")

    dv_payment = DataValidation(type="list", formula1='"' + ",".join(PAYMENTS) + '"', allow_blank=False)
    ws.add_data_validation(dv_payment)
    dv_payment.add(f"L2:L{last}")

    ws.conditional_formatting.add(
        f"H2:H{last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color=GOLD, showValue=True),
    )


def build_productos(wb: Workbook) -> None:
    ws = wb.create_sheet("Productos")
    ws.sheet_properties.tabColor = GOLD
    headers = ["SKU", "Producto", "Categoría", "Precio", "Costo unitario", "Margen", "Activo"]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    for sku, name, category, price, cost in PRODUCTS:
        row = ws.max_row + 1
        ws.append([sku, name, category, price, cost, f"=(D{row}-E{row})/D{row}", "Sí"])

    last = ws.max_row
    style_body(ws, 2, last, len(headers), money_cols={4, 5}, pct_cols={6})
    add_table(ws, f"A1:G{last}", "CatalogoDemo")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:G{last}"
    set_widths(ws, [13, 38, 16, 15, 17, 12, 10])
    ws.sheet_view.showGridLines = False
    ws.oddHeader.center.text = "PUERTO A&M · CATÁLOGO FICTICIO · MXN"
    ws.oddFooter.center.text = "Top Learning · datos ficticios"

    dv_active = DataValidation(type="list", formula1='"Sí,No"', allow_blank=False)
    ws.add_data_validation(dv_active)
    dv_active.add(f"G2:G{last}")


def build_metas(wb: Workbook) -> None:
    ws = wb.create_sheet("Metas")
    ws.sheet_properties.tabColor = GREEN
    headers = ["Mes", "Año", "Meta de ingresos", "Meta de pedidos", "Meta de utilidad"]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    for year, month, label, target, target_orders, target_profit in MONTHS:
        ws.append([label, year, target, target_orders, target_profit])

    last = ws.max_row
    style_body(ws, 2, last, len(headers), money_cols={3, 5}, int_cols={4})
    add_table(ws, f"A1:E{last}", "MetasDemo")
    ws.freeze_panes = "A2"
    set_widths(ws, [18, 11, 21, 19, 20])
    ws.sheet_view.showGridLines = False
    ws.oddHeader.center.text = "PUERTO A&M · METAS FICTICIAS · MXN"
    ws.oddFooter.center.text = "Top Learning · datos ficticios"


def build_instrucciones(wb: Workbook, rows: int) -> None:
    ws = wb.create_sheet("Instrucciones")
    ws.sheet_properties.tabColor = GRAPHITE
    ws.sheet_view.showGridLines = False
    set_widths(ws, [4, 108, 4])

    style_title(ws, "A1:C1", "PUERTO A&M  |  BASE DE DATOS DEMO", size=18)
    ws.merge_cells("A2:C2")
    ws["A2"] = "SIMULACIÓN ACADÉMICA · TODOS LOS DATOS SON FICTICIOS · TOP LEARNING"
    ws["A2"].font = Font(name="Aptos", size=11, bold=True, color=PETROLEUM_DARK)
    ws["A2"].fill = PatternFill("solid", fgColor=GOLD_SOFT)
    ws["A2"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 26

    blocks = [
        ("AVISO IMPORTANTE", [
            "Este archivo es una simulación académica creada para la práctica de Business Intelligence.",
            "Ningún cliente, pedido, producto, precio o meta corresponde a una operación real.",
            "Los importes están expresados en pesos mexicanos (MXN).",
        ]),
        ("1. DÓNDE EDITAR", [
            f"Ventas: {rows} pedidos ficticios de enero a marzo de 2026. Edita Cantidad, Precio unitario y Costo unitario.",
            "Productos: catálogo con SKU, categoría, precio y costo. Si cambias un precio aquí, actualízalo también en Ventas si quieres coherencia histórica.",
            "Metas: objetivos mensuales de ingresos, pedidos y utilidad para enero, febrero y marzo de 2026.",
            "Las columnas Ingreso, Costo total y Utilidad de la hoja Ventas son fórmulas de Excel: se recalculan al abrir el archivo.",
        ]),
        ("2. REGLAS PARA QUE EL DASHBOARD SIGA FUNCIONANDO", [
            "No cambies los nombres de las hojas ni los encabezados de la fila 1.",
            "Usa fechas reales de Excel en la columna Fecha (formato dd/mm/aaaa).",
            "Escribe números sin símbolos de moneda en Cantidad, Precio unitario y Costo unitario.",
            "No dejes filas intermedias vacías dentro de la tabla de ventas.",
            "Puedes agregar pedidos nuevos: el dashboard los detecta automáticamente.",
        ]),
        ("3. CÓMO ACTUALIZAR EL DASHBOARD", [
            "Guarda el archivo con el mismo nombre y en la misma carpeta data/.",
            "En el dashboard pulsa el botón «Actualizar datos» (barra superior).",
            "El backend vuelve a leer el Excel con openpyxl: no hace falta tocar código ni reiniciar el servidor.",
            "El indicador «Fuente conectada» muestra la fecha y hora de la última lectura del archivo.",
        ]),
        ("4. QUÉ CALCULA EL DASHBOARD", [
            "Ingreso = Cantidad × Precio unitario.  Costo = Cantidad × Costo unitario.  Utilidad = Ingreso − Costo.",
            "Ticket promedio = Ingreso total ÷ número de pedidos.  Margen bruto = Utilidad ÷ Ingreso.",
            "Crecimiento mensual = variación del ingreso del último mes respecto al mes anterior presente en el filtro.",
            "Cumplimiento de metas = ingreso real del mes ÷ meta de ingresos de ese mes (meses sin meta se omiten).",
            "Los filtros de fecha, producto y categoría se aplican antes de todos los cálculos.",
        ]),
        ("5. SOPORTE", [
            "Si el dashboard muestra un error, revisa que el archivo no esté abierto con cambios sin guardar.",
            "Para regenerar la base original: python scripts/generar_excel.py",
            "Para iniciar la aplicación: uvicorn app.main:app --reload --port 8124",
        ]),
    ]

    row = 4
    for heading, lines in blocks:
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=2)
        cell = ws.cell(row=row, column=2, value=heading)
        cell.font = Font(name="Aptos", size=11, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=PETROLEUM)
        cell.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[row].height = 24
        row += 1
        for line in lines:
            c = ws.cell(row=row, column=2, value="•  " + line)
            c.font = Font(name="Aptos", size=10, color=INK)
            c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
            c.fill = PatternFill("solid", fgColor=WARM_WHITE if row % 2 == 0 else WHITE)
            ws.row_dimensions[row].height = 30 if len(line) > 95 else 22
            row += 1
        row += 1

    ws.merge_cells(start_row=row + 1, start_column=2, end_row=row + 1, end_column=2)
    note = ws.cell(row=row + 1, column=2, value="PUERTO A&M · BUSINESS INTELLIGENCE · USO EXCLUSIVAMENTE ACADÉMICO")
    note.font = Font(name="Aptos", size=9, italic=True, color=MUTED)
    note.alignment = Alignment(horizontal="center")


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    build_ventas(wb)
    build_productos(wb)
    build_metas(wb)
    build_instrucciones(wb, 360)
    wb.properties.title = "Puerto A&M · Base de datos demo"
    wb.properties.subject = "Simulación académica · Business Intelligence"
    wb.properties.creator = "Puerto A&M · Top Learning"
    wb.properties.description = "Datos 100% ficticios para práctica académica. Importes en MXN."
    wb.save(OUT)
    print(f"Generado: {OUT}")
    print(f"Hojas: {', '.join(wb.sheetnames)}")
    print("Registros de ventas: 360")


if __name__ == "__main__":
    main()
