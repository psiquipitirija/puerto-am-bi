"""Lectura, filtrado y agregación del Excel de demostración.

Toda cifra mostrada por el dashboard se calcula aquí a partir del archivo
Puerto_AM_Base_Datos_DEMO.xlsx. No existen datos incrustados en el frontend.
"""
from __future__ import annotations

import unicodedata
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

DATA_FILE = Path(__file__).resolve().parents[1] / "data" / "Puerto_AM_Base_Datos_DEMO.xlsx"

_MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11,
    "diciembre": 12,
}
_MONTH_LABELS = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

_SALES_ALIASES = {
    "fecha": "fecha", "pedido": "pedido", "producto": "producto",
    "categoria": "categoria", "cantidad": "cantidad",
    "preciounitario": "precio", "precio": "precio",
    "costounitario": "costo", "costo": "costo",
    "ingreso": "ingreso", "costototal": "costo_total", "utilidad": "utilidad",
    "canal": "canal", "metododepago": "pago", "metodopago": "pago",
    "formadepago": "pago", "pago": "pago",
}
_PRODUCT_ALIASES = {
    "sku": "sku", "producto": "producto", "categoria": "categoria",
    "precio": "precio", "costounitario": "costo", "costo": "costo",
    "margen": "margen", "activo": "activo",
}
_GOAL_ALIASES = {
    "mes": "mes", "anio": "anio", "ano": "anio", "year": "anio",
    "metaingresos": "meta_ingresos", "metadeingresos": "meta_ingresos",
    "metapedidos": "meta_pedidos", "metadepedidos": "meta_pedidos",
    "metautilidad": "meta_utilidad", "metadeutilidad": "meta_utilidad",
}

_cache: dict[str, Any] = {"key": None, "payload": None}


class ExcelSourceError(Exception):
    """Error legible cuando el archivo no tiene la estructura esperada."""


def normalize(text: Any) -> str:
    """Minúsculas, sin acentos ni espacios: permite tolerar variaciones de encabezado."""
    decomposed = unicodedata.normalize("NFD", str(text or ""))
    plain = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    return "".join(plain.lower().split())


def _signature() -> tuple[int, int]:
    stat = DATA_FILE.stat()
    return stat.st_mtime_ns, stat.st_size


def _header_index(ws, aliases: dict[str, str], sheet: str) -> dict[str, int]:
    header = [normalize(cell.value) for cell in ws[1]]
    index: dict[str, int] = {}
    for pos, name in enumerate(header):
        key = aliases.get(name)
        if key and key not in index:
            index[key] = pos
    required = {
        "ventas": {"fecha", "pedido", "producto", "cantidad", "precio", "costo"},
        "productos": {"producto", "precio"},
        "metas": {"mes", "meta_ingresos"},
    }.get(sheet, set())
    missing = sorted(required - set(index))
    if missing:
        raise ExcelSourceError(
            f"La hoja '{sheet}' no tiene las columnas requeridas: {', '.join(missing)}. "
            "No cambies los encabezados de la fila 1."
        )
    return index


def _as_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        converted = from_excel(value)
        return converted.date() if isinstance(converted, datetime) else converted
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Fecha no reconocida: {value!r}")


def _as_decimal(value: Any, field: str, row_number: int) -> Decimal:
    if value is None or str(value).strip() == "":
        return Decimal("0")
    if isinstance(value, bool):
        raise ValueError(f"Fila {row_number}: {field} no puede ser booleano.")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip().replace("$", "").replace("MXN", "").replace(",", "").strip()
    if text.startswith("="):
        raise ValueError(
            f"Fila {row_number}: la columna '{field}' contiene una fórmula sin valor calculado."
        )
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Fila {row_number}: '{value}' no es un número válido en {field}.") from exc


def _as_month(value: Any) -> tuple[int, int] | None:
    if isinstance(value, (datetime, date)):
        return value.year, value.month
    text = normalize(value)
    for name, month in _MONTHS.items():
        if name in text:
            return None, month
    digits = [int(part) for part in text.replace("/", " ").split() if part.isdigit()]
    if len(digits) >= 2 and 1 <= digits[1] <= 12:
        return digits[0], digits[1]
    if len(digits) == 1 and 1 <= digits[0] <= 12:
        return None, digits[0]
    return None


def _money(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.01")))


def _read_products(wb) -> dict[str, dict[str, Any]]:
    ws = wb["Productos"]
    idx = _header_index(ws, _PRODUCT_ALIASES, "productos")
    catalog: dict[str, dict[str, Any]] = {}
    for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(value is None or str(value).strip() == "" for value in row):
            continue

        def cell(key: str) -> Any:
            return row[idx[key]] if key in idx and idx[key] < len(row) else None

        name = str(cell("producto") or "").strip()
        if not name:
            continue
        catalog[name] = {
            "sku": str(cell("sku") or "").strip(),
            "categoria": str(cell("categoria") or "Sin categoría").strip() or "Sin categoría",
            "precio": _as_decimal(cell("precio"), "precio", row_number),
            "costo": _as_decimal(cell("costo"), "costo", row_number),
            "activo": str(cell("activo") or "Sí").strip().lower() not in {"no", "false", "0"},
        }
    return catalog


def _read_goals(wb) -> dict[tuple[int, int], dict[str, Decimal]]:
    ws = wb["Metas"]
    idx = _header_index(ws, _GOAL_ALIASES, "metas")
    goals: dict[tuple[int, int], dict[str, Decimal]] = {}
    for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(value is None or str(value).strip() == "" for value in row):
            continue

        def cell(key: str) -> Any:
            return row[idx[key]] if key in idx and idx[key] < len(row) else None

        month_info = _as_month(cell("mes"))
        if month_info is None:
            raise ValueError(f"Fila {row_number}: no se pudo interpretar el mes de la meta.")
        year = cell("anio")
        year = int(year) if year not in (None, "") else month_info[0]
        if year is None:
            raise ValueError(f"Fila {row_number}: la meta no especifica año.")
        goals[(int(year), month_info[1])] = {
            "ingresos": _as_decimal(cell("meta_ingresos"), "meta de ingresos", row_number),
            "pedidos": _as_decimal(cell("meta_pedidos"), "meta de pedidos", row_number),
            "utilidad": _as_decimal(cell("meta_utilidad"), "meta de utilidad", row_number),
        }
    return goals


def _read_sales(wb, catalog: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    ws = wb["Ventas"]
    idx = _header_index(ws, _SALES_ALIASES, "ventas")
    sales: list[dict[str, Any]] = []
    for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(value is None or str(value).strip() == "" for value in row):
            continue

        def cell(key: str) -> Any:
            return row[idx[key]] if key in idx and idx[key] < len(row) else None

        try:
            sale_date = _as_date(cell("fecha"))
            product = str(cell("producto") or "").strip()
            if not product:
                raise ValueError("producto vacío")
            category = str(cell("categoria") or "").strip()
            if not category:
                category = catalog.get(product, {}).get("categoria", "Sin categoría")
            quantity = int(_as_decimal(cell("cantidad"), "cantidad", row_number))
            price = _as_decimal(cell("precio"), "precio unitario", row_number)
            cost = _as_decimal(cell("costo"), "costo unitario", row_number)
            if quantity <= 0:
                raise ValueError("la cantidad debe ser mayor a cero")
        except (ValueError, TypeError) as exc:
            raise ValueError(f"Fila {row_number} de Ventas: {exc}") from exc

        income = Decimal(quantity) * price
        total_cost = Decimal(quantity) * cost
        sales.append({
            "fecha": sale_date,
            "pedido": str(cell("pedido") or f"FILA-{row_number}").strip(),
            "producto": product,
            "categoria": category,
            "cantidad": quantity,
            "precio": price,
            "costo_unitario": cost,
            "ingreso": income,
            "costo": total_cost,
            "utilidad": income - total_cost,
            "canal": str(cell("canal") or "Sin especificar").strip(),
            "pago": str(cell("pago") or "Sin especificar").strip(),
        })
    if not sales:
        raise ExcelSourceError("La hoja 'Ventas' no contiene registros.")
    return sales


def load_source(force: bool = False) -> dict[str, Any]:
    """Lee el Excel una vez y guarda en caché mientras no cambie mtime/tamaño."""
    if not DATA_FILE.exists():
        raise ExcelSourceError(f"No se encontró el archivo: {DATA_FILE}")
    signature = _signature()
    if force:
        _cache["key"] = None
    if _cache["key"] == signature and _cache["payload"] is not None:
        payload = dict(_cache["payload"])
        payload["cache_hit"] = True
        return payload

    try:
        wb = load_workbook(DATA_FILE, data_only=False, read_only=True)
    except Exception as exc:  # noqa: BLE001 - error de archivo abierto/corrupto
        raise ExcelSourceError(f"No se pudo abrir el Excel: {exc}") from exc

    try:
        missing = [name for name in ("Ventas", "Productos", "Metas") if name not in wb.sheetnames]
        if missing:
            raise ExcelSourceError(f"Faltan hojas en el Excel: {', '.join(missing)}")
        catalog = _read_products(wb)
        goals = _read_goals(wb)
        sales = _read_sales(wb, catalog)
    finally:
        wb.close()

    modified = datetime.fromtimestamp(signature[0] / 1_000_000_000)
    payload = {
        "sales": sales,
        "catalog": catalog,
        "goals": goals,
        "cache_hit": False,
        "source": {
            "file": DATA_FILE.name,
            "modified": modified.isoformat(timespec="seconds"),
            "modified_label": modified.strftime("%d/%m/%Y %H:%M:%S"),
            "bytes": signature[1],
            "rows_total": len(sales),
            "products_total": len(catalog),
            "goals_total": len(goals),
            "simulation": True,
        },
    }
    _cache["key"] = signature
    _cache["payload"] = payload
    return dict(payload, cache_hit=False)


def get_source_meta(force: bool = False) -> dict[str, Any]:
    payload = load_source(force=force)
    return {
        **payload["source"],
        "cache_hit": payload["cache_hit"],
        "server_time": datetime.now().isoformat(timespec="seconds"),
    }


def filter_sales(sales: list[dict[str, Any]], filters: dict[str, Any]) -> list[dict[str, Any]]:
    date_from = filters.get("fecha_desde")
    date_to = filters.get("fecha_hasta")
    product = (filters.get("producto") or "").strip()
    category = (filters.get("categoria") or "").strip()
    channel = (filters.get("canal") or "").strip()

    result = []
    for row in sales:
        if date_from and row["fecha"] < date_from:
            continue
        if date_to and row["fecha"] > date_to:
            continue
        if product and normalize(row["producto"]) != normalize(product):
            continue
        if category and normalize(row["categoria"]) != normalize(category):
            continue
        if channel and normalize(row["canal"]) != normalize(channel):
            continue
        result.append(row)
    return result


def _monthly_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    buckets: dict[tuple[int, int], dict[str, Any]] = {}
    for row in rows:
        key = (row["fecha"].year, row["fecha"].month)
        bucket = buckets.setdefault(key, {
            "key": f"{key[0]:04d}-{key[1]:02d}",
            "year": key[0], "month": key[1],
            "label": f"{_MONTH_LABELS[key[1]]} {key[0]}",
            "ingresos": Decimal("0"), "costos": Decimal("0"),
            "utilidad": Decimal("0"), "pedidos": 0, "unidades": 0,
        })
        bucket["ingresos"] += row["ingreso"]
        bucket["costos"] += row["costo"]
        bucket["utilidad"] += row["utilidad"]
        bucket["unidades"] += row["cantidad"]
    orders_by_month: dict[tuple[int, int], set[str]] = {}
    for row in rows:
        orders_by_month.setdefault((row["fecha"].year, row["fecha"].month), set()).add(row["pedido"])
    for key, bucket in buckets.items():
        bucket["pedidos"] = len(orders_by_month.get(key, set()))
    return [buckets[key] for key in sorted(buckets)]


def build_dashboard(filters: dict[str, Any] | None = None) -> dict[str, Any]:
    filters = filters or {}
    payload = load_source()
    sales = payload["sales"]
    catalog = payload["catalog"]
    goals = payload["goals"]
    rows = filter_sales(sales, filters)

    ingresos = sum((row["ingreso"] for row in rows), Decimal("0"))
    costos = sum((row["costo"] for row in rows), Decimal("0"))
    utilidad = ingresos - costos
    unidades = sum(row["cantidad"] for row in rows)
    pedidos = len({row["pedido"] for row in rows})
    margen = (utilidad / ingresos * 100) if ingresos > 0 else None

    monthly = _monthly_rows(rows)
    if len(monthly) >= 2 and monthly[-2]["ingresos"] > 0:
        growth = (monthly[-1]["ingresos"] / monthly[-2]["ingresos"] - 1) * 100
    else:
        growth = None

    by_product: dict[str, dict[str, Any]] = {}
    by_category: dict[str, dict[str, Any]] = {}
    by_channel: dict[str, Decimal] = {}
    by_payment: dict[str, dict[str, Any]] = {}
    for row in rows:
        p = by_product.setdefault(row["producto"], {
            "producto": row["producto"], "ingresos": Decimal("0"), "unidades": 0, "utilidad": Decimal("0")
        })
        p["ingresos"] += row["ingreso"]
        p["unidades"] += row["cantidad"]
        p["utilidad"] += row["utilidad"]

        c = by_category.setdefault(row["categoria"], {
            "categoria": row["categoria"], "ingresos": Decimal("0"), "unidades": 0
        })
        c["ingresos"] += row["ingreso"]
        c["unidades"] += row["cantidad"]
        by_channel[row["canal"]] = by_channel.get(row["canal"], Decimal("0")) + row["ingreso"]

        pay = by_payment.setdefault(row["pago"], {
            "label": row["pago"], "ingresos": Decimal("0"), "unidades": 0, "pedidos": set()
        })
        pay["ingresos"] += row["ingreso"]
        pay["unidades"] += row["cantidad"]
        pay["pedidos"].add(row["pedido"])

    products_sorted = sorted(by_product.values(), key=lambda item: item["ingresos"], reverse=True)
    categories_sorted = sorted(by_category.values(), key=lambda item: item["ingresos"], reverse=True)

    goal_rows = []
    total_real = Decimal("0")
    total_goal = Decimal("0")
    for month in monthly:
        target = goals.get((month["year"], month["month"]))
        if not target:
            continue
        total_real += month["ingresos"]
        total_goal += target["ingresos"]
        goal_rows.append({
            "key": month["key"],
            "label": month["label"],
            "real": _money(month["ingresos"]),
            "meta": _money(target["ingresos"]),
            "cumplimiento": (float(month["ingresos"] / target["ingresos"] * 100)
                             if target["ingresos"] > 0 else None),
            "pedidos": month["pedidos"],
            "meta_pedidos": _money(target["pedidos"]),
        })
    overall_attainment = (float(total_real / total_goal * 100) if total_goal > 0 else None)

    table = [{
        "fecha": row["fecha"].isoformat(),
        "pedido": row["pedido"],
        "producto": row["producto"],
        "categoria": row["categoria"],
        "cantidad": row["cantidad"],
        "precio": _money(row["precio"]),
        "ingreso": _money(row["ingreso"]),
        "costo": _money(row["costo"]),
        "utilidad": _money(row["utilidad"]),
        "canal": row["canal"],
        "pago": row["pago"],
    } for row in sorted(rows, key=lambda item: (item["fecha"], item["pedido"]), reverse=True)]

    options = {
        "productos": sorted({row["producto"] for row in sales}),
        "categorias": sorted({row["categoria"] for row in sales}),
        "canales": sorted({row["canal"] for row in sales}),
        "metodos_pago": sorted({row["pago"] for row in sales}),
        "fecha_min": min(row["fecha"] for row in sales).isoformat(),
        "fecha_max": max(row["fecha"] for row in sales).isoformat(),
    }

    return {
        "meta": {**payload["source"], "cache_hit": payload["cache_hit"],
                 "server_time": datetime.now().isoformat(timespec="seconds")},
        "filters_applied": {
            "fecha_desde": filters.get("fecha_desde") or None,
            "fecha_hasta": filters.get("fecha_hasta") or None,
            "producto": filters.get("producto") or None,
            "categoria": filters.get("categoria") or None,
            "canal": filters.get("canal") or None,
        },
        "options": options,
        "kpis": {
            "ingresos": _money(ingresos),
            "pedidos": pedidos,
            "ticket_promedio": _money(ingresos / pedidos) if pedidos else 0.0,
            "unidades": unidades,
            "costos": _money(costos),
            "utilidad": _money(utilidad),
            "margen": float(margen) if margen is not None else None,
            "crecimiento": float(growth) if growth is not None else None,
            "cumplimiento_metas": overall_attainment,
            "filas_filtradas": len(rows),
        },
        "charts": {
            "tendencia": [{
                "key": m["key"], "label": m["label"],
                "ingresos": _money(m["ingresos"]), "pedidos": m["pedidos"],
            } for m in monthly],
            "ingresos_costos": [{
                "key": m["key"], "label": m["label"],
                "ingresos": _money(m["ingresos"]), "costos": _money(m["costos"]),
                "utilidad": _money(m["utilidad"]),
            } for m in monthly],
            "categorias": [{
                "label": c["categoria"], "value": _money(c["ingresos"]), "unidades": c["unidades"]
            } for c in categories_sorted],
            "productos": [{
                "label": p["producto"], "value": _money(p["ingresos"]),
                "unidades": p["unidades"], "utilidad": _money(p["utilidad"])
            } for p in products_sorted[:8]],
            "metas": goal_rows,
            "canales": [{"label": key, "value": _money(value)}
                        for key, value in sorted(by_channel.items(), key=lambda item: item[1], reverse=True)],
            "metodos_pago": [{
                "label": value["label"], "value": _money(value["ingresos"]),
                "pedidos": len(value["pedidos"]), "unidades": value["unidades"]
            } for value in sorted(by_payment.values(), key=lambda item: item["ingresos"], reverse=True)],
        },
        "table": table,
    }


def export_rows(filters: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    payload = load_source()
    return filter_sales(payload["sales"], filters or {})
