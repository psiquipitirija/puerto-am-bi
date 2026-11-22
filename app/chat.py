"""Asistente conversacional estrictamente fundamentado en el Excel.

Regla absoluta: cada cifra proviene del payload que build_dashboard() calcula
leyendo Puerto_AM_Base_Datos_DEMO.xlsx. Si el dato no existe en ese payload,
la respuesta lo dice explicitamente y no inventa nada.

Modo por defecto: motor local determinista (sin servicio externo, sin API key).
Opcional: si se define PUERTOAM_LLM_API_KEY, un modelo externo solo reescribe la
pregunta para entenderla mejor; NUNCA ve los datos del Excel y NUNCA da cifras.
Todos los números siguen siendo calculados por _local_answer() desde el payload.
Compatibles: OpenAI (gpt-4o-mini) y OpenRouter (sk-or-... / openai/gpt-4o-mini).
"""
from __future__ import annotations

import json
import os
import re
import unicodedata
import urllib.error
import urllib.request
from typing import Any

from .excel_reader import normalize

PERSONA = """Eres «A&M Copiloto», el asistente de Business Intelligence de Puerto A&M.

Tu personalidad: ejecutiva, cálida, breve, precisa y radicalmente honesta.
Hablas español de México. Usas importes en MXN como $1,234 y porcentajes con un decimal.
Tu tono es el de un analista senior que responde a dirección: vas al grano, citas la
fuente de cada número y nunca presumes datos que no están en el archivo.

REGLA ABSOLUTA (no negociable):
- Solo puedes usar los datos que aparecen en el JSON «DATOS DISPONIBLES».
- Nunca uses conocimiento general, estadísticas externas, suposiciones ni estimaciones.
- Si el dato solicitado no está en el JSON, responde exactamente con el texto SIN_DATO.
- No hagas proyecciones, pronósticos ni recomendaciones de inversión: reporta lo ocurrido.
- Cuando des un número, menciona brevemente de dónde sale (hoja Ventas, Productos o Metas).
- Si la pregunta es ambigua, pide precisión en lugar de adivinar.
- Máximo 5 renglones, salvo que pidan un desglose.
"""

SIN_DATO = (
    "No tengo ese dato en el Excel cargado. Solo respondo con información de las hojas "
    "Ventas, Productos y Metas. Puedo ayudarte con ventas, pedidos, ticket promedio, unidades, "
    "costos, utilidad, margen, crecimiento, metas, categorías, productos, canales, métodos de pago "
    "y con la venta más grande o más pequeña calculada de cada fila."
)

_MONTH_NAMES = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04", "mayo": "05", "junio": "06",
    "julio": "07", "agosto": "08", "septiembre": "09", "octubre": "10", "noviembre": "11",
    "diciembre": "12", "ene": "01", "feb": "02", "mar": "03", "abr": "04", "may": "05",
    "jun": "06", "jul": "07", "ago": "08", "sep": "09", "oct": "10", "nov": "11", "dic": "12",
}

_METRIC_WORDS = {
    "ingresos": ("ingresos", "ventas totales", "total vendido", "facturacion", "facturación",
                 "cuanto vendimos", "cuanto vendio", "cuanto se vendio", "vendimos", "vendio"),
    "pedidos": ("pedidos", "ordenes", "numero de pedidos", "cuantas ventas", "cuantos pedidos"),
    "ticket": ("ticket", "promedio por pedido", "venta promedio", "valor promedio"),
    "unidades": ("unidades", "cantidad vendida", "piezas", "articulos vendidos"),
    "costos": ("costos", "costo total", "cuanto costo", "gasto"),
    "utilidad": ("utilidad", "ganancia", "utilidad bruta", "margen en pesos"),
    "margen": ("margen", "rentabilidad", "porcentaje de ganancia"),
    "crecimiento": ("crecimiento", "variacion", "mes anterior", "evolucion", "subio", "bajo"),
    "metas": ("meta", "metas", "cumplimiento", "objetivo", "planificado", "plan"),
}


def money(value: float) -> str:
    return f"${float(value):,.0f}"


def percent(value: float | None) -> str:
    return "—" if value is None else f"{float(value):,.1f}%"


def number(value: float) -> str:
    return f"{float(value):,.0f}"


def plain(message: str) -> str:
    """Texto sin tildes pero CON espacios: permite límites de palabra fiables."""
    decomposed = unicodedata.normalize("NFD", str(message or ""))
    stripped = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    return stripped.lower()


def has(text: str, phrase: str) -> bool:
    """Compara contra texto normalizado (sin acentos ni espacios)."""
    return normalize(phrase) in text


def has_any(text: str, phrases) -> bool:
    return any(has(text, phrase) for phrase in phrases)


def _extreme_row(rows: list[dict[str, Any]], biggest: bool = True) -> dict[str, Any] | None:
    """Fila de mayor o menor ingreso dentro de las filas dadas."""
    if not rows:
        return None
    return max(rows, key=lambda row: row["ingreso"]) if biggest else min(rows, key=lambda row: row["ingreso"])


def build_evidence(data: dict[str, Any]) -> dict[str, Any]:
    """Contexto cerrado para el modelo externo: únicamente cifras del Excel."""
    return {
        "fuente": data["meta"]["file"],
        "ultima_lectura": data["meta"]["modified_label"],
        "filtros_activos": data["filters_applied"],
        "kpis": data["kpis"],
        "resumenes": {
            "por_mes": data["charts"]["tendencia"],
            "ingresos_y_costos": data["charts"]["ingresos_costos"],
            "por_categoria": data["charts"]["categorias"],
            "top_productos": data["charts"]["productos"],
            "metas": data["charts"]["metas"],
            "por_canal": data["charts"]["canales"],
            "por_metodo_pago": data["charts"].get("metodos_pago", []),
        },
        "opciones_de_filtro": data["options"],
        "venta_mayor": _extreme_row(data["table"], biggest=True),
        "venta_menor": _extreme_row(data["table"], biggest=False),
        "detalle_ventas": [
            {key: row[key] for key in ("fecha", "pedido", "producto", "categoria", "cantidad",
                                       "ingreso", "costo", "utilidad", "canal", "pago")}
            for row in data["table"][:40]
        ],
        "nota_detalle": (
            f"Se muestran 40 de {len(data['table'])} filas filtradas; además se incluyen la venta "
            "mayor y la menor calculadas sobre todas las filas. "
            "Si necesitas otra fila no mostrada, responde SIN_DATO."
        ),
        "SIN_DATO": SIN_DATO,
    }


def _entities(text: str, options: dict[str, Any]) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    groups = [
        ("producto", options.get("productos", [])),
        ("categoria", options.get("categorias", [])),
        ("canal", options.get("canales", [])),
        ("metodo de pago", options.get("metodos_pago", [])),
    ]
    for kind, values in groups:
        matches = [value for value in values if normalize(value) and normalize(value) in text]
        if matches:
            found.append((kind, max(matches, key=len)))
    return found


def _months(message: str) -> list[str]:
    text = plain(message)
    keys: list[str] = []
    # Los nombres completos primero: evita que "ene" gane sobre "enero".
    for word in sorted(_MONTH_NAMES, key=len, reverse=True):
        if re.search(rf"\b{re.escape(word)}\b", text):
            key = f"2026-{_MONTH_NAMES[word]}"
            if key not in keys:
                keys.append(key)
    for year_month in re.findall(r"\b(20\d{2})-(0[1-9]|1[0-2])\b", text):
        key = f"{year_month[0]}-{year_month[1]}"
        if key not in keys:
            keys.append(key)
    return keys


def _aggregate(table: list[dict[str, Any]], field: str, value: str) -> dict[str, float] | None:
    rows = [row for row in table if normalize(row[field]) == normalize(value)]
    if not rows:
        return None
    ingresos = sum(row["ingreso"] for row in rows)
    costos = sum(row["costo"] for row in rows)
    return {
        "ingresos": ingresos,
        "pedidos": len({row["pedido"] for row in rows}),
        "unidades": sum(row["cantidad"] for row in rows),
        "utilidad": ingresos - costos,
        "margen": (ingresos - costos) / ingresos * 100 if ingresos else 0,
    }


def _month_answer(months: list[str], data: dict[str, Any], message: str) -> dict[str, Any]:
    trend = {item["key"]: item for item in data["charts"]["tendencia"]}
    goals = {item["key"]: item for item in data["charts"]["metas"]}
    text = normalize(message)
    lines: list[str] = []
    sources = {"Ventas: Fecha, Cantidad × Precio unitario"}
    found = 0

    for key in months:
        item = trend.get(key)
        if not item:
            lines.append(f"No hay registros de {key} con los filtros actuales.")
            continue
        found += 1
        costs = sum(row["costo"] for row in data["table"] if row["fecha"].startswith(key))
        line = (f"{item['label']}: ingresos {money(item['ingresos'])}, {number(item['pedidos'])} pedidos, "
                f"utilidad {money(item['ingresos'] - costs)}.")
        goal = goals.get(key)
        if goal:
            line += f" Meta {money(goal['meta'])}: cumplimiento {percent(goal['cumplimiento'])}."
            sources.add("Metas: Meta de ingresos")
        lines.append(line)

    if found >= 2 and (has(text, "comparar") or has(text, "comparado") or has(text, "diferencia") or has(text, "versus")):
        values = [trend[key]["ingresos"] for key in months if key in trend]
        if len(values) >= 2 and min(values) > 0:
            lines.append(f"El mes mayor supera al menor por {money(max(values) - min(values))} "
                         f"({percent((max(values) / min(values) - 1) * 100)}).")

    if not found:
        lines.append(SIN_DATO)
    return {"respuesta": "\n".join(lines), "fuentes": sorted(sources)}


def _entity_answer(kind: str, value: str, data: dict[str, Any]) -> dict[str, Any]:
    field = {"producto": "producto", "categoria": "categoria", "canal": "canal",
             "metodo de pago": "pago"}[kind]
    result = _aggregate(data["table"], field, value)
    if result is None:
        return {"respuesta": SIN_DATO, "fuentes": []}
    label = {"producto": "Producto", "categoria": "Categoría", "canal": "Canal",
             "metodo de pago": "Método de pago"}[kind]
    respuesta = (
        f"{label} «{value}»: ingresos {money(result['ingresos'])}, {number(result['pedidos'])} pedidos, "
        f"{number(result['unidades'])} unidades, utilidad {money(result['utilidad'])} "
        f"y margen {percent(result['margen'])}."
    )
    return {
        "respuesta": respuesta,
        "fuentes": ["Ventas: " + label, "Ventas: Cantidad × Precio unitario y Costo unitario"],
    }


def _listing_answer(kind: str, data: dict[str, Any]) -> dict[str, Any]:
    chart_key = {"categoria": "categorias", "canal": "canales",
                 "metodo de pago": "metodos_pago", "producto": "productos"}[kind]
    items = data["charts"].get(chart_key, [])
    if not items:
        return {"respuesta": SIN_DATO, "fuentes": []}
    top = items[:5]
    title = {"categoria": "categorías", "canal": "canales", "metodo de pago": "métodos de pago",
             "producto": "productos"}[kind]
    lines = [f"Top {len(top)} {title} por ingresos (hoja Ventas):"]
    lines += [f"{index}. {item['label']} — {money(item['value'])}"
              for index, item in enumerate(top, start=1)]
    return {"respuesta": "\n".join(lines), "fuentes": ["Ventas: Cantidad × Precio unitario"]}


_SALE_FIELDS = {
    "producto": "producto", "categoria": "categoria", "canal": "canal",
    "metodo de pago": "pago",
}


def _scoped_rows(message: str, data: dict[str, Any]) -> tuple[list[dict[str, Any]], str, bool]:
    """Filas del Excel aplicando mes y/o entidad mencionados en la pregunta.

    Devuelve (filas, texto de alcance, sin_dato). El alcance se usa para decir
    explícitamente si la respuesta es, por ejemplo, de marzo o de una categoría.
    """
    rows = list(data.get("table") or [])
    scope: list[str] = []
    months = _months(message)
    if months:
        rows = [row for row in rows if any(row["fecha"].startswith(key) for key in months)]
        if not rows:
            return [], "", True
        labels = {item["key"]: item["label"] for item in data["charts"].get("tendencia", [])}
        scope.append("en " + ", ".join(labels.get(key, key) for key in months))

    # Solo aplicamos la entidad si la pregunta la introduce con "de", "del",
    # "en", "por" o "para": evita tomar "venta" como si fuera un producto.
    if re.search(r"\b(de|del|en|por|para)\b", plain(message)):
        entities = _entities(normalize(message), data["options"])
        if entities:
            kind, value = max(entities, key=lambda item: len(normalize(item[1])))
            field = _SALE_FIELDS[kind]
            filtered = [row for row in rows if normalize(row[field]) == normalize(value)]
            if not filtered:
                return [], "", True
            rows = filtered
            scope.append(f"de {kind} «{value}»")
    return rows, (" " + " ".join(scope) if scope else ""), False


def _sale_extreme_answer(message: str, data: dict[str, Any], biggest: bool = True) -> dict[str, Any]:
    """Venta individual de mayor o menor ingreso, calculada desde la hoja Ventas."""
    rows, scope, missing = _scoped_rows(message, data)
    if missing or not rows:
        return {"respuesta": SIN_DATO, "fuentes": []}
    row = _extreme_row(rows, biggest=biggest)
    word = "mayor" if biggest else "menor"
    respuesta = (
        f"La venta {word} del periodo{scope}: {money(row['ingreso'])} "
        f"(pedido {row['pedido']} del {row['fecha']}, {row['producto']}, "
        f"categoría {row['categoria']}, {number(row['cantidad'])} unidades, "
        f"canal {row['canal']}, pago {row['pago']}; hoja Ventas)."
    )
    return {
        "respuesta": respuesta,
        "fuentes": [
            "Ventas: Pedido, Fecha, Producto, Categoría y Cantidad",
            "Ventas: Cantidad × Precio unitario",
        ],
    }


def _top_sales_answer(message: str, data: dict[str, Any]) -> dict[str, Any]:
    """Listado de las ventas individuales más altas, calculado desde la hoja Ventas."""
    rows, scope, missing = _scoped_rows(message, data)
    if missing or not rows:
        return {"respuesta": SIN_DATO, "fuentes": []}
    top = sorted(rows, key=lambda row: row["ingreso"], reverse=True)[:5]
    lines = [f"Top {len(top)} ventas por ingreso{scope} (hoja Ventas):"]
    lines += [
        f"{index}. {row['pedido']} · {row['fecha']} · {money(row['ingreso'])} — {row['producto']}"
        for index, row in enumerate(top, start=1)
    ]
    return {
        "respuesta": "\n".join(lines),
        "fuentes": ["Ventas: Pedido, Fecha, Producto y Cantidad × Precio unitario"],
    }


_BIG_SALE_PHRASES = (
    "venta mas grande", "venta mayor", "mayor venta", "pedido mas grande",
    "pedido mayor", "pedido mas caro", "pedido mas alto", "monto mas alto",
    "monto mas grande", "importe mas alto", "venta mas alta", "venta mas cara",
    "facturo mas", "venta facturo mas", "ingreso mas alto", "venta maxima",
    "record de ventas",
)
_SMALL_SALE_PHRASES = (
    "venta mas pequena", "venta mas baja", "menor venta", "venta minima",
    "pedido mas pequeno", "monto mas bajo", "ingreso mas bajo",
)
_TOP_SALES_PHRASES = (
    "top ventas", "top de ventas", "top pedidos", "top de pedidos",
    "top 5 ventas", "top 5 de ventas", "las 5 ventas mas altas",
    "cinco ventas mas grandes", "mejores ventas", "ventas mas altas",
    "ventas mas grandes",
)


def _local_answer(message: str, data: dict[str, Any]) -> dict[str, Any]:
    text = normalize(message)
    if not text:
        return {"respuesta": "Escribe una pregunta sobre los datos del dashboard y respondo solo con el Excel.", "fuentes": []}
    k = data["kpis"]
    trend = data["charts"]["tendencia"]

    # 1. Personalidad: saludo y ayuda.
    if has_any(text, ("hola", "buenos dias", "buenas tardes", "que tal")) and len(text) < 24:
        return {
            "respuesta": "Hola, soy A&M Copiloto. Respondo únicamente con los datos del Excel cargado. "
                         "Pregúntame por ventas, pedidos, metas, categorías, productos o canales.",
            "fuentes": [],
        }
    if has_any(text, ("ayuda", "que puedes hacer", "como funciona", "que sabes")):
        return {
            "respuesta": "Puedo responder sobre ingresos, pedidos, ticket promedio, unidades, costos, "
                         "utilidad, margen, crecimiento mensual, metas, categorías, productos, canales, "
                         "métodos de pago y la venta más grande o más pequeña. Si el dato no está en el "
                         "Excel, te lo digo sin inventarlo.",
            "fuentes": [],
        }

    # 2. Metadatos de la fuente.
    if has_any(text, ("archivo", "fuente", "actualizado", "actualizacion", "modificacion", "Excel")) \
            and not has_any(text, tuple(word for words in _METRIC_WORDS.values() for word in words)):
        meta = data["meta"]
        return {
            "respuesta": (f"Fuente: {meta['file']}. Última lectura: {meta['modified_label']}. "
                          f"Contiene {number(meta['rows_total'])} filas en Ventas, {meta['products_total']} productos "
                          f"y {meta['goals_total']} metas. "
                          f"Estado: {'caché vigente' if meta.get('cache_hit') else 'lectura nueva'}."),
            "fuentes": ["Metadatos del archivo Excel"],
        }

    # 3. Ventas individuales: se calculan fila por fila sobre la hoja Ventas.
    #    Van antes que los meses para permitir "venta más grande en marzo".
    if has_any(text, _BIG_SALE_PHRASES):
        return _sale_extreme_answer(message, data, biggest=True)
    if has_any(text, _SMALL_SALE_PHRASES):
        return _sale_extreme_answer(message, data, biggest=False)
    if has_any(text, _TOP_SALES_PHRASES):
        return _top_sales_answer(message, data)

    # 4. Meses explícitos.
    months = _months(message)
    if months:
        return _month_answer(months, data, message)

    # 5. Entidades concretas (producto / categoría / canal / método de pago).
    entities = _entities(text, data["options"])
    if entities:
        kind, value = max(entities, key=lambda item: len(normalize(item[1])))
        return _entity_answer(kind, value, data)

    # 6. Listados y ranking.
    if has_any(text, ("top", "mas vendido", "destacado", "ranking", "mejor vendido")):
        return _listing_answer("producto", data)
    if has_any(text, ("categoria", "categorias")):
        return _listing_answer("categoria", data)
    if has_any(text, ("canal", "canales")):
        return _listing_answer("canal", data)
    if has_any(text, ("metodo de pago", "metodos de pago", "forma de pago")):
        return _listing_answer("metodo de pago", data)

    # 7. Mejor / peor mes.
    if has_any(text, ("cual mes", "que mes", "mejor mes", "peor mes")):
        if not trend:
            return {"respuesta": SIN_DATO, "fuentes": []}
        best = max(trend, key=lambda item: item["ingresos"])
        worst = min(trend, key=lambda item: item["ingresos"])
        return {
            "respuesta": (f"El mes con más ingresos fue {best['label']} con {money(best['ingresos'])}. "
                          f"El de menos fue {worst['label']} con {money(worst['ingresos'])}. "
                          f"Diferencia: {money(best['ingresos'] - worst['ingresos'])} (hoja Ventas)."),
            "fuentes": ["Ventas: Fecha, Cantidad × Precio unitario"],
        }

    # 8. Temas fuera del Excel: responder SIN_DATO antes que cualquier métrica.
    if has_any(text, ("cliente", "empleado", "sueldo", "impuesto", "predice", "pronostico",
                      "proyeccion", "futuro", "invertir", "opinion", "recomiendas")):
        return {"respuesta": SIN_DATO, "fuentes": []}

    # 9. Métricas globales del periodo filtrado.
    for metric, words in _METRIC_WORDS.items():
        if has_any(text, words):
            if metric == "ingresos":
                return {"respuesta": f"Ingresos totales: {money(k['ingresos'])} en el periodo filtrado "
                                     f"({number(k['filas_filtradas'])} registros de la hoja Ventas).",
                        "fuentes": ["Ventas: Cantidad × Precio unitario"]}
            if metric == "pedidos":
                return {"respuesta": f"Pedidos: {number(k['pedidos'])} en el periodo filtrado (hoja Ventas).",
                        "fuentes": ["Ventas: Pedido"]}
            if metric == "ticket":
                return {"respuesta": f"Ticket promedio: {money(k['ticket_promedio'])} por pedido "
                                     f"(ingresos ÷ pedidos, hoja Ventas).",
                        "fuentes": ["Ventas: Cantidad × Precio unitario", "Ventas: Pedido"]}
            if metric == "unidades":
                return {"respuesta": f"Unidades vendidas: {number(k['unidades'])} (hoja Ventas, Cantidad).",
                        "fuentes": ["Ventas: Cantidad"]}
            if metric == "costos":
                return {"respuesta": f"Costos totales: {money(k['costos'])} (Cantidad × Costo unitario, hoja Ventas).",
                        "fuentes": ["Ventas: Cantidad × Costo unitario"]}
            if metric == "utilidad":
                return {"respuesta": f"Utilidad bruta: {money(k['utilidad'])} (ingresos menos costos, hoja Ventas).",
                        "fuentes": ["Ventas: Cantidad × Precio unitario y Costo unitario"]}
            if metric == "margen":
                return {"respuesta": f"Margen bruto: {percent(k['margen'])} (utilidad ÷ ingresos, hoja Ventas).",
                        "fuentes": ["Ventas: Cantidad × Precio unitario y Costo unitario"]}
            if metric == "crecimiento":
                if k["crecimiento"] is None:
                    return {"respuesta": "No puedo calcular el crecimiento con los filtros actuales: "
                                         "se necesitan al menos dos meses en el periodo (hoja Ventas).",
                            "fuentes": []}
                detail = "; ".join(f"{item['label']}: {money(item['ingresos'])}" for item in trend)
                return {"respuesta": f"Crecimiento del último mes frente al anterior: {percent(k['crecimiento'])}. {detail}.",
                        "fuentes": ["Ventas: Fecha, Cantidad × Precio unitario"]}
            if metric == "metas":
                goals = data["charts"]["metas"]
                if not goals or k["cumplimiento_metas"] is None:
                    return {"respuesta": "No hay metas mensuales en el Excel para el periodo filtrado.",
                            "fuentes": []}
                detail = "; ".join(f"{item['label']}: {percent(item['cumplimiento'])}" for item in goals)
                return {"respuesta": f"Cumplimiento global de metas: {percent(k['cumplimiento_metas'])}. Por mes — {detail}.",
                        "fuentes": ["Metas: Meta de ingresos", "Ventas: Cantidad × Precio unitario"]}

    return {"respuesta": SIN_DATO, "fuentes": []}


_INTERPRETER_SYSTEM = """Eres un normalizador de preguntas para el asistente de un dashboard de ventas.
NO respondas la pregunta. Solo reescríbela de forma clara, breve y canónica en español de México,
conservando la intención y los filtros explícitos (mes, categoría, producto, canal, forma de pago).
No agregues cifras, nombres ni datos que el usuario no escribió.

Usa estas frases canónicas cuando aplique:
- "¿Cuánto vendimos en total?"
- "¿Cuántos pedidos hay?"
- "¿Cuál es el ticket promedio?"
- "¿Cuántas unidades se vendieron?"
- "¿Cuáles son los costos totales?"
- "¿Cuál es la utilidad bruta?"
- "¿Cuál es el margen bruto?"
- "¿Cómo fue el crecimiento mensual?"
- "¿Cuál es el cumplimiento de metas?"
- "¿Cuál fue la venta más grande?" (también para "la más alta", "el pedido más caro", "el pedido más alto", "el monto mayor", "el importe más alto", "lo que más facturó", "el récord")
- "¿Cuál fue la venta más pequeña?"
- "Top ventas" (también para "top 5 de ventas", "las ventas más altas", "los mejores pedidos"; NUNCA para productos ni categorías)
- "¿Cuáles son las categorías que más venden?"
- "¿Cuáles son los productos destacados?"
- "¿Cómo se distribuyen las ventas por canal?"
- "¿Cuáles son los métodos de pago?"
- "¿Cuál fue el mejor mes?"

Si el usuario menciona un mes, inclúyelo como "en marzo", "en febrero", etc.
Si menciona una categoría, producto, canal o forma de pago, inclúyelo como "de <nombre>".
Si pide algo ajeno a ventas, productos o metas (clientes, empleados, sueldos, pronósticos, opiniones),
devuelve exactamente SIN_DATO. Devuelve SOLO la pregunta reescrita, sin comillas ni explicación."""


def _llm_endpoint() -> str:
    configured = os.getenv("PUERTOAM_LLM_BASE_URL")
    if configured:
        return configured
    key = os.getenv("PUERTOAM_LLM_API_KEY", "")
    return ("https://openrouter.ai/api/v1/chat/completions" if key.startswith("sk-or-")
            else "https://api.openai.com/v1/chat/completions")


def _llm_model() -> str:
    configured = os.getenv("PUERTOAM_LLM_MODEL")
    if configured:
        return configured
    key = os.getenv("PUERTOAM_LLM_API_KEY", "")
    return "openai/gpt-4o-mini" if key.startswith("sk-or-") else "gpt-4o-mini"


def _interpret_message(message: str) -> str | None:
    """Opcional: el modelo solo interpreta la intención; nunca ve los datos del Excel."""
    api_key = os.getenv("PUERTOAM_LLM_API_KEY")
    if not api_key:
        return None
    body = json.dumps({
        "model": _llm_model(),
        "temperature": 0,
        "max_tokens": 120,
        "messages": [
            {"role": "system", "content": _INTERPRETER_SYSTEM},
            {"role": "user", "content": message},
        ],
    }).encode("utf-8")
    request = urllib.request.Request(
        _llm_endpoint(),
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://puerto-am-bi.local",
            "X-Title": "Puerto A&M Copiloto",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            payload = json.load(response)
        content = payload["choices"][0]["message"]["content"].strip().strip('"').strip()
    except (urllib.error.URLError, KeyError, IndexError, json.JSONDecodeError, OSError):
        return None
    if not content or content.upper().startswith("SIN_DATO"):
        return None
    return content[:300]


def answer_chat(message: str, data: dict[str, Any]) -> dict[str, Any]:
    """Los números siempre los calcula el motor local; el LLM solo ayuda a entender la pregunta."""
    base = _local_answer(message, data)
    interpreted = _interpret_message(message)
    result = base
    mode = "motor local sobre datos del Excel"

    if interpreted:
        candidate = _local_answer(interpreted, data)
        base_failed = "No tengo ese dato" in base["respuesta"] or not base["fuentes"]
        candidate_ok = "No tengo ese dato" not in candidate["respuesta"] and bool(candidate["fuentes"])
        if base_failed and candidate_ok:
            result = candidate
        mode = "interpretación con modelo + cálculo local sobre el Excel"

    return {
        "respuesta": result["respuesta"],
        "fuentes": result["fuentes"],
        "modo": mode,
        "verificado": True,
    }
