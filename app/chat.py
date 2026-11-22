"""Asistente conversacional estrictamente fundamentado en el Excel.

Regla absoluta: cada cifra proviene del payload que build_dashboard() calcula
leyendo Puerto_AM_Base_Datos_DEMO.xlsx. Si el dato no existe en ese payload,
la respuesta lo dice explicitamente y no inventa nada.

Modo por defecto: motor local determinista (sin servicio externo, sin API key).
Opcional: modelo externo con contexto cerrado si se define PUERTOAM_LLM_API_KEY.
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
    "costos, utilidad, margen, crecimiento, metas, categorías, productos, canales y métodos de pago."
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
        "detalle_ventas": [
            {key: row[key] for key in ("fecha", "pedido", "producto", "categoria", "cantidad",
                                       "ingreso", "costo", "utilidad", "canal", "pago")}
            for row in data["table"][:40]
        ],
        "nota_detalle": (
            f"Se muestran 40 de {len(data['table'])} filas filtradas. "
            "Si necesitas un dato de una fila no mostrada, responde SIN_DATO."
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
                         "utilidad, margen, crecimiento mensual, metas, categorías, productos, canales "
                         "y métodos de pago. Si el dato no está en el Excel, te lo digo sin inventarlo.",
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

    # 3. Meses explícitos.
    months = _months(message)
    if months:
        return _month_answer(months, data, message)

    # 4. Entidades concretas (producto / categoría / canal / método de pago).
    entities = _entities(text, data["options"])
    if entities:
        kind, value = max(entities, key=lambda item: len(normalize(item[1])))
        return _entity_answer(kind, value, data)

    # 5. Listados y ranking.
    if has_any(text, ("top", "mas vendido", "destacado", "ranking", "mejor vendido")):
        return _listing_answer("producto", data)
    if has_any(text, ("categoria", "categorias")):
        return _listing_answer("categoria", data)
    if has_any(text, ("canal", "canales")):
        return _listing_answer("canal", data)
    if has_any(text, ("metodo de pago", "metodos de pago", "forma de pago")):
        return _listing_answer("metodo de pago", data)

    # 6. Mejor / peor mes.
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

    # 7. Temas fuera del Excel: responder SIN_DATO antes que cualquier métrica.
    if has_any(text, ("cliente", "empleado", "sueldo", "impuesto", "predice", "pronostico",
                      "proyeccion", "futuro", "invertir", "opinion", "recomiendas")):
        return {"respuesta": SIN_DATO, "fuentes": []}

    # 8. Métricas globales del periodo filtrado.
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


def _llm_answer(message: str, data: dict[str, Any]) -> str | None:
    """Opcional: modelo externo con contexto cerrado. Sin clave no se llama."""
    api_key = os.getenv("PUERTOAM_LLM_API_KEY")
    if not api_key:
        return None
    model = os.getenv("PUERTOAM_LLM_MODEL", "gpt-4o-mini")
    system = PERSONA + "\n\nDATOS DISPONIBLES (JSON):\n" + json.dumps(build_evidence(data), ensure_ascii=False)
    body = json.dumps({
        "model": model,
        "temperature": 0,
        "max_tokens": 400,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": message},
        ],
    }).encode("utf-8")
    request = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=body,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            payload = json.load(response)
        content = payload["choices"][0]["message"]["content"].strip()
        return content or None
    except (urllib.error.URLError, KeyError, IndexError, json.JSONDecodeError):
        return None


def answer_chat(message: str, data: dict[str, Any]) -> dict[str, Any]:
    llm = _llm_answer(message, data)
    if llm:
        return {
            "respuesta": llm,
            "fuentes": ["Excel · " + data["meta"]["file"]],
            "modo": "modelo externo con contexto cerrado",
            "verificado": False,
        }
    result = _local_answer(message, data)
    return {
        "respuesta": result["respuesta"],
        "fuentes": result["fuentes"],
        "modo": "motor local sobre datos del Excel",
        "verificado": True,
    }
