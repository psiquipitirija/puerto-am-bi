"""Prueba funcional del asistente contra el servidor local.

Ejecutar con el servidor levantado:
    python scripts/probar_chat.py

Muestra la respuesta del endpoint /api/chat para preguntas representativas
y comprueba que las respuestas SIN_DATO no inventen información.
"""
from __future__ import annotations

import json
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

BASE = "http://127.0.0.1:8124"
LOCAL = "--local" in sys.argv

if LOCAL:
    from app.chat import answer_chat
    from app.excel_reader import build_dashboard

QUESTIONS = [
    "Hola",
    "¿Cuánto vendimos en total?",
    "¿Cuántos pedidos hay?",
    "¿Cuál es el ticket promedio?",
    "¿Cuál es la utilidad bruta?",
    "¿Cuál es el margen bruto?",
    "¿Cómo fue el crecimiento mensual?",
    "¿Cuánto vendimos en marzo?",
    "¿Cómo fue febrero comparado con enero?",
    "¿Cuál es el cumplimiento de metas?",
    "¿Cuáles son las categorías que más venden?",
    "¿Cuáles son los productos destacados?",
    "¿Cómo se distribuyen las ventas por canal?",
    "¿Cuáles son los métodos de pago?",
    "¿Cuánto vendió la categoría Oficina?",
    "¿Cuántas unidades se vendieron de Mochila ejecutiva impermeable?",
    "¿Cuál es la fuente del archivo?",
    "¿Quién es el cliente más importante?",
    "¿Qué empleados tienen el mejor sueldo?",
    "¿Cuánto venderemos en abril?",
]


def chat(mensaje: str, **filters: str) -> dict:
    if LOCAL:
        from app.excel_reader import filter_sales, load_source  # noqa: F401
        from app.excel_reader import _as_date
        parsed = {
            "fecha_desde": _as_date(filters["fecha_desde"]) if filters.get("fecha_desde") else None,
            "fecha_hasta": _as_date(filters["fecha_hasta"]) if filters.get("fecha_hasta") else None,
            "producto": filters.get("producto"),
            "categoria": filters.get("categoria"),
            "canal": filters.get("canal"),
        }
        return answer_chat(mensaje, build_dashboard(parsed))
    body = json.dumps({"mensaje": mensaje, **filters}).encode("utf-8")
    request = urllib.request.Request(
        BASE + "/api/chat",
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        return json.load(response)


def main() -> None:
    print("=" * 78)
    print("PRUEBA DEL ASISTENTE · Puerto A&M Copiloto")
    print("=" * 78)
    failures = []
    for question in QUESTIONS:
        result = chat(question)
        answer = result["respuesta"].replace("\n", " | ")
        print(f"\nP: {question}")
        print(f"R: {answer}")
        print(f"   modo={result['modo']} fuentes={result['fuentes']}")
        if question in {
            "¿Quién es el cliente más importante?",
            "¿Qué empleados tienen el mejor sueldo?",
            "¿Cuánto venderemos en abril?",
        } and "No tengo ese dato" not in result["respuesta"]:
            failures.append(question)

    print("\n" + "=" * 78)
    print("PRUEBA CON FILTRO ACTIVO (solo enero)")
    print("=" * 78)
    filtered = chat("¿Cuánto vendimos en total?", fecha_desde="2026-01-01", fecha_hasta="2026-01-31")
    print(filtered["respuesta"])
    assert "$177,387" in filtered["respuesta"], "El chat no respetó el filtro de enero"

    print("\nRESULTADO:", "OK" if not failures else f"FALLÓ: {failures}")
    if failures:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
