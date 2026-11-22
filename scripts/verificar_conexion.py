"""Prueba automatizada de la conexión real con el Excel.

1. Lee KPIs desde la API.
2. Modifica una cantidad en la hoja Ventas y guarda el archivo.
3. Fuerza la actualización y comprueba que el KPI cambia exactamente.
4. Regenera la base original para dejar el demo limpio.
"""
from __future__ import annotations

import json
import subprocess
import sys
import urllib.request
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "data" / "Puerto_AM_Base_Datos_DEMO.xlsx"
BASE = "http://127.0.0.1:8124"


def get_json(path: str, method: str = "GET"):
    request = urllib.request.Request(BASE + path, method=method)
    with urllib.request.urlopen(request, timeout=15) as response:
        return json.load(response)


def main() -> None:
    before = get_json("/api/dashboard")["kpis"]["ingresos"]
    wb = load_workbook(XLSX)
    ws = wb["Ventas"]
    old_qty = int(ws["E2"].value)
    price = float(ws["F2"].value)
    ws["E2"] = old_qty + 10
    wb.save(XLSX)
    expected = before + 10 * price

    meta = get_json("/api/refresh", method="POST")
    after = get_json("/api/dashboard")["kpis"]["ingresos"]

    print(f"Ingresos antes:        {before:,.2f}")
    print(f"Filas Ventas!E2:       {old_qty} -> {old_qty + 10} (precio {price:,.2f})")
    print(f"Ingresos esperados:    {expected:,.2f}")
    print(f"Ingresos después:      {after:,.2f}")
    print(f"Última modificación:   {meta['modified_label']}")
    ok = abs(after - expected) < 0.01 and after > before
    print("RESULTADO:", "CONEXIÓN REAL VERIFICADA" if ok else "FALLÓ")
    if not ok:
        sys.exit(1)

    subprocess.run([sys.executable, str(ROOT / "scripts" / "generar_excel.py")], check=True)
    get_json("/api/refresh", method="POST")
    restored = get_json("/api/dashboard")["kpis"]["ingresos"]
    print(f"Base regenerada:       {restored:,.2f}")
    assert abs(restored - before) < 0.01, "La regeneración no devolvió el valor original"


if __name__ == "__main__":
    main()
