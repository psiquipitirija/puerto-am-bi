"""Genera una versión 100% estática del dashboard para GitHub Pages.

Los datos se leen del mismo Excel de siempre y se guardan en `site/data/dashboard.json`.
El navegador calcula filtros, KPIs, gráficas, exportaciones y respuestas del chat
con ese archivo: no necesita backend. Ejecutar desde la raíz del repo:

    python scripts/generar_estatico.py

Después publica la carpeta `site/` con:

    git add site && git commit -m "Actualiza sitio estático" && git push origin gh-pages
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.excel_reader import build_dashboard, load_source  # noqa: E402

SITE = ROOT / "site"
STATIC = ROOT / "static"


def main() -> None:
    source = load_source()
    base = build_dashboard()
    goals = [
        {
            "year": year,
            "month": month,
            "meta_ingresos": float(values["ingresos"]),
            "meta_pedidos": float(values["pedidos"]),
        }
        for (year, month), values in sorted(source["goals"].items())
    ]
    data = {
        "meta": base["meta"],
        "options": base["options"],
        "goals": goals,
        "rows": base["table"],
    }

    if SITE.exists():
        shutil.rmtree(SITE)
    (SITE / "data").mkdir(parents=True)
    (SITE / "vendor").mkdir()

    # El Excel original también se publica: queda en site/data/ y en el repo.
    excel = ROOT / "data" / "Puerto_AM_Base_Datos_DEMO.xlsx"
    shutil.copy2(excel, SITE / "data" / excel.name)

    for name in ("index.html", "styles.css", "app.js", "static-bridge.js"):
        shutil.copy2(STATIC / name, SITE / name)
    shutil.copy2(STATIC / "vendor" / "chart.umd.min.js", SITE / "vendor" / "chart.umd.min.js")

    vendor_xlsx = STATIC / "vendor" / "xlsx.full.min.js"
    if vendor_xlsx.exists():
        shutil.copy2(vendor_xlsx, SITE / "vendor" / "xlsx.full.min.js")

    html = (SITE / "index.html").read_text(encoding="utf-8")
    html = html.replace(
        '  <script src="vendor/chart.umd.min.js"></script>\n  <script src="app.js"></script>',
        '  <script src="vendor/chart.umd.min.js"></script>\n'
        '  <script src="vendor/xlsx.full.min.js"></script>\n'
        '  <script src="static-bridge.js"></script>\n'
        '  <script src="app.js"></script>',
    )
    (SITE / "index.html").write_text(html, encoding="utf-8")
    (SITE / "data" / "dashboard.json").write_text(
        json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8"
    )
    print(f"Sitio estático generado en {SITE}")
    print(f"  {len(data['rows'])} filas · {len(data['goals'])} metas · "
          f"{len(data['options']['productos'])} productos")


if __name__ == "__main__":
    main()
