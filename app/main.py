"""FastAPI: expone el Excel como API y sirve el dashboard estático."""
from __future__ import annotations

import csv
import io
from datetime import date
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel, Field
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .chat import answer_chat
from .excel_reader import (
    DATA_FILE,
    ExcelSourceError,
    _as_date,
    _money,
    build_dashboard,
    export_rows,
    get_source_meta,
)

STATIC_DIR = Path(__file__).resolve().parents[1] / "static"

app = FastAPI(
    title="Puerto A&M · Business Intelligence API",
    description="Dashboard conectado a un archivo Excel de demostración (datos ficticios).",
    version="1.0.0",
)


@app.exception_handler(ExcelSourceError)
async def excel_error_handler(_request: Request, exc: ExcelSourceError) -> JSONResponse:
    return JSONResponse(status_code=502, content={"detail": str(exc)})


def _parse_filters(
    fecha_desde: str | None,
    fecha_hasta: str | None,
    producto: str | None,
    categoria: str | None,
    canal: str | None,
) -> dict[str, Any]:
    def parse_date(value: str | None, label: str) -> date | None:
        if not value:
            return None
        try:
            return _as_date(value)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=f"{label} no es una fecha válida (usa aaaa-mm-dd).") from exc

    return {
        "fecha_desde": parse_date(fecha_desde, "fecha_desde"),
        "fecha_hasta": parse_date(fecha_hasta, "fecha_hasta"),
        "producto": producto,
        "categoria": categoria,
        "canal": canal,
    }


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {"status": "ok", "excel_exists": DATA_FILE.exists(), "simulation": True}


@app.get("/api/meta")
def meta() -> dict[str, Any]:
    return get_source_meta()


@app.post("/api/refresh")
def refresh() -> dict[str, Any]:
    """Fuerza una nueva lectura del Excel y devuelve metadatos de la fuente."""
    return get_source_meta(force=True)


@app.get("/api/dashboard")
def dashboard(
    fecha_desde: str | None = Query(default=None),
    fecha_hasta: str | None = Query(default=None),
    producto: str | None = Query(default=None),
    categoria: str | None = Query(default=None),
    canal: str | None = Query(default=None),
) -> dict[str, Any]:
    return build_dashboard(_parse_filters(fecha_desde, fecha_hasta, producto, categoria, canal))


class ChatRequest(BaseModel):
    mensaje: str = Field(min_length=1, max_length=500)
    fecha_desde: str | None = None
    fecha_hasta: str | None = None
    producto: str | None = None
    categoria: str | None = None
    canal: str | None = None


@app.post("/api/chat")
def chat(request: ChatRequest) -> dict[str, Any]:
    """Responde usando exclusivamente los datos calculados desde el Excel."""
    filters = _parse_filters(
        request.fecha_desde, request.fecha_hasta,
        request.producto, request.categoria, request.canal,
    )
    return answer_chat(request.mensaje.strip(), build_dashboard(filters))


@app.get("/api/export")
def export_report(
    formato: str = Query(default="csv", pattern="^(csv|xlsx)$"),
    fecha_desde: str | None = Query(default=None),
    fecha_hasta: str | None = Query(default=None),
    producto: str | None = Query(default=None),
    categoria: str | None = Query(default=None),
    canal: str | None = Query(default=None),
) -> Response:
    filters = _parse_filters(fecha_desde, fecha_hasta, producto, categoria, canal)
    rows = export_rows(filters)
    filename_base = "PuertoAM_Reporte_Ventas"

    if formato == "csv":
        output = io.StringIO(newline="")
        writer = csv.writer(output, delimiter=",", quoting=csv.QUOTE_MINIMAL)
        writer.writerow(["Fecha", "Pedido", "Producto", "Categoría", "Cantidad",
                         "Precio unitario", "Ingreso", "Costo", "Utilidad", "Canal", "Método de pago"])
        for row in rows:
            writer.writerow([
                row["fecha"].isoformat(), row["pedido"], row["producto"], row["categoria"],
                row["cantidad"], _money(row["precio"]), _money(row["ingreso"]),
                _money(row["costo"]), _money(row["utilidad"]), row["canal"], row["pago"],
            ])
        content = "\ufeff" + output.getvalue()
        return Response(
            content=content.encode("utf-8"),
            media_type="text/csv; charset=utf-8",
            headers={"Content-Disposition": f'attachment; filename="{filename_base}.csv"'},
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte de ventas"
    headers = ["Fecha", "Pedido", "Producto", "Categoría", "Cantidad",
               "Precio unitario", "Ingreso", "Costo", "Utilidad", "Canal", "Método de pago"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", name="Aptos")
        cell.fill = PatternFill("solid", fgColor="0B3948")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for row in rows:
        ws.append([
            row["fecha"], row["pedido"], row["producto"], row["categoria"], row["cantidad"],
            _money(row["precio"]), _money(row["ingreso"]), _money(row["costo"]),
            _money(row["utilidad"]), row["canal"], row["pago"],
        ])
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=9):
        for cell in row:
            cell.number_format = '"$"#,##0.00'
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        for cell in row:
            cell.number_format = "dd/mm/yyyy"
    widths = [13, 18, 34, 16, 11, 16, 15, 15, 15, 19, 18]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{max(2, ws.max_row)}"
    ws.oddHeader.center.text = "PUERTO A&M · REPORTE SIMULADO · DATOS FICTICIOS"
    ws.oddFooter.center.text = "Top Learning · uso académico"

    buffer = io.BytesIO()
    wb.save(buffer)
    return Response(
        content=buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename_base}.xlsx"'},
    )


app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="dashboard")
