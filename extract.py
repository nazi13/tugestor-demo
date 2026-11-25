#!/usr/bin/env python3
"""Invoice extractor demo.

Reads invoice photos/PDFs from a folder, extracts the data with Claude, runs
sanity checks (NIF checksum, VAT arithmetic, dates, duplicates) and writes an
Excel file where doubtful rows are highlighted for human review.

    export ANTHROPIC_API_KEY=sk-...
    python extract.py ./facturas -o facturas.xlsx
    python extract.py ./facturas --mock      # canned data, no API key needed
"""
from __future__ import annotations

import argparse
import base64
import io
import os
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from validators import check_invoice, clean_tax_id

MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")
MAX_IMAGE_BYTES = 4_500_000
IMAGE_TYPES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
SUPPORTED = set(IMAGE_TYPES) | {".pdf"}

SYSTEM_PROMPT = """You extract accounting data from Spanish invoices, simplified invoices and receipts (photos or PDFs).
Rules:
- Copy tax IDs (NIF/CIF/NIE) exactly as printed. Never invent or "fix" characters.
- Dates in ISO format YYYY-MM-DD. Amounts as plain numbers with a dot as decimal separator (1.234,56 -> 1234.56).
- Add one entry to desglose_iva per VAT rate shown (base, tipo in %, cuota). Use tipo 0 for exempt lines.
- retencion_irpf is the positive amount withheld, if any.
- The supplier (proveedor) issues the document; the client (cliente) receives it.
- If a value is missing or unreadable, use null. Never guess; lower the confidence instead.
- If the document is not an invoice or receipt, set tipo_documento to "otro"."""

TOOL = {
    "name": "registrar_factura",
    "description": "Registra los datos extraídos de una factura, ticket o documento similar.",
    "input_schema": {
        "type": "object",
        "properties": {
            "tipo_documento": {"type": "string", "enum": ["factura", "factura_simplificada", "ticket", "otro"]},
            "numero_factura": {"type": ["string", "null"]},
            "fecha": {"type": ["string", "null"], "description": "YYYY-MM-DD"},
            "proveedor_nombre": {"type": ["string", "null"]},
            "proveedor_nif": {"type": ["string", "null"]},
            "cliente_nombre": {"type": ["string", "null"]},
            "cliente_nif": {"type": ["string", "null"]},
            "desglose_iva": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "base": {"type": "number"},
                        "tipo": {"type": "number", "description": "VAT rate in percent, e.g. 21"},
                        "cuota": {"type": "number"},
                    },
                    "required": ["base", "tipo", "cuota"],
                },
            },
            "retencion_irpf": {"type": ["number", "null"]},
            "total": {"type": ["number", "null"]},
            "confianza": {"type": "string", "enum": ["alta", "media", "baja"]},
            "notas": {"type": ["string", "null"]},
        },
        "required": ["tipo_documento", "desglose_iva", "total", "confianza"],
    },
}


# --------------------------------------------------------------------------- extraction


def _downscale(data: bytes) -> bytes:
    """Shrink large phone photos so they fit the API image size limit."""
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise RuntimeError("Imagen demasiado grande: instala pillow para reducirla") from exc
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    img.thumbnail((2400, 2400))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return buf.getvalue()


def build_content_block(path: Path) -> dict:
    raw = path.read_bytes()
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        media_type, kind = "application/pdf", "document"
    else:
        media_type, kind = IMAGE_TYPES[suffix], "image"
        if len(raw) > MAX_IMAGE_BYTES:
            raw, media_type = _downscale(raw), "image/jpeg"
    data = base64.standard_b64encode(raw).decode()
    return {"type": kind, "source": {"type": "base64", "media_type": media_type, "data": data}}


def extract_with_claude(path: Path, client) -> dict:
    response = client.messages.create(
        model=MODEL,
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        tools=[TOOL],
        tool_choice={"type": "tool", "name": TOOL["name"]},
        messages=[
            {
                "role": "user",
                "content": [build_content_block(path), {"type": "text", "text": "Extrae los datos de este documento."}],
            }
        ],
    )
    for block in response.content:
        if block.type == "tool_use":
            return block.input
    raise RuntimeError("La respuesta no contenía datos estructurados")


MOCK_INVOICES = [
    {  # clean invoice
        "tipo_documento": "factura", "numero_factura": "2025-0142", "fecha": "2025-10-14",
        "proveedor_nombre": "Suministros Ejemplo, S.L.", "proveedor_nif": "B12345674",
        "cliente_nombre": "Bar Ejemplo", "cliente_nif": "12345678Z",
        "desglose_iva": [{"base": 100.0, "tipo": 21, "cuota": 21.0}],
        "retencion_irpf": None, "total": 121.0, "confianza": "alta", "notas": None,
    },
    {  # bad supplier NIF + total does not add up
        "tipo_documento": "factura", "numero_factura": "A-889", "fecha": "2025-09-02",
        "proveedor_nombre": "Reformas Demo, S.L.", "proveedor_nif": "B12345679",
        "cliente_nombre": "Bar Ejemplo", "cliente_nif": "12345678Z",
        "desglose_iva": [{"base": 250.0, "tipo": 21, "cuota": 52.5}],
        "retencion_irpf": None, "total": 312.5, "confianza": "media", "notas": None,
    },
    {  # receipt without NIF (fine)
        "tipo_documento": "ticket", "numero_factura": None, "fecha": "2025-11-02",
        "proveedor_nombre": "Cafetería Ejemplo", "proveedor_nif": None,
        "cliente_nombre": None, "cliente_nif": None,
        "desglose_iva": [{"base": 9.09, "tipo": 10, "cuota": 0.91}],
        "retencion_irpf": None, "total": 10.0, "confianza": "alta", "notas": None,
    },
    {  # freelancer invoice with IRPF withholding
        "tipo_documento": "factura", "numero_factura": "F-7", "fecha": "2025-09-30",
        "proveedor_nombre": "María Ejemplo García", "proveedor_nif": "12345678Z",
        "cliente_nombre": "Bar Ejemplo", "cliente_nif": "B12345674",
        "desglose_iva": [{"base": 500.0, "tipo": 21, "cuota": 105.0}],
        "retencion_irpf": 75.0, "total": 530.0, "confianza": "media", "notas": None,
    },
]
MOCK_INVOICES.append(dict(MOCK_INVOICES[0]))  # same invoice twice -> duplicate check


def extract_mock(index: int) -> dict:
    return dict(MOCK_INVOICES[index % len(MOCK_INVOICES)])


# --------------------------------------------------------------------------- pipeline


def process_folder(folder: Path, use_mock: bool) -> list[dict]:
    files = sorted(p for p in folder.iterdir() if p.suffix.lower() in SUPPORTED)
    if not files:
        sys.exit(f"No hay imágenes ni PDF en {folder}")
    client = None
    if not use_mock:
        import anthropic

        client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY
    rows, seen = [], {}
    for i, path in enumerate(files):
        print(f"[{i + 1}/{len(files)}] {path.name}")
        try:
            data = extract_mock(i) if use_mock else extract_with_claude(path, client)
            issues = check_invoice(data)
        except Exception as exc:  # keep going: one bad file must not stop the batch
            data, issues = {}, [f"No se pudo procesar: {exc}"]
        key = (clean_tax_id(data.get("proveedor_nif")), data.get("numero_factura"))
        if key[0] and key[1]:
            if key in seen:
                issues.append(f"Posible duplicado de {seen[key]}")
            else:
                seen[key] = path.name
        rows.append({"archivo": path.name, "data": data, "issues": issues})
    return rows


# --------------------------------------------------------------------------- excel

HEADERS = [
    "Archivo", "Tipo", "Nº factura", "Fecha", "Proveedor", "NIF proveedor", "Cliente", "NIF cliente",
    "Base 21%", "IVA 21%", "Base 10%", "IVA 10%", "Base 4%", "IVA 4%", "Base 0%",
    "Otras bases", "Otras cuotas", "Retención IRPF", "Total", "Confianza", "Revisar",
]
MONEY_COLUMNS = range(9, 20)  # "Base 21%" .. "Total"
REVIEW_FILL = PatternFill("solid", start_color="FFF2CC")
HEADER_FILL = PatternFill("solid", start_color="1F3864")


def _money(x):
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return None


def split_vat(lines) -> dict:
    """Group VAT lines by rate: 21, 10, 4, 0 or 'other'. Returns {key: (base, fee)}."""
    by_rate: dict = {}
    for line in lines or []:
        try:
            rate, base, fee = float(line["tipo"]), float(line["base"]), float(line["cuota"])
        except (KeyError, TypeError, ValueError):
            continue
        key = rate if rate in (21.0, 10.0, 4.0, 0.0) else "other"
        old_base, old_fee = by_rate.get(key, (0.0, 0.0))
        by_rate[key] = (old_base + base, old_fee + fee)
    return by_rate


def build_row(item: dict) -> list:
    d = item["data"]
    vat = split_vat(d.get("desglose_iva"))
    base21, fee21 = vat.get(21.0, (None, None))
    base10, fee10 = vat.get(10.0, (None, None))
    base4, fee4 = vat.get(4.0, (None, None))
    base0, _ = vat.get(0.0, (None, None))
    other_base, other_fee = vat.get("other", (None, None))
    return [
        item["archivo"], d.get("tipo_documento"), d.get("numero_factura"), d.get("fecha"),
        d.get("proveedor_nombre"), d.get("proveedor_nif"), d.get("cliente_nombre"), d.get("cliente_nif"),
        _money(base21), _money(fee21), _money(base10), _money(fee10), _money(base4), _money(fee4),
        _money(base0), _money(other_base), _money(other_fee),
        _money(d.get("retencion_irpf")), _money(d.get("total")), d.get("confianza"),
        "; ".join(item["issues"]),
    ]


def write_excel(rows: list[dict], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for item in rows:
        ws.append(build_row(item))
        row_idx = ws.max_row
        for col in MONEY_COLUMNS:
            ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=len(HEADERS)).alignment = Alignment(wrap_text=True, vertical="top")
        if item["issues"]:
            for cell in ws[row_idx]:
                cell.fill = REVIEW_FILL
    widths = {1: 22, 2: 18, 3: 14, 4: 12, 5: 28, 6: 14, 7: 24, 8: 14, len(HEADERS): 60}
    for col in range(1, len(HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = widths.get(col, 12)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    total = len(rows)
    flagged = sum(1 for r in rows if r["issues"])
    summary = wb.create_sheet("Resumen")
    for line in [
        ("Documentos procesados", total),
        ("Sin incidencias", total - flagged),
        ("A revisar", flagged),
        ("% sin incidencias", round(100 * (total - flagged) / total, 1) if total else 0),
    ]:
        summary.append(line)
    summary.column_dimensions["A"].width = 28
    wb.save(out_path)


# --------------------------------------------------------------------------- cli


def main() -> None:
    ap = argparse.ArgumentParser(description="Extrae datos de facturas (fotos/PDF) a un Excel revisable.")
    ap.add_argument("folder", type=Path, help="Carpeta con facturas (jpg, png, webp, pdf)")
    ap.add_argument("-o", "--output", type=Path, default=Path("facturas.xlsx"))
    ap.add_argument("--mock", action="store_true", help="Usa datos de ejemplo, sin llamar a la API")
    args = ap.parse_args()

    if not args.folder.is_dir():
        sys.exit(f"No existe la carpeta: {args.folder}")
    if not args.mock and not os.environ.get("ANTHROPIC_API_KEY"):
        sys.exit("Falta ANTHROPIC_API_KEY (o usa --mock para probar sin API).")

    rows = process_folder(args.folder, args.mock)
    write_excel(rows, args.output)
    flagged = sum(1 for r in rows if r["issues"])
    print(f"\n{len(rows)} documentos -> {args.output}  ({flagged} a revisar, {len(rows) - flagged} sin incidencias)")


if __name__ == "__main__":
    main()
