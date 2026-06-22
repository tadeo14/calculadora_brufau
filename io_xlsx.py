import unicodedata
import io
import openpyxl


_ALIAS = {
    "sku":         ["sku", "codigo", "cod", "ref", "code"],
    "descripcion": ["descripcion", "detalle", "nombre", "producto", "description"],
    "precio_usd":  ["precio", "usd", "dolar", "price", "valor", "costo"],
    "moneda":      ["moneda", "currency"],
    "unidad":      ["unidad", "unit", "um"],
    "observaciones": ["observaciones", "obs", "nota", "notas"],
}


def _norm(text: str) -> str:
    """Minúsculas sin acentos."""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(text).lower())
        if unicodedata.category(c) != "Mn"
    )


def _map_headers(headers: list[str]) -> dict:
    """Devuelve {campo_interno: índice_columna}."""
    mapping = {}
    for idx, h in enumerate(headers):
        n = _norm(h)
        for campo, aliases in _ALIAS.items():
            if campo not in mapping and any(a in n for a in aliases):
                mapping[campo] = idx
    return mapping


def parse_excel(file_bytes: bytes) -> tuple[list[dict], list[str]]:
    """
    Parsea un Excel con lista de productos en USD.
    Retorna (items, advertencias).
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], ["El archivo está vacío."]

    headers = [str(c) if c is not None else "" for c in rows[0]]
    mapping = _map_headers(headers)

    if "precio_usd" not in mapping:
        return [], ["No se encontró columna de precio (USD/precio/valor/costo)."]

    items = []
    advertencias = []

    for i, row in enumerate(rows[1:], start=2):
        def col(campo):
            idx = mapping.get(campo)
            return row[idx] if idx is not None and idx < len(row) else None

        precio_raw = col("precio_usd")
        try:
            precio = float(precio_raw)
            if precio <= 0:
                raise ValueError
        except (TypeError, ValueError):
            advertencias.append(f"Fila {i}: precio inválido ({precio_raw!r}), ignorada.")
            continue

        sku = str(col("sku") or "").strip()
        desc = str(col("descripcion") or "").strip()
        if not sku and not desc:
            advertencias.append(f"Fila {i}: sin SKU ni descripción, ignorada.")
            continue

        items.append({
            "sku":           sku,
            "descripcion":   desc,
            "precio_usd":    precio,
            "moneda":        str(col("moneda") or "USD").strip(),
            "unidad":        str(col("unidad") or "").strip(),
            "observaciones": str(col("observaciones") or "").strip(),
        })

    return items, advertencias


def export_excel(items: list[dict], cotizacion: float, recargo: float) -> bytes:
    """Genera un Excel con precios convertidos a ARS."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lista ARS"

    headers = ["SKU", "Descripción", "Precio USD", "Cotización", "Recargo %", "Precio ARS", "Unidad", "Observaciones"]
    ws.append(headers)

    # Estilo de encabezado
    from openpyxl.styles import Font, PatternFill, Alignment
    header_fill = PatternFill("solid", fgColor="1A3C6E")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for it in items:
        precio_ars = it["precio_usd"] * cotizacion * (1 + recargo / 100)
        ws.append([
            it["sku"],
            it["descripcion"],
            it["precio_usd"],
            cotizacion,
            recargo,
            round(precio_ars, 2),
            it["unidad"],
            it["observaciones"],
        ])

    # Formato numérico
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=6):
        for cell in row:
            cell.number_format = "#,##0.00"

    # Ancho de columnas
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
