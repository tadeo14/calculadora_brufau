"""
oc_processor.py
Lógica de negocio para el módulo de Órdenes de Compra.
Lee propuesta Flexxus + lista de precios, cruza datos y devuelve
una lista de filas lista para mostrar en st.data_editor.
"""

import io
import unicodedata
import openpyxl
from io_xlsx import norm, map_headers

# ── Aliases de columnas ──────────────────────────────────────────────────────

_ALIAS_PROPUESTA = {
    "codigo":     ["codigo", "cod", "sku", "ref", "articulo", "art"],
    "descripcion":["descripcion", "detalle", "nombre", "producto"],
    "proveedor":  ["proveedor", "prov", "razon", "razon social", "supplier"],
    "cantidad":   ["c opt", "copt", "cantidad", "cant", "sugerido", "sugerida",
                   "comprar", "a comprar", "cant a comprar", "cantidad sugerida",
                   "cant. a comprar", "c. opt."],
    "costo":      ["costo", "precio", "precio actual", "costo actual", "ultimo costo"],
    "stock":      ["stock", "existencia", "disponible", "saldo"],
    "sucursal":   ["sucursal", "deposito", "deposito", "sede", "almacen"],
}

_ALIAS_PRECIOS = {
    "codigo":      ["codigo", "cod", "sku", "ref", "articulo", "art"],
    # cod_prov se mapea ANTES que proveedor para evitar falsos positivos
    "cod_prov":    ["cod_prov", "cod prov", "codprov", "codigo prov", "cod. prov",
                   "codigo proveedor", "art prov", "articulo proveedor", "cod.prov"],
    "descripcion": ["descripcion", "detalle", "nombre", "producto"],
    # "proveedor" debe ser nombre largo, no "prov" suelto que matchea dentro de "cod_prov"
    "proveedor":   ["proveedor", "razon social", "razon", "supplier", "nombre proveedor"],
    "precio":      ["precio", "lista", "costo", "precio lista", "precio unitario",
                   "precio usd", "usd", "valor"],
    "moneda":      ["moneda", "currency", "tipo"],
    "bulto":       ["bulto", "bulto minimo", "unid x bulto", "x bulto", "cant bulto",
                   "minimo", "multiplo"],
    "activo":      ["activo", "habilitado", "estado", "vigente", "active"],
}


# ── Lectura de Excel ─────────────────────────────────────────────────────────

def leer_excel_raw(file_bytes: bytes) -> tuple[list[str], list[tuple]]:
    """Devuelve (headers, rows) de la primera hoja con datos."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("El archivo está vacío.")
    # Buscar primera fila con contenido (por si hay filas vacías arriba)
    header_row = 0
    for i, r in enumerate(rows):
        if any(c is not None for c in r):
            header_row = i
            break
    headers = [str(c).strip() if c is not None else "" for c in rows[header_row]]
    data = rows[header_row + 1:]
    if not headers:
        raise ValueError("No se encontraron encabezados en el archivo.")
    if not data:
        raise ValueError("El archivo no tiene filas de datos (solo encabezado).")
    return headers, data


def detectar_columnas_propuesta(headers: list[str]) -> dict:
    return map_headers(headers, _ALIAS_PROPUESTA)


def detectar_columnas_precios(headers: list[str]) -> dict:
    return map_headers(headers, _ALIAS_PRECIOS)


def validar_mapeo_propuesta(mapeo: dict) -> list[str]:
    errores = []
    if "codigo" not in mapeo:
        errores.append("No se encontró la columna de código de artículo en la propuesta.")
    if "cantidad" not in mapeo:
        errores.append("No se encontró una columna de cantidad sugerida en la propuesta "
                       "(buscadas: 'C Opt.', 'Cantidad', 'Sugerido', 'Comprar', etc.).")
    return errores


def validar_mapeo_precios(mapeo: dict) -> list[str]:
    errores = []
    if "codigo" not in mapeo:
        errores.append("No se encontró la columna de código de artículo en la lista de precios.")
    if "precio" not in mapeo:
        errores.append("La lista de precios no tiene precios válidos "
                       "(buscadas: 'Precio', 'Lista', 'Costo', 'Precio lista', etc.).")
    return errores


# ── Helpers ──────────────────────────────────────────────────────────────────

def _val(row, idx):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _str(row, idx) -> str:
    v = _val(row, idx)
    return str(v).strip() if v is not None else ""


def _float(row, idx) -> float | None:
    v = _val(row, idx)
    if v is None:
        return None
    try:
        return float(str(v).replace(",", ".").strip())
    except (ValueError, TypeError):
        return None


def _es_activo(row, idx) -> bool:
    """Devuelve True si la columna activo indica que el precio está vigente."""
    v = _val(row, idx)
    if v is None:
        return True  # sin columna = asumir activo
    n = norm(str(v))
    if n in ("0", "no", "false", "inactivo", "inhabilitado", "n"):
        return False
    return True


# ── Procesamiento principal ──────────────────────────────────────────────────

def procesar(
    propuesta_bytes: bytes,
    precios_bytes: bytes,
    mapeo_prop: dict,
    mapeo_prec: dict,
) -> tuple[list[dict], list[str]]:
    """
    Cruza propuesta con lista de precios.
    Devuelve (filas, advertencias_globales).
    Cada fila es un dict con todas las columnas que mostrará data_editor.
    """
    _, prop_rows = leer_excel_raw(propuesta_bytes)
    _, prec_rows = leer_excel_raw(precios_bytes)

    # ── Indexar precios por código ────────────────────────────────────────
    # precio_idx: {cod_normalizado: [lista de entradas]}
    precio_idx: dict[str, list[dict]] = {}
    for row in prec_rows:
        cod = norm(_str(row, mapeo_prec.get("codigo")))
        if not cod:
            continue
        precio = _float(row, mapeo_prec.get("precio"))
        activo = _es_activo(row, mapeo_prec.get("activo"))
        entrada = {
            "cod_prov":    _str(row, mapeo_prec.get("cod_prov")),
            "proveedor":   _str(row, mapeo_prec.get("proveedor")),
            "precio":      precio,
            "moneda":      _str(row, mapeo_prec.get("moneda")) or "ARS",
            "bulto":       _float(row, mapeo_prec.get("bulto")),
            "activo":      activo,
            "descripcion": _str(row, mapeo_prec.get("descripcion")),
        }
        precio_idx.setdefault(cod, []).append(entrada)

    advertencias_globales: list[str] = []
    filas: list[dict] = []

    for row in prop_rows:
        cod_orig = _str(row, mapeo_prop.get("codigo"))
        if not cod_orig:
            continue

        cant_sug = _float(row, mapeo_prop.get("cantidad"))
        if cant_sug is None or cant_sug <= 0:
            continue  # solo artículos con cantidad > 0

        cod_norm = norm(cod_orig)
        desc_prop = _str(row, mapeo_prop.get("descripcion"))
        prov_prop = _str(row, mapeo_prop.get("proveedor"))
        costo_act = _float(row, mapeo_prop.get("costo"))
        stock_act = _float(row, mapeo_prop.get("stock"))
        sucursal  = _str(row, mapeo_prop.get("sucursal"))

        # Buscar en lista de precios
        entradas = precio_idx.get(cod_norm, [])

        if not entradas:
            estado = "Sin precio"
            obs    = "No se encontró en la lista de precios."
            precio_u = None
            cod_prov = ""
            prov_sug = prov_prop
            bulto    = None
            moneda   = ""
        elif len(entradas) > 1:
            # Múltiples coincidencias → usar la primera activa, marcar duplicado
            activas = [e for e in entradas if e["activo"]]
            entrada = activas[0] if activas else entradas[0]
            estado  = "Revisar duplicado"
            obs     = f"Se encontraron {len(entradas)} entradas para este código."
            precio_u = entrada["precio"]
            cod_prov = entrada["cod_prov"]
            prov_sug = entrada["proveedor"] or prov_prop
            bulto    = entrada["bulto"]
            moneda   = entrada["moneda"]
        else:
            entrada  = entradas[0]
            precio_u = entrada["precio"]
            cod_prov = entrada["cod_prov"]
            prov_sug = entrada["proveedor"] or prov_prop
            bulto    = entrada["bulto"]
            moneda   = entrada["moneda"]
            if not entrada["activo"]:
                estado = "Precio inactivo"
                obs    = "El precio encontrado está marcado como inactivo."
            elif precio_u is None or precio_u <= 0:
                estado = "Sin precio"
                obs    = "La entrada existe pero no tiene precio válido."
            else:
                estado = "OK"
                obs    = ""

        # Sugerencia de bulto
        cant_bulto_sug = None
        if bulto and bulto > 1 and cant_sug:
            import math
            cant_bulto_sug = math.ceil(cant_sug / bulto) * bulto
            if cant_bulto_sug != cant_sug and estado == "OK":
                obs = f"Bulto mínimo {bulto:.0f} → se sugiere pedir {cant_bulto_sug:.0f}."

        precio_total = None
        if precio_u and cant_sug:
            precio_total = round(precio_u * cant_sug, 2)

        fila = {
            "seleccionar":       True,
            "codigo":            cod_orig,
            "descripcion":       desc_prop,
            "proveedor_sugerido":prov_sug,
            "proveedor_elegido": prov_sug,
            "cod_prov":          cod_prov,
            "stock_actual":      stock_act,
            "sucursal":          sucursal,
            "costo_actual":      costo_act,
            "cantidad_sugerida": cant_sug,
            "cant_bulto_sug":    cant_bulto_sug,
            "cantidad_final":    cant_bulto_sug if cant_bulto_sug else cant_sug,
            "bulto_minimo":      bulto,
            "precio_unitario":   precio_u,
            "moneda":            moneda,
            "precio_total":      precio_total,
            "estado":            estado,
            "observaciones":     obs,
        }
        filas.append(fila)

    if not filas:
        advertencias_globales.append(
            "No se encontraron artículos con cantidad sugerida mayor a cero."
        )

    sin_precio = sum(1 for f in filas if f["estado"] == "Sin precio")
    if sin_precio:
        advertencias_globales.append(
            f"Hay {sin_precio} artículo(s) sin precio. Revisalos antes de exportar."
        )

    return filas, advertencias_globales
