import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "brufau.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS proveedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL UNIQUE,
            tipo_cambio TEXT NOT NULL DEFAULT 'BLUE',
            cotizacion_manual REAL,
            recargo REAL NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS cotizaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL DEFAULT (date('now')),
            tipo TEXT NOT NULL,
            valor REAL NOT NULL,
            UNIQUE(fecha, tipo)
        );

        CREATE TABLE IF NOT EXISTS items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proveedor_id INTEGER NOT NULL,
            sku TEXT,
            descripcion TEXT,
            precio_usd REAL NOT NULL,
            moneda TEXT DEFAULT 'USD',
            unidad TEXT,
            observaciones TEXT,
            fecha_importacion TEXT DEFAULT (date('now')),
            FOREIGN KEY (proveedor_id) REFERENCES proveedores(id)
        );
    """)
    conn.commit()
    conn.close()


# ── Cotizaciones ─────────────────────────────────────────────────────────────

def upsert_cotizacion(tipo: str, valor: float, fecha: str = None):
    conn = get_conn()
    if fecha:
        conn.execute(
            "INSERT INTO cotizaciones (fecha, tipo, valor) VALUES (?, ?, ?) "
            "ON CONFLICT(fecha, tipo) DO UPDATE SET valor=excluded.valor",
            (fecha, tipo, valor),
        )
    else:
        conn.execute(
            "INSERT INTO cotizaciones (tipo, valor) VALUES (?, ?) "
            "ON CONFLICT(fecha, tipo) DO UPDATE SET valor=excluded.valor",
            (tipo, valor),
        )
    conn.commit()
    conn.close()


def get_cotizaciones_hoy():
    conn = get_conn()
    rows = conn.execute(
        "SELECT tipo, valor FROM cotizaciones WHERE fecha = date('now') ORDER BY tipo"
    ).fetchall()
    conn.close()
    return {r["tipo"]: r["valor"] for r in rows}


# ── Proveedores ───────────────────────────────────────────────────────────────

def get_proveedores():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM proveedores ORDER BY nombre").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def upsert_proveedor(nombre: str, tipo_cambio: str, recargo: float, cotizacion_manual: float = None):
    conn = get_conn()
    conn.execute(
        "INSERT INTO proveedores (nombre, tipo_cambio, recargo, cotizacion_manual) VALUES (?, ?, ?, ?) "
        "ON CONFLICT(nombre) DO UPDATE SET tipo_cambio=excluded.tipo_cambio, "
        "recargo=excluded.recargo, cotizacion_manual=excluded.cotizacion_manual",
        (nombre, tipo_cambio, recargo, cotizacion_manual),
    )
    conn.commit()
    conn.close()


def delete_proveedor(proveedor_id: int):
    conn = get_conn()
    conn.execute("DELETE FROM items WHERE proveedor_id = ?", (proveedor_id,))
    conn.execute("DELETE FROM proveedores WHERE id = ?", (proveedor_id,))
    conn.commit()
    conn.close()


# ── Items ─────────────────────────────────────────────────────────────────────

def replace_items(proveedor_id: int, items: list[dict]):
    conn = get_conn()
    conn.execute("DELETE FROM items WHERE proveedor_id = ?", (proveedor_id,))
    conn.executemany(
        "INSERT INTO items (proveedor_id, sku, descripcion, precio_usd, moneda, unidad, observaciones) "
        "VALUES (:proveedor_id, :sku, :descripcion, :precio_usd, :moneda, :unidad, :observaciones)",
        [{"proveedor_id": proveedor_id, **it} for it in items],
    )
    conn.commit()
    conn.close()


def get_items(proveedor_id: int):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM items WHERE proveedor_id = ? ORDER BY descripcion", (proveedor_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]
