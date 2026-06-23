# Brufau Sanitarios – Herramientas de Compras

App web interna para el área de Compras de **Brufau Sanitarios**, construida en Python con Streamlit.  
Disponible desde cualquier navegador o celular, sin instalación.

---

## Módulos

### 🧮 Calculadora

Tres herramientas de uso diario para el área de compras:

| # | Función | Descripción |
|---|---|---|
| 1 | **Diferencia porcentual** | Calcula la variación % entre dos valores (aumento o reducción de precio) |
| 2 | **Porcentaje de un valor** | Responde "¿A es qué % de B?" |
| 3 | **Descuento escalonado** | Aplica bonificaciones en cascada (`10,10,5` → 23,05 % efectivo). El precio es **opcional**: si solo ingresás los descuentos muestra el % total; si ingresás precio original o precio final también calcula el otro extremo |

### 💵 Listas Dolarizadas

Gestión completa de listas de precios en USD y conversión a pesos argentinos.

**Flujo de trabajo:**

```
📊 Cotizaciones  →  🏭 Proveedores  →  📥 Importar lista  →  📤 Exportar ARS
```

| Pestaña | Qué hace |
|---|---|
| **Cotizaciones** | Registra el valor del día para OFICIAL, BLUE, MEP, CCL, DIVISA o MANUAL |
| **Proveedores** | Crea/edita proveedores con su tipo de cambio y % de recargo |
| **Importar lista** | Sube un Excel `.xlsx` con precios en USD (incluye botón para bajar la planilla de ejemplo) |
| **Exportar lista** | Genera y descarga un Excel en ARS con la fórmula: `precio_ARS = USD × cotización × (1 + recargo%)` |

---

## Tecnologías

| Componente | Tecnología |
|---|---|
| Lenguaje | Python 3 |
| Interfaz web | Streamlit |
| Base de datos | SQLite (archivo local `brufau.db`) |
| Excel (importar/exportar) | openpyxl |
| Deploy | Streamlit Community Cloud (gratuito) |

---

## Estructura de archivos

```
calculadora_brufau/
├── app.py              # App principal (UI y lógica de navegación)
├── db.py               # Base de datos SQLite: proveedores, cotizaciones, items
├── io_xlsx.py          # Importación y exportación de Excel + plantilla de ejemplo
├── requirements.txt    # Dependencias: streamlit, openpyxl
└── .gitignore          # Excluye __pycache__/ y brufau.db
```

---

## Base de datos (SQLite)

Tres tablas:

- **proveedores** — nombre, tipo de cambio, cotización manual, % de recargo.
- **cotizaciones** — snapshot diario de los seis tipos de cambio.
- **items** — SKU, descripción, precio USD, unidad, observaciones, vinculados a un proveedor.

---

## Deploy

### Local
```bash
pip install -r requirements.txt
streamlit run app.py
# Abre en http://localhost:8501
```

### Nube (Streamlit Community Cloud)
1. Subir el repo a GitHub.
2. Entrar a [share.streamlit.io](https://share.streamlit.io) con la cuenta de GitHub.
3. **New app** → seleccionar repo → rama `main` → archivo `app.py` → **Deploy**.
4. URL pública lista en ~2 minutos.

> ⚠️ **Nota:** En Streamlit Community Cloud la base de datos SQLite se reinicia al reiniciar el contenedor. Para datos persistentes en producción se recomienda migrar a una base en la nube (ej. Supabase).

---

## Ramas del repositorio

| Rama | Contenido |
|---|---|
| `main` | Calculadora mejorada (versión base) |
| `claude/dolarizados` | App completa: calculadora + módulo de Listas Dolarizadas |

**PR abierto:** `claude/dolarizados` → `main` ([PR #1](https://github.com/tadeo14/calculadora_brufau/pull/1))

---

## Mejoras futuras sugeridas

- Historial de cálculos exportable a Excel.
- Precio con IVA (10,5 % / 21 %) en el módulo de descuentos.
- Persistencia de la base de datos en la nube (Supabase / Google Drive).
- Autenticación con Google OAuth (disponible en Streamlit Community Cloud).
- Logo e identidad visual de Brufau Sanitarios.
