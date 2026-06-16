import streamlit as st

# ─── Configuración de página ────────────────────────────────────────────────
st.set_page_config(
    page_title="Calculadora de Compras – Brufau Sanitarios",
    page_icon="🧮",
    layout="centered",
)

# ─── Estilos personalizados ──────────────────────────────────────────────────
st.markdown(
    """
    <style>
        .resultado {
            background-color: #f0f7ff;
            border-left: 5px solid #1a73e8;
            padding: 14px 20px;
            border-radius: 6px;
            font-size: 1.1rem;
            margin-top: 12px;
        }
        .error-box {
            background-color: #fff0f0;
            border-left: 5px solid #d93025;
            padding: 10px 16px;
            border-radius: 6px;
            margin-top: 10px;
        }
        h1 { color: #1a3c6e; }
        h2 { color: #1a3c6e; border-bottom: 2px solid #1a73e8; padding-bottom: 4px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧮 Calculadora de Compras")
st.caption("Brufau Sanitarios – Área de Compras")
st.markdown("---")


# ─── Helpers ────────────────────────────────────────────────────────────────

def mostrar_resultado(html: str):
    st.markdown(f'<div class="resultado">{html}</div>', unsafe_allow_html=True)

def mostrar_error(msg: str):
    st.markdown(f'<div class="error-box">⚠️ {msg}</div>', unsafe_allow_html=True)

def parsear_float(texto: str, nombre: str):
    """Convierte texto a float; lanza ValueError con mensaje amigable."""
    texto = texto.strip().replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        raise ValueError(f'El campo "{nombre}" no contiene un número válido.')


# ════════════════════════════════════════════════════════════════════════════
# 1. DIFERENCIA PORCENTUAL
# ════════════════════════════════════════════════════════════════════════════
st.header("1 · Diferencia porcentual entre dos valores")
st.markdown(
    "Calcula cuánto representa **Valor B** respecto de **Valor A** "
    "(positivo = aumento, negativo = reducción)."
)

col1, col2 = st.columns(2)
with col1:
    txt_a = st.text_input("Valor A (base / precio anterior)", key="dp_a")
with col2:
    txt_b = st.text_input("Valor B (precio nuevo)", key="dp_b")

if st.button("Calcular diferencia porcentual", key="btn_dp"):
    try:
        if not txt_a or not txt_b:
            raise ValueError("Completá ambos campos antes de calcular.")
        a = parsear_float(txt_a, "Valor A")
        b = parsear_float(txt_b, "Valor B")
        if a == 0:
            raise ValueError("El Valor A no puede ser cero (división por cero).")
        diff = ((b - a) / abs(a)) * 100
        signo = "+" if diff >= 0 else ""
        icono = "📈" if diff >= 0 else "📉"
        mostrar_resultado(
            f"{icono} Diferencia porcentual: <strong>{signo}{diff:.2f} %</strong><br>"
            f"De <strong>{a:,.2f}</strong> a <strong>{b:,.2f}</strong>"
        )
    except ValueError as e:
        mostrar_error(str(e))

st.markdown("---")


# ════════════════════════════════════════════════════════════════════════════
# 2. CALCULADORA DE PORCENTAJE SIMPLE
# ════════════════════════════════════════════════════════════════════════════
st.header("2 · Porcentaje de un valor")
st.markdown("Obtené **X % de Y** de forma rápida.")

col3, col4 = st.columns(2)
with col3:
    txt_valor = st.text_input("Valor base", key="ps_val")
with col4:
    txt_pct = st.text_input("Porcentaje (%)", key="ps_pct")

if st.button("Calcular porcentaje", key="btn_ps"):
    try:
        if not txt_valor or not txt_pct:
            raise ValueError("Completá ambos campos antes de calcular.")
        valor = parsear_float(txt_valor, "Valor base")
        pct   = parsear_float(txt_pct, "Porcentaje")
        if pct < 0 or pct > 100:
            raise ValueError("El porcentaje debe estar entre 0 y 100.")
        resultado = valor * pct / 100
        mostrar_resultado(
            f"💡 El <strong>{pct:.2f} %</strong> de <strong>{valor:,.2f}</strong> "
            f"es <strong>{resultado:,.2f}</strong>"
        )
    except ValueError as e:
        mostrar_error(str(e))

st.markdown("---")


# ════════════════════════════════════════════════════════════════════════════
# 3. DESCUENTO ESCALONADO
# ════════════════════════════════════════════════════════════════════════════
st.header("3 · Descuento escalonado (bonificaciones)")
st.markdown(
    "Ingresá los descuentos separados por coma. Ejemplo: `10,10,5` "
    "aplica primero 10 %, luego 10 % sobre el resto, y luego 5 % sobre ese resto."
)

col5, col6 = st.columns(2)
with col5:
    txt_precio = st.text_input("Precio / Lista original", key="de_precio")
with col6:
    txt_desc = st.text_input("Descuentos separados por coma", key="de_desc",
                              placeholder="Ej: 10,10,5")

if st.button("Calcular descuento escalonado", key="btn_de"):
    try:
        if not txt_precio or not txt_desc:
            raise ValueError("Completá ambos campos antes de calcular.")

        precio_original = parsear_float(txt_precio, "Precio original")
        if precio_original <= 0:
            raise ValueError("El precio original debe ser mayor que cero.")

        # Parsear lista de descuentos
        partes = txt_desc.strip().split(",")
        descuentos = []
        for i, p in enumerate(partes, 1):
            d = parsear_float(p, f"descuento #{i}")
            if d < 0 or d >= 100:
                raise ValueError(
                    f"El descuento #{i} ({d}) debe estar entre 0 y 99,99 %."
                )
            descuentos.append(d)

        if not descuentos:
            raise ValueError("Ingresá al menos un descuento.")

        # Aplicar descuentos en cascada
        precio_actual = precio_original
        pasos = []
        for d in descuentos:
            reduccion = precio_actual * d / 100
            precio_actual -= reduccion
            pasos.append((d, reduccion))

        total_descuento_pct = (1 - precio_actual / precio_original) * 100
        ahorro_total = precio_original - precio_actual

        # Armar detalle de pasos
        detalle = "".join(
            f"&nbsp;&nbsp;• Descuento {d:.2f} % → − {r:,.2f}<br>"
            for d, r in pasos
        )

        mostrar_resultado(
            f"📋 <strong>Precio original:</strong> {precio_original:,.2f}<br>"
            f"{detalle}"
            f"<hr style='margin:6px 0; border-color:#ccd'>"
            f"💰 <strong>Descuento total efectivo:</strong> {total_descuento_pct:.2f} %<br>"
            f"🔽 <strong>Ahorro total:</strong> {ahorro_total:,.2f}<br>"
            f"✅ <strong>Precio final:</strong> {precio_actual:,.2f}"
        )

    except ValueError as e:
        mostrar_error(str(e))

st.markdown("---")
st.caption("Brufau Sanitarios · Calculadora interna de Compras")
