import streamlit as st

# ─── Configuración de página ────────────────────────────────────────────────
st.set_page_config(
    page_title="Calculadora de Compras – Brufau Sanitarios",
    page_icon="🧮",
    layout="wide",
)

# ─── Estilos personalizados ──────────────────────────────────────────────────
st.markdown(
    """
    <style>
        .appview-container .main .block-container {
            padding: 10px 14px 10px;
        }
        .block-container .element-container {
            margin: 0 0 0.35rem 0;
            padding: 0;
        }
        .stTextInput>div, .stButton>button, .stTextArea>div {
            margin-bottom: 0.35rem;
        }
        .stButton>button {
            padding: 0.5rem 0.8rem;
            font-size: 0.92rem;
        }
        .resultado {
            background-color: #f0f7ff;
            border-left: 5px solid #1a73e8;
            color: #0f172a;
            padding: 10px 14px;
            border-radius: 6px;
            font-size: 0.95rem;
            margin-top: 8px;
            box-shadow: 0 2px 8px rgba(26, 115, 232, 0.06);
        }
        .resultado strong {
            color: #0f172a;
        }
        .error-box {
            background-color: #fff0f0;
            border-left: 5px solid #d93025;
            color: #710000;
            padding: 8px 12px;
            border-radius: 6px;
            margin-top: 8px;
            font-size: 0.95rem;
        }
        .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
            margin: 0.2rem 0 0.25rem 0;
        }
        .stMarkdown p {
            margin: 0.1rem 0 0.2rem 0;
        }
        header, footer {
            visibility: hidden;
            height: 0;
            margin: 0;
            padding: 0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧮 Calculadora de Compras")
st.caption("Brufau Sanitarios – Área de Compras")
st.markdown("Compacta, todo visible en una sola vista cuando hay suficiente ancho de pantalla.")


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
st.subheader("1 · Diferencia porcentual")
col1, col2, col3 = st.columns([1.2, 1.2, 0.7])
with col1:
    txt_a = st.text_input("Valor A (base / precio anterior)", key="dp_a")
with col2:
    txt_b = st.text_input("Valor B (precio nuevo)", key="dp_b")
with col3:
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    calcular_dp = st.button("Calcular", key="btn_dp")

if calcular_dp:
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


# ════════════════════════════════════════════════════════════════════════════
# 2. CALCULADORA DE PORCENTAJE SIMPLE
# ════════════════════════════════════════════════════════════════════════════
st.subheader("2 · ¿Qué porcentaje es un valor de otro?")
col4, col5, col6 = st.columns([1.2, 1.0, 0.7])
with col4:
    txt_valor = st.text_input("Valor A (el que querés porcentaje)", key="ps_val")
with col5:
    txt_base = st.text_input("Valor B (respecto a qué)", key="ps_base")
with col6:
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    calcular_ps = st.button("Calcular", key="btn_ps")

if calcular_ps:
    try:
        if not txt_valor or not txt_base:
            raise ValueError("Completá ambos campos antes de calcular.")
        valor = parsear_float(txt_valor, "Valor A")
        base = parsear_float(txt_base, "Valor B")
        if base == 0:
            raise ValueError("Valor B no puede ser cero (división por cero).")
        porcentaje = (valor / base) * 100
        mostrar_resultado(
            f"💡 <strong>{valor:,.2f}</strong> es el <strong>{porcentaje:.2f} %</strong> de <strong>{base:,.2f}</strong>"
        )
    except ValueError as e:
        mostrar_error(str(e))


# ════════════════════════════════════════════════════════════════════════════
# 3. DESCUENTO ESCALONADO
# ════════════════════════════════════════════════════════════════════════════
st.subheader("3 · Descuento escalonado")
col7, col8, col9, col10 = st.columns([1.0, 1.0, 1.5, 0.6])
with col7:
    txt_precio = st.text_input("Precio original (opcional)", key="de_precio")
with col8:
    txt_precio_final = st.text_input("Precio final (opcional)", key="de_precio_final")
with col9:
    txt_desc = st.text_input(
        "Descuentos separados por coma",
        key="de_desc",
        placeholder="Ej: 10,10,5",
    )
with col10:
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    calcular_de = st.button("Calcular", key="btn_de")

if calcular_de:
    try:
        if not txt_desc:
            raise ValueError("Ingresá al menos un descuento.")

        precio_original = None
        precio_final = None

        if txt_precio:
            precio_original = parsear_float(txt_precio, "Precio original")
            if precio_original <= 0:
                raise ValueError("El precio original debe ser mayor que cero.")

        if txt_precio_final:
            precio_final = parsear_float(txt_precio_final, "Precio final")
            if precio_final <= 0:
                raise ValueError("El precio final debe ser mayor que cero.")

        partes = [p.strip() for p in txt_desc.split(",") if p.strip()]
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

        factor_total = 1.0
        for d in descuentos:
            factor_total *= 1 - d / 100

        if precio_original is None and precio_final is None:
            precio_original = 1.0
            precio_actual = precio_original * factor_total
            nota_precio = "Precio relativo base 1,00"
        elif precio_original is None:
            precio_original = precio_final / factor_total
            precio_actual = precio_final
            nota_precio = (
                "Precio original calculado a partir del precio final y los descuentos"
            )
        else:
            precio_actual = precio_original * factor_total
            nota_precio = "Precio original usado para el cálculo"

        pasos = []
        precio_temp = precio_original
        for d in descuentos:
            reduccion = precio_temp * d / 100
            precio_temp -= reduccion
            pasos.append((d, reduccion, precio_temp))

        total_descuento_pct = (1 - factor_total) * 100
        ahorro_total = precio_original - precio_actual

        detalle = "".join(
            f"&nbsp;&nbsp;• Descuento {d:.2f} % → − {r:,.2f} (queda {n:,.2f})<br>"
            for d, r, n in pasos
        )

        html = (
            f"📋 <strong>{nota_precio}:</strong> {precio_original:,.2f}<br>"
            f"{detalle}"
            f"<hr style='margin:6px 0; border-color:#ccd'>"
            f"💰 <strong>Descuento total efectivo:</strong> {total_descuento_pct:.2f} %<br>"
            f"🔽 <strong>Ahorro total:</strong> {ahorro_total:,.2f}<br>"
            f"✅ <strong>Precio final:</strong> {precio_actual:,.2f}"
        )

        if precio_final is not None and txt_precio:
            esperado = precio_original * factor_total
            if abs(precio_final - esperado) > 1e-6:
                html += (
                    f"<br>⚠️ <strong>Nota:</strong> el precio final ingresado "
                    f"({precio_final:,.2f}) no coincide con el precio calculado "
                    f"({esperado:,.2f}) según los descuentos."
                )

        mostrar_resultado(html)

    except ValueError as e:
        mostrar_error(str(e))

st.caption("Brufau Sanitarios · Calculadora interna de Compras")
