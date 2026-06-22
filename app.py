import streamlit as st
import datetime
from db import init_db, upsert_cotizacion, get_cotizaciones_hoy, get_proveedores, upsert_proveedor, delete_proveedor, replace_items, get_items
from io_xlsx import parse_excel, export_excel

# ─── Inicialización ──────────────────────────────────────────────────────────
init_db()

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


st.title("🔧 Brufau Sanitarios – Herramientas de Compras")
st.caption("Área de Compras · Calculadora + Listas Dolarizadas")

# ─── Navegación principal ────────────────────────────────────────────────────
tab_calc, tab_dolar = st.tabs(["🧮 Calculadora", "💵 Listas Dolarizadas"])


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
    except (ValueError, AttributeError):
        raise ValueError(f'El campo "{nombre}" no contiene un número válido.')


# ════════════════════════════════════════════════════════════════════════════
# TAB 1 – CALCULADORA
# ════════════════════════════════════════════════════════════════════════════
with tab_calc:

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



# ════════════════════════════════════════════════════════════════════════════
# TAB 2 – LISTAS DOLARIZADAS
# ════════════════════════════════════════════════════════════════════════════
TIPOS_CAMBIO = ["OFICIAL", "BLUE", "MEP", "CCL", "DIVISA", "MANUAL"]

with tab_dolar:
    sub_cot, sub_prov, sub_import, sub_export = st.tabs(
        ["📊 Cotizaciones", "🏭 Proveedores", "📥 Importar lista", "📤 Exportar lista"]
    )

    # ── Cotizaciones ─────────────────────────────────────────────────────────
    with sub_cot:
        st.header("Cotizaciones del día")
        cotizaciones_hoy = get_cotizaciones_hoy()
        st.caption(f"Fecha: {datetime.date.today().strftime('%d/%m/%Y')}")

        with st.form("form_cotizaciones"):
            cols = st.columns(3)
            valores = {}
            for i, tipo in enumerate(TIPOS_CAMBIO):
                val_actual = cotizaciones_hoy.get(tipo, "")
                valores[tipo] = cols[i % 3].text_input(
                    f"{tipo}",
                    value=str(val_actual) if val_actual else "",
                    placeholder="0.00",
                    key=f"cot_{tipo}",
                )
            guardar = st.form_submit_button("💾 Guardar cotizaciones")

        if guardar:
            errores = []
            guardados = []
            for tipo, txt in valores.items():
                if not txt.strip():
                    continue
                try:
                    v = parsear_float(txt, tipo)
                    if v <= 0:
                        raise ValueError("debe ser mayor que cero")
                    upsert_cotizacion(tipo, v)
                    guardados.append(f"{tipo}: {v:,.2f}")
                except ValueError as e:
                    errores.append(f"{tipo}: {e}")
            if guardados:
                st.success("Guardado: " + " · ".join(guardados))
            if errores:
                for e_msg in errores:
                    mostrar_error(e_msg)

        # Tabla resumen
        cots = get_cotizaciones_hoy()
        if cots:
            st.markdown("**Cotizaciones registradas hoy:**")
            cols_h = st.columns(len(cots))
            for i, (tipo, val) in enumerate(cots.items()):
                cols_h[i].metric(tipo, f"$ {val:,.2f}")

    # ── Proveedores ───────────────────────────────────────────────────────────
    with sub_prov:
        st.header("Gestión de proveedores")

        proveedores = get_proveedores()
        if proveedores:
            st.markdown("**Proveedores registrados:**")
            for p in proveedores:
                col_n, col_tc, col_rec, col_del = st.columns([3, 2, 2, 1])
                col_n.write(f"**{p['nombre']}**")
                col_tc.write(p["tipo_cambio"])
                col_rec.write(f"Recargo: {p['recargo']} %")
                if col_del.button("🗑️", key=f"del_prov_{p['id']}"):
                    delete_proveedor(p["id"])
                    st.rerun()
            st.markdown("---")

        st.subheader("Agregar / actualizar proveedor")
        with st.form("form_proveedor"):
            nombre_p  = st.text_input("Nombre del proveedor")
            tipo_c    = st.selectbox("Tipo de cambio", TIPOS_CAMBIO, index=1)
            recargo_p = st.text_input("Recargo (%)", value="0", placeholder="Ej: 15")
            cot_manual = st.text_input(
                "Cotización manual (solo si elegiste MANUAL)",
                placeholder="Ej: 1250.50",
            )
            guardar_p = st.form_submit_button("💾 Guardar proveedor")

        if guardar_p:
            try:
                if not nombre_p.strip():
                    raise ValueError("El nombre del proveedor es obligatorio.")
                rec = parsear_float(recargo_p, "Recargo")
                if rec < 0:
                    raise ValueError("El recargo no puede ser negativo.")
                cot_m = None
                if tipo_c == "MANUAL":
                    if not cot_manual.strip():
                        raise ValueError("Ingresá la cotización manual.")
                    cot_m = parsear_float(cot_manual, "Cotización manual")
                    if cot_m <= 0:
                        raise ValueError("La cotización manual debe ser mayor que cero.")
                upsert_proveedor(nombre_p.strip(), tipo_c, rec, cot_m)
                st.success(f"Proveedor '{nombre_p.strip()}' guardado.")
                st.rerun()
            except ValueError as e:
                mostrar_error(str(e))

    # ── Importar lista ────────────────────────────────────────────────────────
    with sub_import:
        st.header("Importar lista de productos (Excel USD)")

        proveedores = get_proveedores()
        if not proveedores:
            st.info("Primero creá un proveedor en la pestaña **Proveedores**.")
        else:
            nombres_prov = {p["nombre"]: p for p in proveedores}
            prov_sel = st.selectbox("Proveedor", list(nombres_prov.keys()), key="imp_prov")
            archivo = st.file_uploader("Subí el archivo Excel (.xlsx)", type=["xlsx"])

            if archivo and st.button("📥 Importar", key="btn_import"):
                try:
                    items, advertencias = parse_excel(archivo.read())
                    if not items:
                        mostrar_error("No se encontraron productos válidos en el archivo.")
                    else:
                        prov_id = nombres_prov[prov_sel]["id"]
                        replace_items(prov_id, items)
                        st.success(f"✅ {len(items)} productos importados para **{prov_sel}**.")
                        if advertencias:
                            with st.expander(f"⚠️ {len(advertencias)} advertencia(s)"):
                                for w in advertencias:
                                    st.write(w)
                except Exception as e:
                    mostrar_error(f"Error al leer el archivo: {e}")

    # ── Exportar lista ────────────────────────────────────────────────────────
    with sub_export:
        st.header("Exportar lista convertida a ARS")

        proveedores = get_proveedores()
        if not proveedores:
            st.info("Primero creá un proveedor e importá una lista.")
        else:
            nombres_prov = {p["nombre"]: p for p in proveedores}
            prov_sel_e = st.selectbox("Proveedor", list(nombres_prov.keys()), key="exp_prov")
            prov_data  = nombres_prov[prov_sel_e]

            # Resolver cotización
            cots_hoy = get_cotizaciones_hoy()
            tipo_c_prov = prov_data["tipo_cambio"]

            if tipo_c_prov == "MANUAL":
                cot_resuelta = prov_data.get("cotizacion_manual") or 0.0
            else:
                cot_resuelta = cots_hoy.get(tipo_c_prov, 0.0)

            col_info1, col_info2 = st.columns(2)
            col_info1.metric("Tipo de cambio", tipo_c_prov)
            col_info2.metric("Cotización", f"$ {cot_resuelta:,.2f}" if cot_resuelta else "⚠️ No cargada")

            recargo_e = st.text_input(
                "Recargo (%)",
                value=str(prov_data["recargo"]),
                key="exp_recargo",
            )

            items = get_items(prov_data["id"])
            if items:
                st.markdown(f"**{len(items)} productos** en la lista de {prov_sel_e}.")
                with st.expander("Vista previa (primeros 10)"):
                    for it in items[:10]:
                        st.write(f"• {it['descripcion']} — USD {it['precio_usd']:,.2f}")
            else:
                st.info("Este proveedor no tiene productos. Importá una lista primero.")

            if st.button("📤 Generar Excel ARS", key="btn_export") and items:
                try:
                    if cot_resuelta <= 0:
                        raise ValueError(
                            f"La cotización {tipo_c_prov} no está cargada hoy. "
                            "Ingresala en la pestaña Cotizaciones."
                        )
                    rec = parsear_float(recargo_e, "Recargo")
                    if rec < 0:
                        raise ValueError("El recargo no puede ser negativo.")

                    excel_bytes = export_excel(items, cot_resuelta, rec)
                    fecha_str = datetime.date.today().strftime("%Y%m%d")
                    nombre_archivo = f"lista_pesos_{prov_sel_e}_{fecha_str}.xlsx"

                    st.download_button(
                        label="⬇️ Descargar lista ARS",
                        data=excel_bytes,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.success(f"Excel generado: {nombre_archivo}")
                except ValueError as e:
                    mostrar_error(str(e))

st.markdown("---")

st.caption("Brufau Sanitarios · Herramientas internas de Compras")
