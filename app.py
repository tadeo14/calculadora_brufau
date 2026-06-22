import streamlit as st
import datetime
from db import init_db, upsert_cotizacion, get_cotizaciones_hoy, get_proveedores, upsert_proveedor, delete_proveedor, replace_items, get_items
from io_xlsx import parse_excel, export_excel

# ─── Inicialización ──────────────────────────────────────────────────────────
init_db()

st.set_page_config(
    page_title="Brufau Sanitarios – Herramientas de Compras",
    page_icon="🔧",
    layout="centered",
)

st.markdown("""
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
    h3 { color: #1a3c6e; }
</style>
""", unsafe_allow_html=True)

st.title("🔧 Brufau Sanitarios – Herramientas de Compras")
st.caption("Área de Compras · Calculadora + Listas Dolarizadas")

# ─── Navegación principal ────────────────────────────────────────────────────
tab_calc, tab_dolar = st.tabs(["🧮 Calculadora", "💵 Listas Dolarizadas"])


# ════════════════════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════════════════════

def resultado(html):
    st.markdown(f'<div class="resultado">{html}</div>', unsafe_allow_html=True)

def error(msg):
    st.markdown(f'<div class="error-box">⚠️ {msg}</div>', unsafe_allow_html=True)

def parsear_float(txt, nombre):
    try:
        return float(txt.strip().replace(",", "."))
    except (ValueError, AttributeError):
        raise ValueError(f'El campo "{nombre}" no contiene un número válido.')


# ════════════════════════════════════════════════════════════════════════════
# TAB 1 – CALCULADORA
# ════════════════════════════════════════════════════════════════════════════
with tab_calc:

    # 1. Diferencia porcentual
    st.header("1 · Diferencia porcentual entre dos valores")
    st.markdown("Calcula cuánto varía **Valor B** respecto de **Valor A**.")
    c1, c2 = st.columns(2)
    txt_a = c1.text_input("Valor A (base / precio anterior)", key="dp_a")
    txt_b = c2.text_input("Valor B (precio nuevo)", key="dp_b")
    if st.button("Calcular diferencia porcentual", key="btn_dp"):
        try:
            if not txt_a or not txt_b:
                raise ValueError("Completá ambos campos.")
            a = parsear_float(txt_a, "Valor A")
            b = parsear_float(txt_b, "Valor B")
            if a == 0:
                raise ValueError("El Valor A no puede ser cero.")
            diff = ((b - a) / abs(a)) * 100
            signo = "+" if diff >= 0 else ""
            icono = "📈" if diff >= 0 else "📉"
            resultado(
                f"{icono} Diferencia: <strong>{signo}{diff:.2f} %</strong><br>"
                f"De <strong>{a:,.2f}</strong> a <strong>{b:,.2f}</strong>"
            )
        except ValueError as e:
            error(str(e))

    st.markdown("---")

    # 2. Porcentaje simple
    st.header("2 · Porcentaje de un valor")
    c3, c4 = st.columns(2)
    txt_val = c3.text_input("Valor base", key="ps_val")
    txt_pct = c4.text_input("Porcentaje (%)", key="ps_pct")
    if st.button("Calcular porcentaje", key="btn_ps"):
        try:
            if not txt_val or not txt_pct:
                raise ValueError("Completá ambos campos.")
            v = parsear_float(txt_val, "Valor base")
            p = parsear_float(txt_pct, "Porcentaje")
            if p < 0 or p > 100:
                raise ValueError("El porcentaje debe estar entre 0 y 100.")
            resultado(
                f"💡 El <strong>{p:.2f} %</strong> de <strong>{v:,.2f}</strong> "
                f"es <strong>{v*p/100:,.2f}</strong>"
            )
        except ValueError as e:
            error(str(e))

    st.markdown("---")

    # 3. Descuento escalonado
    st.header("3 · Descuento escalonado (bonificaciones)")
    st.markdown("Ingresá los descuentos separados por coma. Ej: `10,10,5`")
    c5, c6 = st.columns(2)
    txt_precio = c5.text_input("Precio / Lista original", key="de_precio")
    txt_desc   = c6.text_input("Descuentos separados por coma", key="de_desc", placeholder="10,10,5")
    if st.button("Calcular descuento escalonado", key="btn_de"):
        try:
            if not txt_precio or not txt_desc:
                raise ValueError("Completá ambos campos.")
            precio = parsear_float(txt_precio, "Precio original")
            if precio <= 0:
                raise ValueError("El precio debe ser mayor que cero.")
            descuentos = []
            for i, p in enumerate(txt_desc.strip().split(","), 1):
                d = parsear_float(p, f"descuento #{i}")
                if d < 0 or d >= 100:
                    raise ValueError(f"El descuento #{i} ({d}) debe estar entre 0 y 99,99 %.")
                descuentos.append(d)
            actual = precio
            pasos = []
            for d in descuentos:
                red = actual * d / 100
                actual -= red
                pasos.append((d, red))
            total_pct = (1 - actual / precio) * 100
            detalle = "".join(f"&nbsp;&nbsp;• {d:.2f} % → − {r:,.2f}<br>" for d, r in pasos)
            resultado(
                f"📋 <strong>Precio original:</strong> {precio:,.2f}<br>{detalle}"
                f"<hr style='margin:6px 0; border-color:#ccd'>"
                f"💰 <strong>Descuento total efectivo:</strong> {total_pct:.2f} %<br>"
                f"🔽 <strong>Ahorro:</strong> {precio-actual:,.2f}<br>"
                f"✅ <strong>Precio final:</strong> {actual:,.2f}"
            )
        except ValueError as e:
            error(str(e))


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
                    error(e_msg)

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
                error(str(e))

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
                        error("No se encontraron productos válidos en el archivo.")
                    else:
                        prov_id = nombres_prov[prov_sel]["id"]
                        replace_items(prov_id, items)
                        st.success(f"✅ {len(items)} productos importados para **{prov_sel}**.")
                        if advertencias:
                            with st.expander(f"⚠️ {len(advertencias)} advertencia(s)"):
                                for w in advertencias:
                                    st.write(w)
                except Exception as e:
                    error(f"Error al leer el archivo: {e}")

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
                    error(str(e))

st.markdown("---")
st.caption("Brufau Sanitarios · Herramientas internas de Compras")
