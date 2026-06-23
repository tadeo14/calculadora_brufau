import streamlit as st
import datetime
from db import init_db, upsert_cotizacion, get_cotizaciones_hoy, get_proveedores, upsert_proveedor, delete_proveedor, replace_items, get_items
from io_xlsx import parse_excel, export_excel, plantilla_excel
from oc_processor import (
    leer_excel_raw, detectar_columnas_propuesta, detectar_columnas_precios,
    validar_mapeo_propuesta, validar_mapeo_precios, procesar,
)
from oc_export import exportar_un_proveedor, exportar_todas, nombre_archivo_oc

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
st.caption("Área de Compras · Calculadora + Listas Dolarizadas + Órdenes de Compra")

# ─── Navegación principal ────────────────────────────────────────────────────
tab_calc, tab_dolar, tab_oc = st.tabs(["🧮 Calculadora", "💵 Listas Dolarizadas", "📋 Órdenes de Compra"])


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

        st.download_button(
            label="⬇️ Descargar planilla de ejemplo",
            data=plantilla_excel(),
            file_name="plantilla_lista_USD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descargá esta planilla, completala con tus productos y volvé a subirla.",
        )
        st.markdown("---")

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


# ════════════════════════════════════════════════════════════════════════════
# TAB 3 – ÓRDENES DE COMPRA
# ════════════════════════════════════════════════════════════════════════════
with tab_oc:
    st.header("Generador de Órdenes de Compra")
    st.caption(
        "Cargá una propuesta de compra y una lista de precios para revisar, "
        "ajustar y exportar órdenes de compra agrupadas por proveedor."
    )
    st.markdown("---")

    # ── BLOQUE 1: Carga de archivos ──────────────────────────────────────────
    st.subheader("1 · Carga de archivos")
    col_prop, col_prec = st.columns(2)
    with col_prop:
        archivo_prop = st.file_uploader(
            "📄 Propuesta de compra (Flexxus) — .xlsx",
            type=["xlsx"],
            key="oc_prop",
        )
    with col_prec:
        archivo_prec = st.file_uploader(
            "💲 Lista de precios del proveedor — .xlsx",
            type=["xlsx"],
            key="oc_prec",
        )

    procesar_btn = st.button("⚙️ Procesar archivos", key="btn_oc_procesar",
                             disabled=not (archivo_prop and archivo_prec))

    if procesar_btn:
        st.session_state.pop("oc_resultado", None)
        st.session_state.pop("oc_mapeo_prop", None)
        st.session_state.pop("oc_mapeo_prec", None)
        try:
            prop_bytes = archivo_prop.read()
            prec_bytes = archivo_prec.read()
            headers_prop, _ = leer_excel_raw(prop_bytes)
            headers_prec, _ = leer_excel_raw(prec_bytes)
            st.session_state["oc_prop_bytes"]    = prop_bytes
            st.session_state["oc_prec_bytes"]    = prec_bytes
            st.session_state["oc_headers_prop"]  = headers_prop
            st.session_state["oc_headers_prec"]  = headers_prec
            st.session_state["oc_mapeo_prop"]    = detectar_columnas_propuesta(headers_prop)
            st.session_state["oc_mapeo_prec"]    = detectar_columnas_precios(headers_prec)
            st.success("Archivos cargados. Revisá la detección de columnas antes de procesar.")
        except Exception as e:
            mostrar_error(f"Error al leer los archivos: {e}")

    # ── BLOQUE 2: Configuración de columnas ──────────────────────────────────
    if "oc_headers_prop" in st.session_state:
        st.markdown("---")
        st.subheader("2 · Verificación de columnas detectadas")
        st.caption("Si alguna columna no fue detectada correctamente, corregila manualmente.")

        headers_prop = st.session_state["oc_headers_prop"]
        headers_prec = st.session_state["oc_headers_prec"]
        mapeo_prop   = st.session_state["oc_mapeo_prop"]
        mapeo_prec   = st.session_state["oc_mapeo_prec"]

        opciones_prop = ["(no usar)"] + headers_prop
        opciones_prec = ["(no usar)"] + headers_prec

        def sel_col(label, campo, mapeo, opciones, key):
            idx_actual = mapeo.get(campo)
            # +1 porque opciones[0] = "(no usar)"
            default = (idx_actual + 1) if idx_actual is not None else 0
            sel = st.selectbox(label, opciones, index=default, key=key)
            return opciones.index(sel) - 1 if sel != "(no usar)" else None

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Propuesta de compra**")
            mapeo_prop["codigo"]      = sel_col("Código artículo *", "codigo", mapeo_prop, opciones_prop, "mp_cod")
            mapeo_prop["descripcion"] = sel_col("Descripción", "descripcion", mapeo_prop, opciones_prop, "mp_desc")
            mapeo_prop["proveedor"]   = sel_col("Proveedor", "proveedor", mapeo_prop, opciones_prop, "mp_prov")
            mapeo_prop["cantidad"]    = sel_col("Cantidad sugerida *", "cantidad", mapeo_prop, opciones_prop, "mp_cant")
            mapeo_prop["costo"]       = sel_col("Costo actual", "costo", mapeo_prop, opciones_prop, "mp_cost")
            mapeo_prop["stock"]       = sel_col("Stock actual", "stock", mapeo_prop, opciones_prop, "mp_stock")
            mapeo_prop["sucursal"]    = sel_col("Sucursal/depósito", "sucursal", mapeo_prop, opciones_prop, "mp_suc")

        with col_b:
            st.markdown("**Lista de precios**")
            mapeo_prec["codigo"]      = sel_col("Código artículo *", "codigo", mapeo_prec, opciones_prec, "mc_cod")
            mapeo_prec["cod_prov"]    = sel_col("Código proveedor", "cod_prov", mapeo_prec, opciones_prec, "mc_cprov")
            mapeo_prec["descripcion"] = sel_col("Descripción", "descripcion", mapeo_prec, opciones_prec, "mc_desc")
            mapeo_prec["proveedor"]   = sel_col("Proveedor", "proveedor", mapeo_prec, opciones_prec, "mc_prov")
            mapeo_prec["precio"]      = sel_col("Precio *", "precio", mapeo_prec, opciones_prec, "mc_prec")
            mapeo_prec["moneda"]      = sel_col("Moneda", "moneda", mapeo_prec, opciones_prec, "mc_mon")
            mapeo_prec["bulto"]       = sel_col("Bulto mínimo", "bulto", mapeo_prec, opciones_prec, "mc_bulto")
            mapeo_prec["activo"]      = sel_col("Activo/Inactivo", "activo", mapeo_prec, opciones_prec, "mc_activo")

        # Validar y cruzar
        errores_prop = validar_mapeo_propuesta(mapeo_prop)
        errores_prec = validar_mapeo_precios(mapeo_prec)
        for e in errores_prop + errores_prec:
            mostrar_error(e)

        if not errores_prop and not errores_prec:
            if st.button("🔀 Cruzar datos", key="btn_oc_cruzar"):
                try:
                    filas, advertencias = procesar(
                        st.session_state["oc_prop_bytes"],
                        st.session_state["oc_prec_bytes"],
                        mapeo_prop,
                        mapeo_prec,
                    )
                    st.session_state["oc_resultado"] = filas
                    for w in advertencias:
                        st.warning(w)
                    if filas:
                        st.success(f"✅ {len(filas)} artículos procesados.")
                    else:
                        mostrar_error("No se encontraron artículos con cantidad sugerida mayor a cero.")
                except Exception as e:
                    mostrar_error(f"Error al procesar los datos: {e}")

    # ── BLOQUE 3: Tabla de revisión ──────────────────────────────────────────
    if "oc_resultado" in st.session_state:
        filas = st.session_state["oc_resultado"]
        st.markdown("---")
        st.subheader("3 · Revisión de artículos")

        # Filtros
        col_f1, col_f2, col_f3 = st.columns(3)
        proveedores_lista = sorted(set(
            (f.get("proveedor_elegido") or f.get("proveedor_sugerido") or "Sin proveedor")
            for f in filas
        ))
        filtro_prov = col_f1.selectbox("Filtrar por proveedor", ["Todos"] + proveedores_lista, key="oc_fprov")
        filtro_sp   = col_f2.checkbox("Solo sin precio", key="oc_fsp")
        filtro_adv  = col_f3.checkbox("Solo con advertencias", key="oc_fadv")

        filas_vis = filas
        if filtro_prov != "Todos":
            filas_vis = [f for f in filas_vis
                         if (f.get("proveedor_elegido") or f.get("proveedor_sugerido") or "") == filtro_prov]
        if filtro_sp:
            filas_vis = [f for f in filas_vis if f["estado"] == "Sin precio"]
        if filtro_adv:
            filas_vis = [f for f in filas_vis if f["estado"] not in ("OK", "")]

        st.caption(f"Mostrando {len(filas_vis)} de {len(filas)} artículos")

        # Columnas a mostrar en data_editor
        import pandas as pd
        df = pd.DataFrame(filas_vis)[[
            "seleccionar", "codigo", "descripcion",
            "proveedor_sugerido", "proveedor_elegido", "cod_prov",
            "cantidad_sugerida", "cant_bulto_sug", "cantidad_final",
            "bulto_minimo", "precio_unitario", "precio_total",
            "moneda", "estado", "observaciones",
        ]]

        col_config = {
            "seleccionar":       st.column_config.CheckboxColumn("✔", width="small"),
            "codigo":            st.column_config.TextColumn("Código", width="medium"),
            "descripcion":       st.column_config.TextColumn("Descripción", width="large"),
            "proveedor_sugerido":st.column_config.TextColumn("Prov. sugerido", width="medium"),
            "proveedor_elegido": st.column_config.TextColumn("Prov. elegido", width="medium"),
            "cod_prov":          st.column_config.TextColumn("Cód. prov.", width="small"),
            "cantidad_sugerida": st.column_config.NumberColumn("Cant. sug.", format="%.0f", width="small"),
            "cant_bulto_sug":    st.column_config.NumberColumn("Sug. bulto", format="%.0f", width="small"),
            "cantidad_final":    st.column_config.NumberColumn("Cant. final", format="%.0f", width="small"),
            "bulto_minimo":      st.column_config.NumberColumn("Bulto", format="%.0f", width="small"),
            "precio_unitario":   st.column_config.NumberColumn("P. unitario", format="%.2f", width="small"),
            "precio_total":      st.column_config.NumberColumn("P. total", format="%.2f", width="small"),
            "moneda":            st.column_config.TextColumn("Moneda", width="small"),
            "estado":            st.column_config.TextColumn("Estado", width="medium"),
            "observaciones":     st.column_config.TextColumn("Observaciones", width="large"),
        }

        df_editado = st.data_editor(
            df,
            column_config=col_config,
            disabled=[c for c in df.columns if c not in ("seleccionar", "cantidad_final", "proveedor_elegido")],
            use_container_width=True,
            hide_index=True,
            key="oc_editor",
        )

        # Aplicar cambios al estado global
        for i, row in df_editado.iterrows():
            filas_vis[i]["seleccionar"]    = bool(row["seleccionar"])
            filas_vis[i]["cantidad_final"] = row["cantidad_final"]
            filas_vis[i]["proveedor_elegido"] = row["proveedor_elegido"]
            # Recalcular precio total
            pu = filas_vis[i].get("precio_unitario")
            cf = filas_vis[i]["cantidad_final"]
            filas_vis[i]["precio_total"] = round(pu * cf, 2) if pu and cf else None

        # Métricas de totales
        sel = [f for f in filas_vis if f.get("seleccionar")]
        total_items = len(sel)
        total_uds   = sum(f.get("cantidad_final") or 0 for f in sel)
        total_imp   = sum(f.get("precio_total") or 0 for f in sel if f.get("estado") != "Sin precio")
        c1, c2, c3 = st.columns(3)
        c1.metric("Artículos seleccionados", total_items)
        c2.metric("Unidades totales", f"{total_uds:,.0f}")
        c3.metric("Importe estimado", f"{total_imp:,.2f}")

        # ── Resumen por proveedor ─────────────────────────────────────────
        st.markdown("---")
        st.subheader("4 · Resumen por proveedor")

        resumen: dict[str, dict] = {}
        for f in filas_vis:
            if not f.get("seleccionar"):
                continue
            prov = (f.get("proveedor_elegido") or f.get("proveedor_sugerido") or "Sin proveedor").strip()
            r = resumen.setdefault(prov, {"articulos": 0, "unidades": 0.0, "importe": 0.0, "adv": 0, "sp": 0})
            r["articulos"] += 1
            r["unidades"]  += f.get("cantidad_final") or 0
            if f["estado"] != "Sin precio":
                r["importe"] += f.get("precio_total") or 0
            if f["estado"] in ("Revisar duplicado", "Precio inactivo"):
                r["adv"] += 1
            if f["estado"] == "Sin precio":
                r["sp"] += 1

        if resumen:
            res_df = pd.DataFrame([
                {
                    "Proveedor": prov,
                    "Artículos": d["articulos"],
                    "Unidades": d["unidades"],
                    "Importe estimado": d["importe"],
                    "Con advertencias": d["adv"],
                    "Sin precio": d["sp"],
                }
                for prov, d in sorted(resumen.items())
            ])
            st.dataframe(res_df, use_container_width=True, hide_index=True)

        # ── BLOQUE 4: Exportación ─────────────────────────────────────────
        st.markdown("---")
        st.subheader("5 · Exportar órdenes de compra")

        filas_sel = [f for f in filas_vis if f.get("seleccionar")]
        if not filas_sel:
            st.info("Seleccioná al menos un artículo en la tabla para poder exportar.")
        else:
            exp_col1, exp_col2 = st.columns(2)

            with exp_col1:
                st.markdown("**Descargar por proveedor**")
                provs_sel = sorted(set(
                    (f.get("proveedor_elegido") or f.get("proveedor_sugerido") or "Sin proveedor")
                    for f in filas_sel
                ))
                prov_exp = st.selectbox("Proveedor", provs_sel, key="oc_exp_prov")
                if st.button("📥 Descargar OC de este proveedor", key="btn_oc_dl_uno"):
                    try:
                        data = exportar_un_proveedor(filas_sel, prov_exp)
                        st.download_button(
                            label=f"⬇️ {nombre_archivo_oc(prov_exp)}",
                            data=data,
                            file_name=nombre_archivo_oc(prov_exp),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_uno",
                        )
                    except Exception as e:
                        mostrar_error(f"Error al exportar: {e}")

            with exp_col2:
                st.markdown("**Descargar todas las órdenes**")
                st.caption("Un archivo Excel con una hoja por proveedor + hoja de resumen.")
                if st.button("📦 Descargar todas las OC", key="btn_oc_dl_all"):
                    try:
                        data_all = exportar_todas(filas_sel)
                        fecha_str = datetime.date.today().strftime("%Y-%m-%d")
                        st.download_button(
                            label=f"⬇️ OC_TODAS_{fecha_str}.xlsx",
                            data=data_all,
                            file_name=f"OC_TODAS_{fecha_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_all",
                        )
                    except Exception as e:
                        mostrar_error(f"Error al exportar: {e}")


st.markdown("---")
st.caption("Brufau Sanitarios · Herramientas internas de Compras")
