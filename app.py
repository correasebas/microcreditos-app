import streamlit as st
import pandas as pd
from datetime import datetime
import urllib.parse

# ---------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Entre Amigos Capital - Gestión de Créditos",
    page_icon="🤝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# INICIALIZACIÓN DE SESSION STATE
# ---------------------------------------------------------
if 'df_clientes' not in st.session_state:
    st.session_state['df_clientes'] = pd.DataFrame()
if 'df_creditos' not in st.session_state:
    st.session_state['df_creditos'] = pd.DataFrame()
if 'df_estado_cartera' not in st.session_state:
    st.session_state['df_estado_cartera'] = pd.DataFrame()
if 'df_historico_pagos' not in st.session_state:
    st.session_state['df_historico_pagos'] = pd.DataFrame()

# ---------------------------------------------------------
# CARGA DE DATOS (EXCEL)
# ---------------------------------------------------------
st.sidebar.title("🤝 Entre Amigos Capital")
st.sidebar.caption("Créditos justos sobre la base de la confianza")
st.sidebar.markdown("---")

st.sidebar.subheader("📁 Base de Datos Excel")
uploaded_file = st.sidebar.file_uploader("Carga tu archivo de Excel actualizado:", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        st.session_state['df_clientes'] = pd.read_excel(xls, 'Clientes') if 'Clientes' in xls.sheet_names else pd.DataFrame()
        st.session_state['df_creditos'] = pd.read_excel(xls, 'Creditos') if 'Creditos' in xls.sheet_names else pd.DataFrame()
        st.session_state['df_estado_cartera'] = pd.read_excel(xls, 'Estado_Cartera') if 'Estado_Cartera' in xls.sheet_names else pd.DataFrame()
        
        # Soportar tanto 'Historico_Pagos' como 'Pagos'
        if 'Historico_Pagos' in xls.sheet_names:
            st.session_state['df_historico_pagos'] = pd.read_excel(xls, 'Historico_Pagos')
        elif 'Pagos' in xls.sheet_names:
            st.session_state['df_historico_pagos'] = pd.read_excel(xls, 'Pagos')
        else:
            st.session_state['df_historico_pagos'] = pd.DataFrame()
        
        st.sidebar.success("✅ Archivo cargado correctamente")
    except Exception as e:
        st.sidebar.error(f"Error al cargar el archivo Excel: {e}")

# NAVEGACIÓN PRINCIPAL
st.sidebar.markdown("---")
opcion_menu = st.sidebar.radio(
    "Selecciona una sección:",
    [
        "📊 Dashboard General",
        "👤 Ficha por Cliente",
        "📝 Registrar Pago",
        "🔔 Alertas & Cobros",
        "🧮 Simulador de Créditos",
        "ℹ️ Sobre Nosotros & Políticas"
    ]
)

df_clientes = st.session_state['df_clientes']
df_creditos = st.session_state['df_creditos']
df_estado_cartera = st.session_state['df_estado_cartera']
df_historico_pagos = st.session_state['df_historico_pagos']


# =========================================================
# 1. DASHBOARD GENERAL
# =========================================================
if opcion_menu == "📊 Dashboard General":
    st.title("📊 Dashboard General - Entre Amigos Capital")
    st.markdown("Vista general del estado de la cartera de créditos y métricas clave.")
    st.markdown("---")

    if df_estado_cartera.empty:
        st.warning("⚠️ Por favor carga el archivo Excel en el menú lateral para visualizar la información.")
    else:
        col1, col2, col3, col4 = st.columns(4)
        
        total_clientes = len(df_clientes) if not df_clientes.empty else 0
        total_creditos = len(df_creditos) if not df_creditos.empty else 0
        
        deuda_total = df_estado_cartera['deuda_total_pendiente'].sum() if 'deuda_total_pendiente' in df_estado_cartera.columns else 0
        capital_pendiente = df_estado_cartera['capital_pendiente'].sum() if 'capital_pendiente' in df_estado_cartera.columns else 0

        col1.metric("Clientes Registrados", total_clientes)
        col2.metric("Créditos Registrados", total_creditos)
        col3.metric("Capital Pendiente", f"${capital_pendiente:,.0f} COP")
        col4.metric("Deuda Total Pendiente", f"${deuda_total:,.0f} COP")

        st.markdown("---")
        st.subheader("📋 Estado de Cartera Actual")
        st.dataframe(df_estado_cartera, use_container_width=True)


# =========================================================
# 2. FICHA POR CLIENTE
# =========================================================
elif opcion_menu == "👤 Ficha por Cliente":
    st.title("👤 Ficha por Cliente")
    st.markdown("Consulta la información detallada de cada cliente y su historial de pagos.")
    st.markdown("---")

    if df_clientes.empty:
        st.warning("⚠️ Por favor carga el archivo Excel en el menú lateral.")
    else:
        cliente_sel = st.selectbox("Selecciona un Cliente:", df_clientes['nombre'].unique())
        
        if cliente_sel:
            datos_cliente = df_clientes[df_clientes['nombre'] == cliente_sel].iloc[0]
            c_id = datos_cliente['cliente_id']

            col_c1, col_c2 = st.columns(2)
            with col_c1:
                st.write(f"**ID Cliente:** `{datos_cliente['cliente_id']}`")
                st.write(f"**Nombre:** {datos_cliente['nombre']}")
                st.write(f"**Cédula:** {datos_cliente.get('cedula', 'N/A')}")
            with col_c2:
                st.write(f"**Teléfono:** {datos_cliente.get('telefono', 'N/A')}")
                st.write(f"**Ocupación:** {datos_cliente.get('ocupacion', 'N/A')}")

            st.markdown("---")
            st.subheader("Créditos Asociados")
            creditos_cli = df_creditos[df_creditos['cliente_id'] == c_id] if not df_creditos.empty else pd.DataFrame()
            st.dataframe(creditos_cli, use_container_width=True)

            st.subheader("Historial de Pagos")
            pagos_cli = df_historico_pagos[df_historico_pagos['cliente_id'] == c_id] if not df_historico_pagos.empty else pd.DataFrame()
            st.dataframe(pagos_cli, use_container_width=True)


# =========================================================
# 3. REGISTRAR PAGO
# =========================================================
elif opcion_menu == "📝 Registrar Pago":
    st.title("📝 Registrar Pago y Generar Comprobante")
    st.markdown("Registra un abono y genera la notificación para WhatsApp.")
    st.markdown("---")

    if df_clientes.empty or df_creditos.empty:
        st.warning("⚠️ Carga el archivo Excel para poder registrar pagos.")
    else:
        with st.form("form_pago"):
            cliente_sel = st.selectbox("Cliente:", df_clientes['nombre'].unique())
            monto_pago = st.number_input("Monto Pagado (COP):", min_value=1000, step=5000, value=50000)
            fecha_pago = st.date_input("Fecha del Pago:", datetime.today())
            notas = st.text_input("Observaciones / Notas:", value="Abono a cuota mensual")
            
            submit_pago = st.form_submit_button("Confirmar y Registrar Pago")

        if submit_pago:
            datos_cli = df_clientes[df_clientes['nombre'] == cliente_sel].iloc[0]
            tel = str(datos_cli.get('telefono', ''))
            num_tel = "".join(filter(str.isdigit, tel))
            if len(num_tel) == 10 and not num_tel.startswith("57"):
                num_tel = "57" + num_tel

            msg_pago = (
                f"Hola *{cliente_sel}* 👋,\n\n"
                f"Hemos recibido exitosamente tu pago en *Entre Amigos Capital* 🤝.\n\n"
                f"💵 *Monto:* ${monto_pago:,.0f} COP\n"
                f"📅 *Fecha:* {fecha_pago.strftime('%d/%m/%Y')}\n"
                f"📌 *Concepto:* {notas}\n\n"
                f"¡Gracias por tu puntualidad!"
            )
            
            st.success("✅ Pago registrado correctamente.")
            st.markdown("### 💬 Generar Comprobante por WhatsApp")
            msg_enc = urllib.parse.quote(msg_pago)
            st.link_button("📲 Enviar Recibo por WhatsApp", f"https://wa.me/{num_tel}?text={msg_enc}", use_container_width=True)


# =========================================================
# 4. ALERTAS DE MORA Y GESTIÓN DE COBROS (2 OPCIONES)
# =========================================================
elif opcion_menu == "🔔 Alertas & Cobros":
    st.title("🔔 Alertas & Cobros")
    st.markdown("Gestión simplificada de cartera: clientes al día y clientes en mora.")
    st.markdown("---")

    if df_estado_cartera.empty:
        st.info("No hay información disponible. Carga tu archivo Excel en la barra lateral.")
    else:
        df_mora = df_estado_cartera.merge(
            df_clientes[['cliente_id', 'nombre', 'telefono']], on='cliente_id', how='left'
        )

        def clasificar_estado_simple(row):
            est = str(row.get('estado', '')).strip().lower()
            if 'mora' in est or 'vencid' in est or 'atras' in est:
                return "🔴 En Mora"
            return "🟢 Al Día"

        df_mora['clasificacion_simple'] = df_mora.apply(clasificar_estado_simple, axis=1)

        tot_mora = len(df_mora[df_mora['clasificacion_simple'] == "🔴 En Mora"])
        tot_aldia = len(df_mora[df_mora['clasificacion_simple'] == "🟢 Al Día"])
        deuda_mora = df_mora[df_mora['clasificacion_simple'] == "🔴 En Mora"]['deuda_total_pendiente'].sum()

        c_a1, c_a2, c_a3 = st.columns(3)
        c_a1.metric("Clientes Al Día", f"{tot_aldia}")
        c_a2.metric("Clientes En Mora", f"{tot_mora}", delta=f"{tot_mora} casos", delta_color="inverse")
        c_a3.metric("Monto Total Pendiente en Mora", f"${deuda_mora:,.0f} COP")

        st.markdown("---")

        filtro_estado = st.radio(
            "Selecciona el grupo a consultar:",
            options=["🔴 En Mora", "🟢 Al Día"],
            horizontal=True
        )

        df_filtrado = df_mora[df_mora['clasificacion_simple'] == filtro_estado]

        st.subheader(f"📋 Registro de Clientes ({filtro_estado})")

        if df_filtrado.empty:
            st.info(f"No hay registros en la categoría {filtro_estado}.")
        else:
            for _, row in df_filtrado.iterrows():
                deuda_val = row.get('deuda_total_pendiente', 0)
                deuda_venc = row.get('deuda_vencida', 0)
                nombre_cliente = row.get('nombre', 'Cliente')

                with st.expander(f"{row['clasificacion_simple']} | {nombre_cliente} - Crédito `{row['credito_id']}` | Pendiente: ${deuda_val:,.0f} COP"):
                    col_m1, col_m2 = st.columns([1, 1])

                    with col_m1:
                        st.write(f"**Cliente:** {nombre_cliente}")
                        st.write(f"**Teléfono:** {row.get('telefono', 'N/A')}")
                        st.write(f"**Estado en Excel:** `{row.get('estado', 'N/A')}`")
                        st.write(f"**Capital Pendiente:** ${row.get('capital_pendiente', 0):,.0f} COP")
                        st.write(f"**Deuda Vencida:** ${deuda_venc:,.0f} COP")
                        st.write(f"**Deuda Total Pendiente:** ${deuda_val:,.0f} COP")

                    with col_m2:
                        st.write("📲 **Aviso Preventivo WhatsApp:**")

                        if filtro_estado == "🔴 En Mora":
                            texto_base = (
                                f"Hola *{nombre_cliente}* 👋,\n\n"
                                f"Te saludamos amablemente de *Entre Amigos Capital* 🤝.\n\n"
                                f"Te escribimos para informarte de manera preventiva que tu crédito *{row['credito_id']}* registra días de mora / atraso en su pago.\n\n"
                                f"📌 *Valor Pendiente:* ${deuda_val:,.0f} COP\n\n"
                                f"Te invitamos a ponernos al día o confirmarnos la fecha en que podrías efectuar tu abono. ¡Agradecemos tu atención!"
                            )
                        else:
                            texto_base = (
                                f"Hola *{nombre_cliente}* 👋,\n\n"
                                f"Te saludamos de *Entre Amigos Capital* 🤝.\n\n"
                                f"Te confirmamos que tu crédito *{row['credito_id']}* se encuentra al día. Saldo total pendiente: *${deuda_val:,.0f} COP*.\n\n"
                                f"¡Gracias por tu responsabilidad y compromiso!"
                            )

                        msg_editado = st.text_area("Editar mensaje antes de enviar:", value=texto_base, height=140, key=f"txt_{row['credito_id']}")
                        
                        num_tel = "".join(filter(str.isdigit, str(row.get('telefono', ''))))
                        if len(num_tel) == 10 and not num_tel.startswith("57"):
                            num_tel = "57" + num_tel

                        msg_enc = urllib.parse.quote(msg_editado)
                        
                        st.link_button("💬 Enviar Mensaje por WhatsApp", f"https://wa.me/{num_tel}?text={msg_enc}", use_container_width=True)


# =========================================================
# 5. SIMULADOR DE CRÉDITOS
# =========================================================
elif opcion_menu == "🧮 Simulador de Créditos":
    st.title("🧮 Simulador de Créditos")
    st.markdown("Calcula cuotas mensuales con interés simple.")
    st.markdown("---")

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        monto_sim = st.number_input("Monto a prestar (COP):", min_value=100000, max_value=50000000, value=1000000, step=100000)
        plazo_meses = st.slider("Plazo (meses):", min_value=1, max_value=36, value=6)
    with col_s2:
        tasa_interes = st.number_input("Tasa de interés mensual (%):", min_value=1.0, max_value=10.0, value=3.0, step=0.5)

    interes_mensual_cop = monto_sim * (tasa_interes / 100)
    capital_mensual_cop = monto_sim / plazo_meses
    cuota_total_mensual = capital_mensual_cop + interes_mensual_cop
    total_a_pagar = cuota_total_mensual * plazo_meses

    st.markdown("---")
    st.subheader("Resultados de la Simulación")
    res_col1, res_col2, res_col3 = st.columns(3)
    res_col1.metric("Cuota Mensual Estimada", f"${cuota_total_mensual:,.0f} COP")
    res_col2.metric("Interés Mensual", f"${interes_mensual_cop:,.0f} COP")
    res_col3.metric("Total a Pagar", f"${total_a_pagar:,.0f} COP")


# =========================================================
# 6. SOBRE NOSOTROS & POLÍTICAS
# =========================================================
elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
    st.title("ℹ️ Sobre Nosotros & Políticas")
    st.markdown("---")
    st.subheader("🤝 Entre Amigos Capital")
    st.write("**Misión:** Créditos justos sobre la base de la confianza.")
    st.write("**Política de Cobro:** No se cobran intereses de mora por retrasos. Se prioriza el diálogo directo y el acuerdo de pago mutuo.")
