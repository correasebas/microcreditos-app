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
# INICIALIZACIÓN DE SESSION STATE (PREVIENE KEYERROR)
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
        
        # Cargar pestañas
        st.session_state['df_clientes'] = pd.read_excel(xls, 'Clientes') if 'Clientes' in xls.sheet_names else pd.DataFrame()
        st.session_state['df_creditos'] = pd.read_excel(xls, 'Creditos') if 'Creditos' in xls.sheet_names else pd.DataFrame()
        st.session_state['df_estado_cartera'] = pd.read_excel(xls, 'Estado_Cartera') if 'Estado_Cartera' in xls.sheet_names else pd.DataFrame()
        st.session_state['df_historico_pagos'] = pd.read_excel(xls, 'Historico_Pagos') if 'Historico_Pagos' in xls.sheet_names else pd.DataFrame()
        
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

# Recuperar DataFrames desde session_state de forma segura
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
        # Métricas principales
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
# 4. ALERTAS DE MORA Y GESTIÓN DE COBROS (VERSIÓN BASE)
# =========================================================
elif opcion_menu == "🔔 Alertas & Cobros":
    st.title("🔔 Alertas de Mora y Gestión de Cobros")
    st.markdown("Seguimiento de días de retraso en los pagos y generación de avisos vía WhatsApp sin cobro de intereses por mora.")
    st.markdown("---")

    if df_estado_cartera.empty:
        st.info("No hay información disponible en la tabla de estado de cartera. Carga tu archivo Excel en la barra lateral.")
    else:
        # Unir Estado_Cartera con Clientes y Créditos
        df_mora = df_estado_cartera.merge(
            df_clientes[['cliente_id', 'nombre', 'telefono']], on='cliente_id', how='left'
        )
        
        if not df_creditos.empty:
            cols_cred_merge = [c for c in ['credito_id', 'fecha_inicio', 'dia_pago_mes'] if c in df_creditos.columns]
            if len(cols_cred_merge) > 1:
                df_mora = df_mora.merge(df_creditos[cols_cred_merge], on='credito_id', how='left')

        # Filtrar créditos con deuda pendiente
        df_mora = df_mora[df_mora['deuda_total_pendiente'] > 0].copy()

        if df_mora.empty:
            st.success("🎉 ¡Excelente! Toda la cartera está al día y en paz y salvo.")
        else:
            def obtener_dias_mora(row):
                if 'dias_mora' in row and pd.notnull(row['dias_mora']):
                    try:
                        return int(float(row['dias_mora']))
                    except:
                        pass
                
                estado_str = str(row.get('estado', '')).lower()
                if 'mora' in estado_str or 'vencid' in estado_str or 'atras' in estado_str:
                    return 30
                
                return 0

            df_mora['dias_mora_calculados'] = df_mora.apply(obtener_dias_mora, axis=1)

            def clasificar_riesgo(dias):
                if dias <= 0:
                    return "🟢 Al Día / Preventivo"
                elif dias <= 30:
                    return "🟡 Retraso Leve (1-30 días)"
                elif dias <= 60:
                    return "🟠 Retraso Medio (31-60 días)"
                else:
                    return "🔴 Retraso Alto (>60 días)"

            df_mora['nivel_riesgo'] = df_mora['dias_mora_calculados'].apply(clasificar_riesgo)

            tot_clientes_mora = len(df_mora[df_mora['dias_mora_calculados'] > 0])
            monto_pendiente = df_mora['deuda_total_pendiente'].sum()
            max_dias = df_mora['dias_mora_calculados'].max()

            c_a1, c_a2, c_a3, c_a4 = st.columns(4)
            c_a1.metric("Clientes Activos", f"{len(df_mora)}")
            c_a2.metric("Clientes con Retraso", f"{tot_clientes_mora}", delta=f"{tot_clientes_mora} alertas", delta_color="inverse")
            c_a3.metric("Deuda Pendiente Total", f"${monto_pendiente:,.0f} COP")
            c_a4.metric("Máximo Días de Retraso", f"{max_dias} días")

            st.markdown("---")

            filtro_nivel = st.multiselect(
                "Filtrar por Nivel de Retraso:",
                options=["🟢 Al Día / Preventivo", "🟡 Retraso Leve (1-30 días)", "🟠 Retraso Medio (31-60 días)", "🔴 Retraso Alto (>60 días)"],
                default=[]
            )

            if filtro_nivel:
                df_filtrado = df_mora[df_mora['nivel_riesgo'].isin(filtro_nivel)]
            else:
                df_filtrado = df_mora.copy()

            st.subheader("📋 Detalle de Clientes y Gestión de Cobro")

            if df_filtrado.empty:
                st.info("No hay registros que coincidan con los filtros seleccionados.")
            else:
                for _, row in df_filtrado.iterrows():
                    dias = int(row['dias_mora_calculados'])
                    with st.expander(f"{row['nivel_riesgo']} | {row['nombre']} - Crédito `{row['credito_id']}` ({dias} días de retraso)"):
                        col_m1, col_m2 = st.columns([1, 1])

                        with col_m1:
                            st.write(f"**Cliente:** {row['nombre']}")
                            st.write(f"**Teléfono:** {row.get('telefono', 'N/A')}")
                            st.write(f"**Estado registrado:** {row.get('estado', 'N/A')}")
                            st.write(f"**Días de Retraso:** {dias} días")
                            st.write(f"**Capital Pendiente:** ${row.get('capital_pendiente', 0):,.0f} COP")
                            st.write(f"**Intereses Pendientes:** ${row.get('interes_pendiente', 0):,.0f} COP")
                            st.write(f"**Deuda Total Pendiente:** ${row['deuda_total_pendiente']:,.0f} COP")

                        with col_m2:
                            st.write("📲 **Plantillas de Recordatorio de Pago:**")

                            tipo_msg = st.radio(
                                "Selecciona el tipo de aviso:",
                                ["Preventivo", "Recordatorio de Retraso", "Aviso de Urgencia"],
                                key=f"radio_{row['credito_id']}"
                            )

                            if tipo_msg == "Preventivo":
                                texto_base = (
                                    f"Hola *{row['nombre']}* 👋,\n\n"
                                    f"Te saludamos de *Entre Amigos Capital* 🤝.\n"
                                    f"Te recordamos amablemente que tu cuota del crédito *{row['credito_id']}* está próxima a vencer por un saldo pendiente de *${row['deuda_total_pendiente']:,.0f} COP*.\n\n"
                                    f"Quedamos atentos a la confirmación de tu pago. ¡Que tengas un excelente día!"
                                )
                            elif tipo_msg == "Recordatorio de Retraso":
                                texto_base = (
                                    f"Hola *{row['nombre']}* 👋,\n\n"
                                    f"Te escribimos de *Entre Amigos Capital* 🤝 para recordarte que registras un saldo pendiente en tu crédito *{row['credito_id']}*.\n\n"
                                    f"🔹 *Saldo Pendiente:* ${row['deuda_total_pendiente']:,.0f} COP\n\n"
                                    f"Te invitamos a ponerte al día para mantener tus condiciones activas. Por favor confírmanos cuándo podrías realizar el abono. ¡Gracias!"
                                )
                            else:
                                texto_base = (
                                    f"Hola *{row['nombre']}* 👋,\n\n"
                                    f"Nos comunicamos de *Entre Amigos Capital* 🤝 referente a tu crédito *{row['credito_id']}*, el cual presenta un saldo pendiente de *${row['deuda_total_pendiente']:,.0f} COP*.\n\n"
                                    f"Te pedimos ponerte en contacto con nosotros hoy mismo para definir una fecha de pago. ¡Agradecemos tu atención!"
                                )

                            msg_editado = st.text_area("Editar mensaje antes de enviar:", value=texto_base, height=130, key=f"txt_{row['credito_id']}")
                            
                            num_tel = "".join(filter(str.isdigit, str(row.get('telefono', ''))))
                            if len(num_tel) == 10 and not num_tel.startswith("57"):
                                num_tel = "57" + num_tel

                            msg_enc = urllib.parse.quote(msg_editado)
                            
                            st.link_button("💬 Enviar Recordatorio por WhatsApp", f"https://wa.me/{num_tel}?text={msg_enc}", use_container_width=True)


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
