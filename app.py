import streamlit as st
import pandas as pd
import plotly.express as px

# Configuración inicial de la página
st.set_page_config(
    page_title="Entre Amigos Capital - Dashboard",
    page_icon="🤝",
    layout="wide"
)

EXCEL_FILE = "proyecto microcréditos copia 3.xlsx.xlsx"

@st.cache_data
def load_data():
    """Carga todas las hojas del archivo de Excel."""
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        
        df_clientes = pd.read_excel(xls, sheet_name='Clientes')
        df_creditos = pd.read_excel(xls, sheet_name='Creditos')
        df_pagos = pd.read_excel(xls, sheet_name='Pagos')
        
        df_estado_cartera = pd.read_excel(xls, sheet_name='Estado_Cartera') if 'Estado_Cartera' in xls.sheet_names else None
        df_resumen = pd.read_excel(xls, sheet_name='Resumen_Cartera') if 'Resumen_Cartera' in xls.sheet_names else None

        return df_clientes, df_creditos, df_pagos, df_estado_cartera, df_resumen
    except Exception as e:
        st.error(f"Error al cargar el archivo de Excel: {e}")
        return None, None, None, None, None

df_clientes, df_creditos, df_pagos, df_estado_cartera, df_resumen = load_data()

# ---------------------------------------------------------
# MENÚ DE NAVEGACIÓN LATERAL
# ---------------------------------------------------------
st.sidebar.title("🤝 Entre Amigos Capital")
st.sidebar.caption("Créditos justos sobre la base de la confianza")

opcion_menu = st.sidebar.radio(
    "Selecciona una sección:",
    ["📊 Dashboard General", "👤 Ficha por Cliente", "🧮 Simulador de Créditos", "ℹ️ Sobre Nosotros & Políticas"]
)

if df_creditos is not None:

    # =========================================================
    # 1. DASHBOARD GENERAL
    # =========================================================
    if opcion_menu == "📊 Dashboard General":
        st.title("📊 Control General de Cartera")
        st.markdown("---")

        val_prestado = 23900000
        val_pagado = 3033328
        val_cap_pendiente = 20866672
        val_int_pendiente = 775000
        val_deuda_total = 21641672
        val_deuda_vencida = 1827500
        creditos_activos = 14
        creditos_mora = 10
        creditos_aldia = 4

        if df_resumen is not None:
            try:
                dict_res = dict(zip(df_resumen['Indicador'], df_resumen['Resultado']))
                val_prestado = dict_res.get('Capital total prestado', val_prestado)
                val_pagado = dict_res.get('Capital pagado', val_pagado)
                val_cap_pendiente = dict_res.get('Capital pendiente', val_cap_pendiente)
                val_int_pendiente = dict_res.get('Intereses pendientes', val_int_pendiente)
                val_deuda_total = dict_res.get('Deuda total pendiente', val_deuda_total)
                val_deuda_vencida = dict_res.get('Deuda vencida', val_deuda_vencida)
                creditos_activos = dict_res.get('Créditos activos', creditos_activos)
                creditos_mora = dict_res.get('Créditos en mora', creditos_mora)
                creditos_aldia = dict_res.get('Créditos al día', creditos_aldia)
            except:
                pass

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Capital Total Prestado", f"${val_prestado:,.0f} COP")
        c2.metric("Capital Pagado", f"${val_pagado:,.0f} COP")
        c3.metric("Capital Pendiente", f"${val_cap_pendiente:,.0f} COP")
        c4.metric("Deuda Total Pendiente", f"${val_deuda_total:,.0f} COP")

        st.markdown("<br>", unsafe_allow_html=True)

        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Intereses Pendientes", f"${val_int_pendiente:,.0f} COP")
        c6.metric("Deuda Vencida (Mora)", f"${val_deuda_vencida:,.0f} COP", delta=f"-{creditos_mora} créditos", delta_color="inverse")
        c7.metric("Créditos Al Día", f"{creditos_aldia}", delta="Puntuales")
        c8.metric("Total Créditos Activos", f"{creditos_activos}")

        st.markdown("---")

        col_left, col_right = st.columns([1, 1])

        with col_left:
            st.subheader("Estado de Créditos Activos")
            df_estado = pd.DataFrame({
                "Estado": ["Al Día", "En Mora"],
                "Cantidad": [creditos_aldia, creditos_mora]
            })
            fig_pie = px.pie(df_estado, names="Estado", values="Cantidad", hole=0.4,
                             color="Estado",
                             color_discrete_map={"Al Día": "#2ecc71", "En Mora": "#e74c3c"})
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_right:
            st.subheader("Detalle de Créditos Registrados")
            st.dataframe(df_creditos, use_container_width=True)

    # =========================================================
    # 2. FICHA POR CLIENTE (CON ESTADO DE MORA Y CAPITAL/INTERÉS)
    # =========================================================
    elif opcion_menu == "👤 Ficha por Cliente":
        st.title("👤 Ficha de Cliente e Historial de Deuda")
        st.markdown("---")

        lista_clientes = df_clientes['nombre'].dropna().unique()
        cliente_sel = st.selectbox("Selecciona un cliente:", lista_clientes)

        if cliente_sel:
            info_cliente = df_clientes[df_clientes['nombre'] == cliente_sel].iloc[0]
            cliente_id = info_cliente['cliente_id']

            creditos_cliente = df_creditos[df_creditos['cliente_id'] == cliente_id]
            pagos_cliente = df_pagos[df_pagos['cliente_id'] == cliente_id]
            
            # Obtener datos financieros desde Estado_Cartera
            if df_estado_cartera is not None:
                cartera_cliente = df_estado_cartera[df_estado_cartera['cliente_id'] == cliente_id]
            else:
                cartera_cliente = pd.DataFrame()

            cap_pendiente = cartera_cliente['capital_pendiente'].sum() if not cartera_cliente.empty else 0
            int_pendiente = cartera_cliente['interes_pendiente'].sum() if not cartera_cliente.empty else 0
            deuda_vencida = cartera_cliente['deuda_vencida'].sum() if not cartera_cliente.empty else 0
            deuda_total = cartera_cliente['deuda_total_pendiente'].sum() if not cartera_cliente.empty else 0

            # Determinar si tiene créditos en mora
            tiene_mora = any(cartera_cliente['estado'].astype(str).str.contains('mora', case=False, na=False)) if not cartera_cliente.empty else False

            # Muestra de Estado
            if tiene_mora or deuda_vencida > 0:
                st.error(f"⚠️ **Cliente con Cuotas/Intereses Vencidos:** Presenta un valor en mora de **${deuda_vencida:,.0f} COP**.")
            elif deuda_total == 0:
                st.success("🟢 **Paz y Salvo:** El cliente no presenta saldos pendientes.")
            else:
                st.info("🟢 **Al Día:** El cliente cuenta con sus cuotas e intereses al día.")

            # Tarjetas de resumen del cliente
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Capital Pendiente", f"${cap_pendiente:,.0f} COP")
            m2.metric("Intereses Pendientes", f"${int_pendiente:,.0f} COP")
            m3.metric("Deuda Vencida (Mora)", f"${deuda_vencida:,.0f} COP")
            m4.metric("Deuda Total Pendiente", f"${deuda_total:,.0f} COP")

            st.markdown("---")

            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("Información Personal")
                st.write(f"**ID Cliente:** {cliente_id}")
                st.write(f"**Nombre:** {info_cliente.get('nombre', 'N/A')}")
                st.write(f"**Teléfono:** {info_cliente.get('telefono', 'N/A')}")
            with col_b:
                st.subheader("Relación y Registro")
                st.write(f"**Fecha Registro:** {info_cliente.get('fecha_registro', 'N/A')}")
                st.write(f"**Estado Cliente:** {info_cliente.get('estado_cliente', 'N/A')}")
                st.write(f"**Parentesco:** {info_cliente.get('parentesco', 'N/A')}")

            st.markdown("---")
            
            st.subheader("Créditos de este Cliente")
            if not creditos_cliente.empty:
                st.dataframe(creditos_cliente, use_container_width=True)
            else:
                st.info("Este cliente no tiene créditos registrados.")

            st.subheader("Historial de Pagos de este Cliente")
            if not pagos_cliente.empty:
                cols_validas = [c for c in pagos_cliente.columns if not str(c).startswith('Unnamed')]
                st.dataframe(pagos_cliente[cols_validas], use_container_width=True)
            else:
                st.info("Este cliente no registra pagos en el sistema.")

    # =========================================================
    # 3. SIMULADOR DE CRÉDITOS
    # =========================================================
    elif opcion_menu == "🧮 Simulador de Créditos":
        st.title("🧮 Simulador de Créditos - Entre Amigos Capital")
        st.markdown("Calcule la cuota estimada a una **tasa fija del 3% mensual**.")
        st.markdown("---")

        col_sim1, col_sim2 = st.columns(2)

        with col_sim1:
            monto_sim = st.number_input("Monto a solicitar (COP):", min_value=100000, max_value=5000000, value=500000, step=50000, format="%d")
            modalidad_sim = st.selectbox("Modalidad de pago:", ["Intereses periódicos + Capital al final", "Cuotas fijas (Capital + Interés)"])
            plazo_sim = st.slider("Plazo en meses:", min_value=1, max_value=24, value=6)

        tasa_mensual = 0.03

        with col_sim2:
            st.subheader("📋 Resumen de la Simulación")
            if modalidad_sim == "Intereses periódicos + Capital al final":
                interes_mensual = monto_sim * tasa_mensual
                total_intereses = interes_mensual * plazo_sim
                total_pagar = monto_sim + total_intereses

                st.write(f"• **Pago mensual de intereses:** ${interes_mensual:,.0f} COP")
                st.write(f"• **Pago final de capital (Mes {plazo_sim}):** ${monto_sim:,.0f} COP")
                st.write(f"• **Total de intereses:** ${total_intereses:,.0f} COP")
                st.metric("Total General a Cancelar", f"${total_pagar:,.0f} COP")
            else:
                cuota_mensual = (monto_sim * tasa_mensual) / (1 - (1 + tasa_mensual)**(-plazo_sim))
                total_pagar = cuota_mensual * plazo_sim
                total_intereses = total_pagar - monto_sim

                st.write(f"• **Cuota fija mensual:** ${cuota_mensual:,.0f} COP")
                st.write(f"• **Total intereses:** ${total_intereses:,.0f} COP")
                st.metric("Total General a Cancelar", f"${total_pagar:,.0f} COP")

    # =========================================================
    # 4. SOBRE NOSOTROS & POLÍTICAS
    # =========================================================
    elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
        st.title("🤝 Sobre Entre Amigos Capital")
        st.markdown("---")

        st.subheader("💡 Nuestra Propuesta de Valor")
        st.write("Acceso a microcréditos ágiles para familiares y amigos a **tasas justas (3% mensual)**, sin sanciones ocultas y bajo el principio de buena fe.")

        st.subheader("📜 Políticas de Crédito")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("""
            **Montos y Plazos**
            * Desde $100.000 hasta $5.000.000 COP.
            * Modalidades: Interés mensual + Capital final / Cuota fija.
            """)
        with col_p2:
            st.markdown("""
            **Pagos y Mora**
            * Medios: Transferencia (Nequi/Bancolombia) o Efectivo.
            * Cero sanciones financieras por mora.
            """)
