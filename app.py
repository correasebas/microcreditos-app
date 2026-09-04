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
        
        # Cargar hoja de resumen/indicadores
        sheet_resumen = [s for s in xls.sheet_names if 'resumen' in s.lower() or 'indicador' in s.lower() or 'kpi' in s.lower()]
        if sheet_resumen:
            df_resumen = pd.read_excel(xls, sheet_name=sheet_resumen[0])
        else:
            df_resumen = None

        return df_clientes, df_creditos, df_pagos, df_resumen
    except Exception as e:
        st.error(f"Error al cargar el archivo de Excel: {e}")
        return None, None, None, None

df_clientes, df_creditos, df_pagos, df_resumen = load_data()

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

        # Valores predeterminados
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
                dict_resumen = pd.Series(df_resumen.iloc[:, 1].values, index=df_resumen.iloc[:, 0].values).to_dict()
                for k, v in dict_resumen.items():
                    k_str = str(k).lower()
                    if 'prestado' in k_str: val_prestado = v
                    elif 'capital pagado' in k_str: val_pagado = v
                    elif 'capital pendiente' in k_str: val_cap_pendiente = v
                    elif 'intereses pendientes' in k_str: val_int_pendiente = v
                    elif 'deuda total' in k_str: val_deuda_total = v
                    elif 'deuda vencida' in k_str: val_deuda_vencida = v
                    elif 'activos' in k_str: creditos_activos = v
                    elif 'mora' in k_str: creditos_mora = v
                    elif 'día' in k_str or 'dia' in k_str: creditos_aldia = v
            except:
                pass

        # Tarjetas principales de KPI
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
    # 2. FICHA POR CLIENTE (FILTRADA CORRECTAMENTE)
    # =========================================================
    elif opcion_menu == "👤 Ficha por Cliente":
        st.title("👤 Ficha de Cliente e Historial de Confianza")
        st.markdown("---")

        # Detectar columnas clave de manera flexible
        col_nom_cli = next((c for c in df_clientes.columns if 'nombre' in str(c).lower()), df_clientes.columns[0])
        col_id_cli = next((c for c in df_clientes.columns if 'cliente' in str(c).lower() or 'id' in str(c).lower()), df_clientes.columns[0])
        
        lista_clientes = df_clientes[col_nom_cli].dropna().unique()
        cliente_sel = st.selectbox("Selecciona un cliente:", lista_clientes)

        if cliente_sel:
            # Fila del cliente seleccionado
            info_cliente = df_clientes[df_clientes[col_nom_cli] == cliente_sel].iloc[0]
            id_cliente = info_cliente[col_id_cli]

            # Buscar la columna de cliente/crédito en las demás tablas
            col_id_credito_cli = next((c for c in df_creditos.columns if 'cliente' in str(c).lower() or 'id' in str(c).lower()), None)
            col_id_pago_cli = next((c for c in df_pagos.columns if 'cliente' in str(c).lower() or 'id' in str(c).lower()), None)

            # Filtrar créditos del cliente
            if col_id_credito_cli:
                creditos_cliente = df_creditos[df_creditos[col_id_credito_cli] == id_cliente]
                ids_creditos_cliente = creditos_cliente.iloc[:, 0].unique() if not creditos_cliente.empty else []
            else:
                creditos_cliente = pd.DataFrame()
                ids_creditos_cliente = []

            # Filtrar pagos del cliente (por id_cliente o por los id_credito del cliente)
            if col_id_pago_cli:
                pagos_cliente = df_pagos[df_pagos[col_id_pago_cli] == id_cliente]
            else:
                # Intentar filtrar pagos usando los IDs de crédito
                col_credito_en_pagos = next((c for c in df_pagos.columns if 'credito' in str(c).lower() or 'id' in str(c).lower()), None)
                if col_credito_en_pagos:
                    pagos_cliente = df_pagos[df_pagos[col_credito_en_pagos].isin(ids_creditos_cliente)]
                else:
                    pagos_cliente = pd.DataFrame()

            st.subheader("Información Registrada")
            col_a, col_b = st.columns(2)
            with col_a:
                for c in df_clientes.columns[:len(df_clientes.columns)//2]:
                    st.write(f"**{c.capitalize()}:** {info_cliente.get(c, 'N/A')}")
            with col_b:
                for c in df_clientes.columns[len(df_clientes.columns)//2:]:
                    st.write(f"**{c.capitalize()}:** {info_cliente.get(c, 'N/A')}")

            st.markdown("---")
            
            st.subheader("Créditos de este Cliente")
            if not creditos_cliente.empty:
                st.dataframe(creditos_cliente, use_container_width=True)
            else:
                st.info("Este cliente no tiene créditos registrados.")

            st.subheader("Historial de Pagos de este Cliente")
            if not pagos_cliente.empty:
                st.dataframe(pagos_cliente, use_container_width=True)
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
