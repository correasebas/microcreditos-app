import streamlit as st
import pandas as pd
import plotly.express as px

# Configuración inicial de la página
st.set_page_config(
    page_title="Entre Amigos Capital - Dashboard",
    page_icon="🤝",
    layout="wide"
)

# Nombre del archivo de Excel subido a GitHub
EXCEL_FILE = "proyecto microcréditos copia 3.xlsx.xlsx"

@st.cache_data
def load_data():
    """Carga y procesa los datos del archivo de Excel."""
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        
        # Cargar hojas
        df_clientes = pd.read_excel(xls, sheet_name='Clientes')
        df_creditos = pd.read_excel(xls, sheet_name='Creditos')
        df_pagos = pd.read_excel(xls, sheet_name='Pagos')
        
        # Limpieza básica
        df_clientes.columns = df_clientes.columns.str.strip()
        df_creditos.columns = df_creditos.columns.str.strip()
        df_pagos.columns = df_pagos.columns.str.strip()
        
        return df_clientes, df_creditos, df_pagos
    except Exception as e:
        st.error(f"Error al cargar el archivo de Excel: {e}")
        return None, None, None

df_clientes, df_creditos, df_pagos = load_data()

# ---------------------------------------------------------
# MENÚ DE NAVEGACIÓN LATERAL
# ---------------------------------------------------------
st.sidebar.title("🤝 Entre Amigos Capital")
st.sidebar.caption("Créditos justos sobre la base de la confianza")

opcion_menu = st.sidebar.radio(
    "Selecciona una sección:",
    ["📊 Dashboard General", "👤 Ficha por Cliente", "🧮 Simulador de Créditos", "ℹ️ Sobre Nosotros & Políticas"]
)

if df_clientes is not None and df_creditos is not None and df_pagos is not None:

    # =========================================================
    # 1. DASHBOARD GENERAL
    # =========================================================
    if opcion_menu == "📊 Dashboard General":
        st.title("📊 Control General de Cartera")
        st.markdown("---")
        
        # Cálculos de Métricas
        capital_prestado = df_creditos['Monto_Prestado'].sum() if 'Monto_Prestado' in df_creditos.columns else 0
        capital_pagado = df_pagos['Monto_Pago'].sum() if 'Monto_Pago' in df_pagos.columns else 0
        
        if 'Saldo_Pendiente' in df_creditos.columns:
            saldo_pendiente = df_creditos['Saldo_Pendiente'].sum()
        else:
            saldo_pendiente = capital_prestado - capital_pagado

        col1, col2, col3 = st.columns(3)
        col1.metric("Capital Prestado Total", f"${capital_prestado:,.0f} COP")
        col2.metric("Capital Recaudado", f"${capital_pagado:,.0f} COP")
        col3.metric("Saldo Pendiente en Cartera", f"${saldo_pendiente:,.0f} COP")

        st.markdown("---")

        col_left, col_right = st.columns([1, 1])

        with col_left:
            st.subheader("Estado General de Cartera")
            df_estado = pd.DataFrame({
                "Estado": ["Capital Recaudado", "Saldo Pendiente"],
                "Monto": [capital_pagado, max(0, saldo_pendiente)]
            })
            fig_pie = px.pie(df_estado, names="Estado", values="Monto", hole=0.4,
                             color_discrete_sequence=["#2ecc71", "#e74c3c"])
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_right:
            st.subheader("Resumen de Créditos Registrados")
            st.dataframe(df_creditos, use_container_width=True)

    # =========================================================
    # 2. FICHA POR CLIENTE CON HISTORIAL DE CONFIANZA
    # =========================================================
    elif opcion_menu == "👤 Ficha por Cliente":
        st.title("👤 Ficha de Cliente e Historial de Confianza")
        st.markdown("---")

        lista_clientes = df_clientes['Nombre'].unique() if 'Nombre' in df_clientes.columns else []
        cliente_sel = st.selectbox("Selecciona un cliente:", lista_clientes)

        if cliente_sel:
            info_cliente = df_clientes[df_clientes['Nombre'] == cliente_sel].iloc[0]
            
            # Obtener créditos y pagos del cliente
            id_cliente = info_cliente.get('ID_Cliente', None)
            creditos_cli = df_creditos[df_creditos['ID_Cliente'] == id_cliente] if id_cliente else pd.DataFrame()
            pagos_cli = df_pagos[df_pagos['ID_Cliente'] == id_cliente] if id_cliente else pd.DataFrame()

            # Métricas del cliente
            monto_prestado_cli = creditos_cli['Monto_Prestado'].sum() if not creditos_cli.empty else 0
            monto_pagado_cli = pagos_cli['Monto_Pago'].sum() if not pagos_cli.empty else 0
            saldo_cli = monto_prestado_cli - monto_pagado_cli

            # Indicador de Confianza / Scoring
            st.subheader("⭐ Calificación de Confianza")
            if saldo_cli <= 0 and monto_prestado_cli > 0:
                st.success("🟢 **Historial Excelente:** Crédito cancelado en su totalidad. Apto para nuevos créditos con aumento de cupo.")
            elif saldo_cli > 0:
                st.info("🟡 **Crédito Activo:** En proceso de pago puntual bajo principio de buena fe.")
            else:
                st.warning("⚪ **Sin Créditos Activos:** No registra operaciones pendientes.")

            st.markdown("---")

            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("Información Personal")
                st.write(f"**Nombre:** {info_cliente.get('Nombre', 'N/A')}")
                st.write(f"**Contacto:** {info_cliente.get('Telefono', 'N/A')}")
                st.write(f"**Cédula/ID:** {info_cliente.get('Documento', 'N/A')}")

            with col_b:
                st.subheader("Resumen Financiero")
                st.write(f"**Total Prestado:** ${monto_prestado_cli:,.0f} COP")
                st.write(f"**Total Pagado:** ${monto_pagado_cli:,.0f} COP")
                st.write(f"**Saldo Actual:** ${saldo_cli:,.0f} COP")

            st.markdown("---")
            st.subheader("Historial de Pagos")
            st.dataframe(pagos_cli, use_container_width=True)

    # =========================================================
    # 3. SIMULADOR DE CRÉDITOS
    # =========================================================
    elif opcion_menu == "🧮 Simulador de Créditos":
        st.title("🧮 Simulador de Créditos - Entre Amigos Capital")
        st.markdown("Calcule la cuota estimada para su próximo préstamo a una **tasa fija del 3% mensual**.")
        st.markdown("---")

        col_sim1, col_sim2 = st.columns(2)

        with col_sim1:
            monto_sim = st.number_input(
                "Monto a solicitar (COP):",
                min_value=100000,
                max_value=5000000,
                value=500000,
                step=50000,
                format="%d"
            )

            modalidad_sim = st.selectbox(
                "Modalidad de pago:",
                [
                    "Intereses periódicos + Capital al final",
                    "Cuotas fijas (Capital + Interés)"
                ]
            )

            plazo_sim = st.slider("Plazo en meses:", min_value=1, max_value=24, value=6)

        tasa_mensual = 0.03 # 3% fijo

        with col_sim2:
            st.subheader("📋 Resumen de la Simulación")
            
            if modalidad_sim == "Intereses periódicos + Capital al final":
                interes_mensual = monto_sim * tasa_mensual
                total_intereses = interes_mensual * plazo_sim
                total_pagar = monto_sim + total_intereses

                st.write(f"• **Pago mensual de intereses:** ${interes_mensual:,.0f} COP")
                st.write(f"• **Pago final de capital (Mes {plazo_sim}):** ${monto_sim:,.0f} COP")
                st.write(f"• **Total de intereses a pagar:** ${total_intereses:,.0f} COP")
                st.metric("Total General a Cancelar", f"${total_pagar:,.0f} COP")

            else:
                # Amortización con cuota fija (Fórmula de Anualidad)
                cuota_mensual = (monto_sim * tasa_mensual) / (1 - (1 + tasa_mensual)**(-plazo_sim))
                total_pagar = cuota_mensual * plazo_sim
                total_intereses = total_pagar - monto_sim

                st.write(f"• **Cuota fija mensual:** ${cuota_mensual:,.0f} COP")
                st.write(f"• **Total intereses aproximados:** ${total_intereses:,.0f} COP")
                st.metric("Total General a Cancelar", f"${total_pagar:,.0f} COP")

        st.info("💡 **Nota:** Esta simulación es meramente informativa y está sujeta a la aprobación de cupo según el historial del cliente.")

    # =========================================================
    # 4. SOBRE NOSOTROS & POLÍTICAS
    # =========================================================
    elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
        st.title("🤝 Sobre Entre Amigos Capital")
        st.markdown("---")

        st.subheader("💡 Nuestra Propuesta de Valor")
        st.write("""
        Ofrecemos acceso a microcréditos ágiles para familiares y amigos que no cuentan con acceso fácil a la banca tradicional.
        Creemos firmemente en el apoyo mutuo, operando con **tasas justas**, cero costos ocultos y total transparencia.
        """)

        st.subheader("📜 Políticas y Reglas de Juego")
        col_p1, col_p2 = st.columns(2)

        with col_p1:
            st.markdown("""
            **1. Montos y Plazos**
            * **Rango:** Desde $100.000 hasta $5.000.000 COP.
            * **Frecuencia:** Pagos mensuales.
            * **Tasa de interés:** 3% mensual fijo.

            **2. Modalidades de Pago**
            * **Intereses Periódicos + Capital Final:** Ideal para quienes necesitan liquidez de trabajo y liquidan el capital al terminar el plazo.
            * **Abono Progresivo:** Cuotas compuestas de capital e intereses para saldar la deuda mes a mes.
            """)

        with col_p2:
            st.markdown("""
            **3. Política de Mora y Buena Fe**
            * **Cero Sanciones:** No aplicamos multas moratorias financieras.
            * **Acompañamiento:** Si atraviesas por una dificultad, conversamos para reestructurar el pago.
            * **Historial de Confianza:** El cumplimiento puntual abre puertas para incrementos de cupos e historial positivo en futuros créditos.

            **4. Medios de Pago Habilitados**
            * Transferencia electrónica (Bancolombia, Nequi, Daviplata).
            * Efectivo.
            """)

else:
    st.warning("Cargando datos o verificando el archivo...")
