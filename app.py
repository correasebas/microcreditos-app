import streamlit as st
import pandas as pd
import plotly.express as px

# Configuración de la página
st.set_page_config(
    page_title="Gestión de Microcréditos",
    page_icon="💰",
    layout="wide"
)

# Cargar datos desde el archivo Excel
EXCEL_FILE = "proyecto microcréditos copia 3.xlsx"

@st.cache_data(ttl=60)
def load_data():
    clientes = pd.read_excel(EXCEL_FILE, sheet_name="Clientes")
    creditos = pd.read_excel(EXCEL_FILE, sheet_name="Creditos")
    pagos = pd.read_excel(EXCEL_FILE, sheet_name="Pagos")
    estado_cartera = pd.read_excel(EXCEL_FILE, sheet_name="Estado_Cartera")
    resumen = pd.read_excel(EXCEL_FILE, sheet_name="Resumen_Cartera")
    return clientes, creditos, pagos, estado_cartera, resumen

try:
    clientes_df, creditos_df, pagos_df, cartera_df, resumen_df = load_data()
except Exception as e:
    st.error(f"Error al cargar el archivo de Excel: {e}")
    st.stop()

# Menú lateral
st.sidebar.title("📌 Menú de Navegación")
opcion = st.sidebar.radio("Selecciona una opción:", ["Dashboard General", "Consulta por Cliente"])

# ---------------------------------------------------------
# OPCIÓN 1: DASHBOARD GENERAL
# ---------------------------------------------------------
if opcion == "Dashboard General":
    st.title("📊 Panel General de Control")
    st.markdown("---")

    # Tarjetas de Métricas Principales
    col1, col2, col3, col4 = st.columns(4)
    
    cap_prestado = cartera_df["capital_inicial"].sum()
    cap_pagado = cartera_df["capital_pagado"].sum()
    deuda_total = cartera_df["deuda_total_pendiente"].sum()
    deuda_vencida = cartera_df["deuda_vencida"].sum()

    col1.metric("Capital Prestado", f"${cap_prestado:,.0f}")
    col2.metric("Capital Pagado", f"${cap_pagado:,.0f}")
    col3.metric("Deuda Total Pendiente", f"${deuda_total:,.0f}")
    col4.metric("Deuda Vencida", f"${deuda_vencida:,.0f}", delta_color="inverse")

    st.markdown("---")

    # Gráficos e indicadores
    col_g1, col_g2 = st.columns(2)

    with col_g1:
        st.subheader("Estado de Créditos")
        fig_estado = px.pie(
            cartera_df, 
            names="estado", 
            title="Distribución por Estado de Crédito",
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        st.plotly_chart(fig_estado, use_container_width=True)

    with col_g2:
        st.subheader("Tabla de Estado de Cartera")
        st.dataframe(
            cartera_df[["credito_id", "cliente_id", "capital_inicial", "deuda_total_pendiente", "deuda_vencida", "estado"]],
            use_container_width=True,
            hide_index=True
        )

# ---------------------------------------------------------
# OPCIÓN 2: CONSULTA POR CLIENTE
# ---------------------------------------------------------
elif opcion == "Consulta por Cliente":
    st.title("🔍 Consulta de Cliente / Crédito")
    st.markdown("---")

    # Selector de Cliente
    clientes_lista = clientes_df["nombre"].dropna().unique().tolist()
    cliente_sel = st.selectbox("Selecciona un Cliente:", sorted(clientes_lista))

    if cliente_sel:
        # Obtener ID del cliente
        cliente_row = clientes_df[clientes_df["nombre"] == cliente_sel].iloc[0]
        c_id = cliente_row["cliente_id"]

        st.subheader(f"Ficha de: {cliente_sel}")
        st.write(f"**Teléfono:** {cliente_row.get('telefono', 'N/A')} | **Parentesco / Nota:** {cliente_row.get('parentesco', 'N/A')}")

        # Créditos del cliente
        creditos_cli = cartera_df[cartera_df["cliente_id"] == c_id]

        if not creditos_cli.empty:
            st.markdown("### Créditos Asociados")
            st.dataframe(
                creditos_cli[["credito_id", "tipo_interes", "capital_inicial", "capital_pagado", "deuda_total_pendiente", "deuda_vencida", "estado"]],
                use_container_width=True,
                hide_index=True
            )

            # Historial de Pagos
            pagos_cli = pagos_df[pagos_df["cliente_id"] == c_id]
            st.markdown("### Historial de Pagos")
            if not pagos_cli.empty:
                st.dataframe(
                    pagos_cli[["pago_id", "credito_id", "fecha_pago", "valor_pago", "pago_capital", "pago_interes", "medio_pago"]],
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("Este cliente no registra pagos aún.")
        else:
            st.warning("No se encontraron créditos registrados para este cliente.")