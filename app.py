import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import urllib.parse
from fpdf import FPDF
import os
from dateutil.relativedelta import relativedelta
from groq import Groq

# ---------------------------------------------------------
# 0. CONFIGURACIÓN NATIVA DE STREAMLIT (config.toml)
# ---------------------------------------------------------
os.makedirs('.streamlit', exist_ok=True)
config_content = """
[theme]
base="dark"
primaryColor="#00D26A"
backgroundColor="#0A1118"
secondaryBackgroundColor="#111B27"
textColor="#E6EDF3"
font="sans serif"
"""
with open('.streamlit/config.toml', 'w') as f:
    f.write(config_content)

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA & ESTILO FINTECH (Azul & Esmeralda)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Entre Amigos Capital - Fondo Familiar",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS avanzados para unificar el diseño corporativo
st.markdown("""
    <style>
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stToolbar"], .main {
        background-color: #0A1118 !important;
        color: #E6EDF3 !important;
    }
    html, body, [class*="css"], p, span, label, h1, h2, h3, h4, h5, h6 {
        color: #E6EDF3 !important;
        font-family: 'Inter', sans-serif !important;
    }
    [data-testid="stSidebar"], [data-testid="stSidebarContent"] {
        background-color: #111B27 !important;
        border-right: 1px solid #1E2D3D !important;
    }
    [data-testid="stSidebar"] * {
        color: #E6EDF3 !important;
    }
    div[data-testid="stMetric"] {
        background-color: #111B27 !important;
        border: 1px solid #1E2D3D !important;
        padding: 16px !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4) !important;
    }
    div[data-testid="stMetric"] label {
        color: #8B949E !important;
        font-size: 0.85rem !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #00D26A !important;
        font-weight: 700 !important;
    }
    .stButton > button, .stLinkButton > a {
        background: linear-gradient(135deg, #00A859 0%, #00D26A 100%) !important;
        color: #0A1118 !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.5rem 1rem !important;
        transition: all 0.3s ease !important;
    }
    .stButton > button:hover, .stLinkButton > a:hover {
        opacity: 0.9 !important;
        box-shadow: 0 0 15px rgba(0, 210, 106, 0.4) !important;
    }
    .stSelectbox div[data-baseweb="select"] > div, 
    .stDateInput input, .stNumberInput input, .stTextInput input, div[data-baseweb="input"] {
        background-color: #111B27 !important;
        color: #E6EDF3 !important;
        border-color: #1E2D3D !important;
        border-radius: 8px !important;
    }
    [data-testid="stDataFrame"], .dataframe {
        background-color: #111B27 !important;
        border: 1px solid #1E2D3D !important;
        border-radius: 8px !important;
    }
    div[data-testid="stExpander"] {
        background-color: #111B27 !important;
        border: 1px solid #1E2D3D !important;
        border-radius: 8px !important;
    }
    hr {
        border-color: #1E2D3D !important;
    }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE_DEFAULT = "proyecto microcréditos copia 3.xlsx.xlsx"

# ---------------------------------------------------------
# CLASE PDF CON LA PALETA INSTITUCIONAL
# ---------------------------------------------------------
class ComprobantePDF(FPDF):
    def header(self):
        self.set_fill_color(17, 27, 39)
        self.rect(0, 0, 210, 32, 'F')
        self.set_font("Arial", "B", 16)
        self.set_text_color(0, 210, 106)
        self.cell(0, 8, "Entre Amigos Capital", ln=True, align="C")
        self.set_font("Arial", "", 10)
        self.set_text_color(230, 237, 243)
        self.cell(0, 5, "Comprobante Oficial de Recaudo de Pago", ln=True, align="C")
        self.ln(10)

def generar_pdf_comprobante(pago_info):
    pdf = ComprobantePDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_fill_color(0, 210, 106)
    pdf.set_text_color(10, 17, 24)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 8, f"RECIBO N°: {pago_info['pago_id']}", ln=True, align="C", fill=True)
    pdf.ln(4)

    pdf.set_text_color(17, 27, 39)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Datos del Cliente y Crédito", ln=True)
    pdf.set_draw_color(30, 45, 61)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(50, 5, "Cliente:", 0)
    pdf.cell(0, 5, f"{pago_info['cliente_nombre']} ({pago_info['cliente_id']})", ln=True)
    pdf.cell(50, 5, "Crédito N°:", 0)
    pdf.cell(0, 5, str(pago_info['credito_id']), ln=True)
    pdf.cell(50, 5, "Fecha de Pago:", 0)
    pdf.cell(0, 5, str(pago_info['fecha_pago']), ln=True)
    pdf.cell(50, 5, "Medio de Pago:", 0)
    pdf.cell(0, 5, str(pago_info['medio_pago']), ln=True)
    pdf.ln(4)

    pdf.set_text_color(17, 27, 39)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Desglose de la Transacción", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(120, 5, "Abono a Intereses:", 0)
    pdf.cell(0, 5, f"${pago_info['pago_interes']:,.0f} COP", ln=True, align="R")
    pdf.cell(120, 5, "Abono a Capital:", 0)
    pdf.cell(0, 5, f"${pago_info['pago_capital']:,.0f} COP", ln=True, align="R")
    
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(17, 27, 39)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(120, 7, f"TOTAL RECIBIDO ({pago_info['concepto']}):", fill=True)
    pdf.set_text_color(0, 210, 106)
    pdf.cell(0, 7, f"${pago_info['valor_pago']:,.0f} COP", ln=True, align="R", fill=True)
    pdf.ln(4)

    pdf.set_text_color(17, 27, 39)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Estado Actualizado de la Deuda", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(50, 5, "Capital Pendiente:", 0)
    pdf.cell(0, 5, f"${pago_info['nuevo_cap_pend']:,.0f} COP", ln=True)
    pdf.cell(50, 5, "Intereses Pendientes:", 0)
    pdf.cell(0, 5, f"${pago_info['nuevo_int_pend']:,.0f} COP", ln=True)
    
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(192, 57, 43)
    pdf.cell(50, 5, "Deuda Total Pendiente:", 0)
    pdf.cell(0, 5, f"${pago_info['nueva_deuda_total']:,.0f} COP", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(50, 5, "Estado del Crédito:", 0)
    pdf.cell(0, 5, str(pago_info['nuevo_estado']), ln=True)

    if pago_info.get('observaciones'):
        pdf.cell(50, 5, "Observaciones:", 0)
        pdf.cell(0, 5, str(pago_info['observaciones']), ln=True)

    pdf.ln(10)
    pdf.set_font("Arial", "I", 8)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 4, "Gracias por mantener tu crédito al día. Este documento sirve como soporte oficial de recaudo.", ln=True, align="C")
    pdf.cell(0, 4, "Entre Amigos Capital - Crecimiento financiero basado en la confianza.", ln=True, align="C")

    return bytes(pdf.output())

# ---------------------------------------------------------
# MENÚ LATERAL & BRANDING Y MISIÓN DEL FONDO
# ---------------------------------------------------------
st.sidebar.title("💎 Entre Amigos Capital")
st.sidebar.caption("Fondo de Inversión y Microcréditos Familiares")

with st.sidebar.expander("📌 Nuestra Misión", expanded=False):
    st.write(
        "Fomentar el desarrollo económico y la colaboración financiera "
        "dentro de nuestro círculo de confianza, ofreciendo liquidez ágil, "
        "tasas justas y transparencia absoluta en cada operación."
    )

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Base de Datos Excel")
uploaded_file = st.sidebar.file_uploader(
    "Carga tu archivo de Excel actualizado:", 
    type=["xlsx"],
    help="Si no subes un archivo, se cargará el archivo base por defecto."
)

@st.cache_data(show_spinner=False)
def load_data_from_file(file_source):
    try:
        xls = pd.ExcelFile(file_source)
        df_clientes = pd.read_excel(xls, sheet_name='Clientes')
        df_creditos = pd.read_excel(xls, sheet_name='Creditos')
        df_pagos = pd.read_excel(xls, sheet_name='Pagos')
        df_estado_cartera = pd.read_excel(xls, sheet_name='Estado_Cartera') if 'Estado_Cartera' in xls.sheet_names else pd.DataFrame()
        df_resumen = pd.read_excel(xls, sheet_name='Resumen_Cartera') if 'Resumen_Cartera' in xls.sheet_names else pd.DataFrame()
        df_calendario = pd.read_excel(xls, sheet_name='Calendario_Intereses') if 'Calendario_Intereses' in xls.sheet_names else pd.DataFrame()

        return df_clientes, df_creditos, df_pagos, df_estado_cartera, df_resumen, df_calendario
    except Exception as e:
        st.error(f"Error al cargar el archivo de Excel: {e}")
        return None, None, None, None, None, None

file_to_load = uploaded_file if uploaded_file is not None else EXCEL_FILE_DEFAULT

if 'current_loaded_file' not in st.session_state or st.session_state['current_loaded_file'] != file_to_load:
    df_c, df_cr, df_p, df_ec, df_res, df_cal = load_data_from_file(file_to_load)
    
    st.session_state['df_clientes'] = df_c if df_c is not None else pd.DataFrame()
    st.session_state['df_creditos'] = df_cr if df_cr is not None else pd.DataFrame()
    st.session_state['df_pagos'] = df_p if df_p is not None else pd.DataFrame()
    st.session_state['df_estado_cartera'] = df_ec if df_ec is not None else pd.DataFrame()
    st.session_state['df_calendario'] = df_cal if df_cal is not None else pd.DataFrame()
    st.session_state['current_loaded_file'] = file_to_load

if 'df_clientes' not in st.session_state:
    st.session_state['df_clientes'] = pd.DataFrame()
if 'df_creditos' not in st.session_state:
    st.session_state['df_creditos'] = pd.DataFrame()
if 'df_estado_cartera' not in st.session_state:
    st.session_state['df_estado_cartera'] = pd.DataFrame()
if 'df_calendario' not in st.session_state:
    st.session_state['df_calendario'] = pd.DataFrame()
if 'df_pagos' not in st.session_state:
    st.session_state['df_pagos'] = pd.DataFrame()

df_clientes = st.session_state['df_clientes']
df_creditos = st.session_state['df_creditos']
df_pagos = st.session_state['df_pagos']
df_estado_cartera = st.session_state['df_estado_cartera']
df_calendario = st.session_state['df_calendario']

st.sidebar.markdown("---")

opcion_menu = st.sidebar.radio(
    "Selecciona una sección:",
    [
        "📊 Dashboard General", 
        "👤 Ficha por Cliente", 
        "➕ Nuevos Registros", 
        "📝 Registrar Pago", 
        "⚖️ Gestión de Cobro", 
        "🧮 Simulador de Créditos", 
        "🤖 Asistente IA", 
        "ℹ️ Sobre Nosotros & Políticas"
    ]
)

# ---------------------------------------------------------
# FUNCIONES AUXILIARES DE RECALCULO Y EXPORTACIÓN
# ---------------------------------------------------------
def recalcular_resumen_cartera():
    if df_creditos.empty or df_estado_cartera.empty:
        return pd.DataFrame()
    
    cap_prestado = df_creditos['capital_inicial'].sum()
    cap_pagado = df_estado_cartera['capital_pagado'].sum()
    cap_pendiente = df_estado_cartera['capital_pendiente'].sum()
    int_pendiente = df_estado_cartera['interes_pendiente'].sum()
    deuda_total = df_estado_cartera['deuda_total_pendiente'].sum()
    deuda_vencida = df_estado_cartera['deuda_vencida'].sum()
    
    creditos_activos = len(df_estado_cartera[df_estado_cartera['deuda_total_pendiente'] > 0])
    creditos_mora = len(df_estado_cartera[df_estado_cartera['estado'].astype(str).str.contains('mora', case=False, na=False)])
    creditos_aldia = len(df_estado_cartera[(df_estado_cartera['estado'].astype(str).str.contains('día', case=False, na=False)) & (df_estado_cartera['deuda_total_pendiente'] > 0)])
    creditos_finalizados = len(df_estado_cartera[df_estado_cartera['deuda_total_pendiente'] <= 0])

    data_resumen = {
        'Indicador': [
            'Capital total prestado', 'Capital pagado', 'Capital pendiente',
            'Intereses pendientes', 'Deuda total pendiente', 'Deuda vencida',
            'Créditos activos', 'Créditos en mora', 'Créditos al día', 'Créditos finalizados'
        ],
        'Resultado': [
            cap_prestado, cap_pagado, cap_pendiente,
            int_pendiente, deuda_total, deuda_vencida,
            creditos_activos, creditos_mora, creditos_aldia, creditos_finalizados
        ]
    }
    return pd.DataFrame(data_resumen)

def exportar_excel_completo():
    output = io.BytesIO()
    df_resumen_actualizado = recalcular_resumen_cartera()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not st.session_state['df_clientes'].empty:
            st.session_state['df_clientes'].to_excel(writer, sheet_name='Clientes', index=False)
        if not st.session_state['df_creditos'].empty:
            cols_cred = [c for c in st.session_state['df_creditos'].columns if not str(c).startswith('Unnamed')]
            st.session_state['df_creditos'][cols_cred].to_excel(writer, sheet_name='Creditos', index=False)
        if not st.session_state['df_pagos'].empty:
            cols_pagos = [c for c in st.session_state['df_pagos'].columns if not str(c).startswith('Unnamed')]
            st.session_state['df_pagos'][cols_pagos].to_excel(writer, sheet_name='Pagos', index=False)
        if not st.session_state['df_estado_cartera'].empty:
            cols_ec = [c for c in st.session_state['df_estado_cartera'].columns if not str(c).startswith('Unnamed')]
            st.session_state['df_estado_cartera'][cols_ec].to_excel(writer, sheet_name='Estado_Cartera', index=False)
        if not st.session_state['df_calendario'].empty:
            cols_cal = [c for c in st.session_state['df_calendario'].columns if not str(c).startswith('Unnamed')]
            st.session_state['df_calendario'][cols_cal].to_excel(writer, sheet_name='Calendario_Intereses', index=False)
        if not df_resumen_actualizado.empty:
            df_resumen_actualizado.to_excel(writer, sheet_name='Resumen_Cartera', index=False)
            
    return output.getvalue()

if not df_creditos.empty or not df_clientes.empty:

    # =========================================================
    # 1. DASHBOARD GENERAL
    # =========================================================
    if opcion_menu == "📊 Dashboard General":
        st.title("📊 Control General de Cartera")
        st.markdown("---")

        df_res = recalcular_resumen_cartera()
        dict_res = dict(zip(df_res['Indicador'], df_res['Resultado'])) if not df_res.empty else {}

        val_prestado = dict_res.get('Capital total prestado', 0)
        val_pagado = dict_res.get('Capital pagado', 0)
        val_cap_pendiente = dict_res.get('Capital pendiente', 0)
        val_int_pendiente = dict_res.get('Intereses pendientes', 0)
        val_deuda_total = dict_res.get('Deuda total pendiente', 0)
        val_deuda_vencida = dict_res.get('Deuda vencida', 0)
        creditos_activos = dict_res.get('Créditos activos', 0)
        creditos_mora = dict_res.get('Créditos en mora', 0)
        creditos_aldia = dict_res.get('Créditos al día', 0)

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
                             color_discrete_map={"Al Día": "#00D26A", "En Mora": "#E74C3C"})
            fig_pie.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", 
                plot_bgcolor="rgba(0,0,0,0)", 
                font_color="#E6EDF3",
                legend=dict(font=dict(color="#E6EDF3"))
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_right:
            st.subheader("Detalle de Créditos Registrados")
            cols_show = [c for c in df_creditos.columns if not str(c).startswith('Unnamed')]
            st.dataframe(df_creditos[cols_show], use_container_width=True)

    # =========================================================
    # 2. FICHA POR CLIENTE
    # =========================================================
    elif opcion_menu == "👤 Ficha por Cliente":
        st.title("👤 Ficha de Cliente e Historial de Deuda")
        st.markdown("---")

        if not df_clientes.empty:
            lista_clientes = df_clientes['nombre'].dropna().unique()
            cliente_sel = st.selectbox("Selecciona un cliente:", lista_clientes)

            if cliente_sel:
                info_cliente = df_clientes[df_clientes['nombre'] == cliente_sel].iloc[0]
                cliente_id = info_cliente['cliente_id']

                creditos_cliente = df_creditos[df_creditos['cliente_id'] == cliente_id]
                pagos_cliente = st.session_state['df_pagos'][st.session_state['df_pagos']['cliente_id'] == cliente_id]
                cartera_cliente = st.session_state['df_estado_cartera'][st.session_state['df_estado_cartera']['cliente_id'] == cliente_id] if not st.session_state['df_estado_cartera'].empty else pd.DataFrame()

                cap_pendiente = cartera_cliente['capital_pendiente'].sum() if not cartera_cliente.empty else 0
                int_pendiente = cartera_cliente['interes_pendiente'].sum() if not cartera_cliente.empty else 0
                deuda_vencida = cartera_cliente['deuda_vencida'].sum() if not cartera_cliente.empty else 0
                deuda_total = cartera_cliente['deuda_total_pendiente'].sum() if not cartera_cliente.empty else 0

                tiene_mora = any(cartera_cliente['estado'].astype(str).str.contains('mora', case=False, na=False)) if not cartera_cliente.empty else False

                if tiene_mora or deuda_vencida > 0:
                    st.error(f"⚠️ **Alerta Individual - En Mora:** Este cliente presenta cuotas vencidas por un valor total de **${deuda_vencida:,.0f} COP** (Deuda total pendiente: ${deuda_total:,.0f} COP).")
                elif deuda_total == 0:
                    st.success("🟢 **Paz y Salvo:** El cliente no presenta saldos pendientes.")
                else:
                    st.info("🟢 **Al Día:** El cliente cuenta con sus cuotas e intereses al día.")

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
                
                st.subheader("Créditos del Cliente")
                if not creditos_cliente.empty:
                    cols_c = [c for c in creditos_cliente.columns if not str(c).startswith('Unnamed')]
                    st.dataframe(creditos_cliente[cols_c], use_container_width=True)

                st.subheader("Historial de Pagos")
                if not pagos_cliente.empty:
                    cols_validas = [c for c in pagos_cliente.columns if not str(c).startswith('Unnamed')]
                    st.dataframe(pagos_cliente[cols_validas], use_container_width=True)
        else:
            st.warning("No hay clientes registrados en el sistema.")

    # =========================================================
    # 3. NUEVOS REGISTROS (CLIENTES Y PRÉSTAMOS INTEGRADOS)
    # =========================================================
    elif opcion_menu == "➕ Nuevos Registros":
        st.title("➕ Módulo Integrado de Nuevos Registros")
        st.markdown("Da de alta nuevos clientes y otorga créditos actualizando automáticamente todas las hojas del Excel (Créditos, Estado de Cartera y Calendario de Intereses).")
        st.markdown("---")

        tab_cli, tab_cred = st.tabs(["👤 Registrar Nuevo Cliente", "💳 Otorgar Nuevo Préstamo"])

        with tab_cli:
            st.subheader("Ingresar Nuevo Cliente al Fondo")
            with st.form("form_nuevo_cliente"):
                col_nc1, col_nc2 = st.columns(2)
                with col_nc1:
                    nombre_nuevo = st.text_input("Nombre Completo:")
                    telefono_nuevo = st.text_input("Teléfono / WhatsApp (ej. 3001234567):")
                with col_nc2:
                    parentesco_nuevo = st.selectbox("Parentesco / Relación con el Fondo:", ["Familiar", "Amigo", "Conocido", "Socio"])
                    estado_cli_nuevo = st.selectbox("Estado del Cliente:", ["Activo", "Inactivo"])

                btn_guardar_cli = st.form_submit_button("💾 Guardar Nuevo Cliente", type="primary")

                if btn_guardar_cli:
                    if nombre_nuevo:
                        existentes_cli_ids = st.session_state['df_clientes']['cliente_id'].dropna().tolist() if not st.session_state['df_clientes'].empty else []
                        nums_c = [int(str(x).replace('CLI', '')) for x in existentes_cli_ids if str(x).startswith('CLI') and str(x).replace('CLI', '').isdigit()]
                        nuevo_num_c = max(nums_c) + 1 if nums_c else 1
                        proximo_cli_id = f"CLI{nuevo_num_c:03d}"

                        nueva_fila_cliente = {
                            'cliente_id': proximo_cli_id,
                            'nombre': nombre_nuevo,
                            'telefono': telefono_nuevo,
                            'fecha_registro': datetime.today().strftime('%Y-%m-%d'),
                            'estado_cliente': estado_cli_nuevo,
                            'parentesco': parentesco_nuevo
                        }

                        st.session_state['df_clientes'] = pd.concat([st.session_state['df_clientes'], pd.DataFrame([nueva_fila_cliente])], ignore_index=True)
                        st.success(f"✅ ¡Cliente **{nombre_nuevo}** registrado con éxito bajo el ID `{proximo_cli_id}`! Ya puedes seleccionarlo en la pestaña de préstamos.")
                    else:
                        st.error("⚠️ El campo de nombre completo es obligatorio.")

        with tab_cred:
            st.subheader("Otorgar Crédito y Sincronizar Hojas de Excel")
            
            df_clientes_actual = st.session_state.get('df_clientes', pd.DataFrame())

            if df_clientes_actual.empty:
                st.warning("⚠️ Primero debes registrar al menos un cliente en la pestaña anterior.")
            else:
                with st.form("form_nuevo_prestamo"):
                    lista_c_nombres = df_clientes_actual['nombre'].dropna().unique()
                    cli_sel_cred = st.selectbox("Selecciona al Cliente Beneficiario:", lista_c_nombres)

                    info_c_sel = df_clientes_actual[df_clientes_actual['nombre'] == cli_sel_cred].iloc[0]
                    id_cli_sel = info_c_sel['cliente_id']

                    col_np1, col_np2 = st.columns(2)
                    with col_np1:
                        capital_inicial = st.number_input("Capital Inicial del Préstamo (COP):", min_value=100000, value=1000000, step=50000, format="%d")
                        tasa_interes_mensual = st.number_input("Tasa de Interés Mensual (%):", min_value=0.0, value=3.0, step=0.5) / 100.0
                        plazo_meses = st.number_input("Plazo en Meses:", min_value=1, value=6, step=1)
                    with col_np2:
                        fecha_desembolso = st.date_input("Fecha de Desembolso:", datetime.today())
                        modalidad_cred = st.selectbox("Modalidad de Pago:", ["Intereses periódicos + Capital al final", "Cuotas fijas (Capital + Interés)"])

                    btn_guardar_cred = st.form_submit_button("🚀 Generar y Sincronizar Préstamo", type="primary")

                    if btn_guardar_cred:
                        existentes_cr_ids = st.session_state['df_creditos']['credito_id'].dropna().tolist() if not st.session_state['df_creditos'].empty else []
                        nums_cr = [int(str(x).replace('CR', '')) for x in existentes_cr_ids if str(x).startswith('CR') and str(x).replace('CR', '').isdigit()]
                        nuevo_num_cr = max(nums_cr) + 1 if nums_cr else 1
                        proximo_cred_id = f"CR{nuevo_num_cr:03d}"

                        nueva_fila_credito = {
                            'credito_id': proximo_cred_id,
                            'cliente_id': id_cli_sel,
                            'capital_inicial': capital_inicial,
                            'tasa_interes': tasa_interes_mensual,
                            'plazo_meses': plazo_meses,
                            'fecha_desembolso': pd.to_datetime(fecha_desembolso),
                            'modalidad': modalidad_cred,
                            'estado_credito': 'Activo',
                            'saldo_capital': capital_inicial
                        }
                        st.session_state['df_creditos'] = pd.concat([st.session_state['df_creditos'], pd.DataFrame([nueva_fila_credito])], ignore_index=True)

                        interes_inicial_est = capital_inicial * tasa_interes_mensual
                        nueva_fila_ec = {
                            'credito_id': proximo_cred_id,
                            'cliente_id': id_cli_sel,
                            'capital_pagado': 0,
                            'capital_pendiente': capital_inicial,
                            'interes_pendiente': interes_inicial_est,
                            'deuda_vencida': 0,
                            'deuda_total_pendiente': capital_inicial + interes_inicial_est,
                            'estado': 'Al día'
                        }
                        st.session_state['df_estado_cartera'] = pd.concat([st.session_state['df_estado_cartera'], pd.DataFrame([nueva_fila_ec])], ignore_index=True)

                        nuevas_filas_cal = []
                        if modalidad_cred == "Intereses periódicos + Capital al final":
                            cuota_val = capital_inicial * tasa_interes_mensual
                            for i in range(1, int(plazo_meses) + 1):
                                fecha_cuota = pd.to_datetime(fecha_desembolso) + relativedelta(months=i)
                                es_ultima = (i == plazo_meses)
                                cap_cuota = capital_inicial if es_ultima else 0
                                total_cuota = cuota_val + cap_cuota
                                
                                nuevas_filas_cal.append({
                                    'credito_id': proximo_cred_id,
                                    'numero_cuota': i,
                                    'fecha_programada': fecha_cuota,
                                    'capital_programado': cap_cuota,
                                    'interes_programado': cuota_val,
                                    'valor_cuota_calculado': total_cuota,
                                    'interes_pagado': 0,
                                    'interes_pendiente': cuota_val,
                                    'estado_cuota': 'Pendiente'
                                })
                        else: 
                            cuota_fija = (capital_inicial * tasa_interes_mensual) / (1 - (1 + tasa_interes_mensual)**(-plazo_meses))
                            saldo_temp = capital_inicial
                            for i in range(1, int(plazo_meses) + 1):
                                fecha_cuota = pd.to_datetime(fecha_desembolso) + relativedelta(months=i)
                                int_cuota = saldo_temp * tasa_interes_mensual
                                cap_cuota = cuota_fija - int_cuota
                                saldo_temp -= cap_cuota

                                nuevas_filas_cal.append({
                                    'credito_id': proximo_cred_id,
                                    'numero_cuota': i,
                                    'fecha_programada': fecha_cuota,
                                    'capital_programado': max(0, cap_cuota),
                                    'interes_programado': max(0, int_cuota),
                                    'valor_cuota_calculado': cuota_fija,
                                    'interes_pagado': 0,
                                    'interes_pendiente': max(0, int_cuota),
                                    'estado_cuota': 'Pendiente'
                                })

                        df_nuevo_cal = pd.DataFrame(nuevas_filas_cal)
                        st.session_state['df_calendario'] = pd.concat([st.session_state['df_calendario'], df_nuevo_cal], ignore_index=True)

                        st.success(f"✅ ¡Préstamo `{proximo_cred_id}` creado y sincronizado con éxito para **{cli_sel_cred}**! Se han actualizado automáticamente las tablas.")

        st.markdown("---")
        st.subheader("📥 Descargar Libro de Excel Actualizado")
        excel_bytes = exportar_excel_completo()
        st.download_button(
            label="📥 Descargar Excel Actualizado (.xlsx)",
            data=excel_bytes,
            file_name="proyecto_microcreditos_actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # =========================================================
    # 4. REGISTRAR PAGO
    # =========================================================
    elif opcion_menu == "📝 Registrar Pago":
        st.title("📝 Formulario de Registro de Pagos")
        st.markdown("Ingresa los datos del pago para actualizar las tablas de Excel y generar el comprobante PDF institucional.")
        st.markdown("---")

        col_f1, col_f2 = st.columns([1, 1])

        with col_f1:
            lista_cli = df_clientes['nombre'].dropna().unique()
            cliente_pago = st.selectbox("Selecciona el cliente:", lista_cli)

            info_cli_pago = df_clientes[df_clientes['nombre'] == cliente_pago].iloc[0]
            cliente_id_pago = info_cli_pago['cliente_id']
            telefono_cliente = info_cli_pago.get('telefono', '')

            creditos_cli = df_creditos[df_creditos['cliente_id'] == cliente_id_pago]

            if creditos_cli.empty:
                st.warning("Este cliente no tiene créditos activos registrados.")
            else:
                opciones_credito = [f"{row['credito_id']} - Capital: ${row['capital_inicial']:,.0f} ({row['modalidad']})" for _, row in creditos_cli.iterrows()]
                credito_sel_str = st.selectbox("Selecciona el Crédito:", opciones_credito)
                credito_id_pago = credito_sel_str.split(" - ")[0]

                fecha_pago = st.date_input("Fecha del Pago:", datetime.today())
                medio_pago = st.selectbox("Medio de Pago:", ["Transferencia", "Efectivo"])
                valor_pago = st.number_input("Valor Pagado (COP):", min_value=1000, value=1500000, step=5000, format="%d")
                concepto = st.selectbox("Concepto del Pago:", ["Intereses", "Abono a Capital", "Intereses y capital"])

        with col_f2:
            st.subheader("⚙️ Desglose del Pago")
            
            if not creditos_cli.empty:
                if concepto == "Intereses":
                    pago_interes = valor_pago
                    pago_capital = 0
                elif concepto == "Abono a Capital":
                    pago_interes = 0
                    pago_capital = valor_pago
                else:
                    pago_interes = st.number_input("Monto destinado a Intereses:", min_value=0, max_value=int(valor_pago), value=int(valor_pago*0.3))
                    pago_capital = valor_pago - pago_interes

                observaciones = st.text_input("Observaciones (opcional):", value="")

                existentes_ids = st.session_state['df_pagos']['pago_id'].dropna().tolist() if not st.session_state['df_pagos'].empty else []
                nums = [int(str(x).replace('PAG', '')) for x in existentes_ids if str(x).startswith('PAG') and str(x).replace('PAG', '').isdigit()]
                nuevo_num = max(nums) + 1 if nums else 1
                proximo_pago_id = f"PAG{nuevo_num:03d}"

                st.markdown("### Resumen a Procesar:")
                st.write(f"• **ID Nuevo Pago:** `{proximo_pago_id}`")
                st.write(f"• **Cliente:** {cliente_pago} (`{cliente_id_pago}`)")
                st.write(f"• **Crédito:** `{credito_id_pago}`")
                st.write(f"• **Fecha:** {fecha_pago}")
                st.write(f"• **Medio de Pago:** {medio_pago}")
                st.write(f"• **Abono a Intereses:** ${pago_interes:,.0f} COP")
                st.write(f"• **Abono a Capital:** ${pago_capital:,.0f} COP")

                if st.button("💾 Registrar Pago y Generar Comprobante", type="primary"):
                    nueva_fila_pago = {
                        'pago_id': proximo_pago_id,
                        'credito_id': credito_id_pago,
                        'cliente_id': cliente_id_pago,
                        'fecha_pago': pd.to_datetime(fecha_pago),
                        'medio_pago': medio_pago,
                        'valor_pago': valor_pago,
                        'pago_interes': pago_interes,
                        'pago_capital': pago_capital,
                        'numero_cuota': None,
                        'concepto': concepto,
                        'observaciones': observaciones if observaciones else None,
                        'interes_adicional': 0,
                        'valor_cuota_calculado': 0
                    }

                    st.session_state['df_pagos'] = pd.concat([st.session_state['df_pagos'], pd.DataFrame([nueva_fila_pago])], ignore_index=True)

                    idx_cred = df_creditos.index[df_creditos['credito_id'] == credito_id_pago].tolist()
                    if idx_cred:
                        ic = idx_cred[0]
                        saldo_cap_prev = df_creditos.loc[ic, 'saldo_capital'] if pd.notnull(df_creditos.loc[ic, 'saldo_capital']) else df_creditos.loc[ic, 'capital_inicial']
                        nuevo_saldo_cap = max(0, saldo_cap_prev - pago_capital)
                        st.session_state['df_creditos'].loc[ic, 'saldo_capital'] = nuevo_saldo_cap
                        if nuevo_saldo_cap == 0:
                            st.session_state['df_creditos'].loc[ic, 'estado_credito'] = "Finalizado"

                    if not st.session_state['df_calendario'].empty and pago_interes > 0:
                        remanente_int = pago_interes
                        filas_cal = st.session_state['df_calendario'][st.session_state['df_calendario']['credito_id'] == credito_id_pago].index.tolist()
                        for idx_c in filas_cal:
                            if remanente_int <= 0:
                                break
                            int_pend = st.session_state['df_calendario'].loc[idx_c, 'interes_pendiente']
                            if int_pend > 0:
                                abono = min(remanente_int, int_pend)
                                st.session_state['df_calendario'].loc[idx_c, 'interes_pagado'] += abono
                                st.session_state['df_calendario'].loc[idx_c, 'interes_pendiente'] -= abono
                                remanente_int -= abono

                    nuevo_cap_pend = 0
                    nuevo_int_pend = 0
                    nueva_deuda_total = 0
                    nuevo_estado = "Al día"

                    idx_ec = st.session_state['df_estado_cartera'].index[st.session_state['df_estado_cartera']['credito_id'] == credito_id_pago].tolist()
                    if idx_ec:
                        ie = idx_ec[0]
                        st.session_state['df_estado_cartera'].loc[ie, 'capital_pagado'] += pago_capital
                        nuevo_cap_pend = max(0, st.session_state['df_estado_cartera'].loc[ie, 'capital_pendiente'] - pago_capital)
                        st.session_state['df_estado_cartera'].loc[ie, 'capital_pendiente'] = nuevo_cap_pend

                        nuevo_int_pend = max(0, st.session_state['df_estado_cartera'].loc[ie, 'interes_pendiente'] - pago_interes)
                        st.session_state['df_estado_cartera'].loc[ie, 'interes_pendiente'] = nuevo_int_pend

                        nuevo_vencido = max(0, st.session_state['df_estado_cartera'].loc[ie, 'deuda_vencida'] - (pago_interes + pago_capital))
                        st.session_state['df_estado_cartera'].loc[ie, 'deuda_vencida'] = nuevo_vencido

                        nueva_deuda_total = nuevo_cap_pend + nuevo_int_pend
                        st.session_state['df_estado_cartera'].loc[ie, 'deuda_total_pendiente'] = nueva_deuda_total

                        if nueva_deuda_total == 0:
                            nuevo_estado = "Paz y salvo"
                        elif nuevo_vencido == 0:
                            nuevo_estado = "Al día"
                        else:
                            nuevo_estado = "En mora"
                        st.session_state['df_estado_cartera'].loc[ie, 'estado'] = nuevo_estado

                    st.success(f"✅ ¡Pago **{proximo_pago_id}** registrado exitosamente!")

                    datos_pago_pdf = {
                        'pago_id': proximo_pago_id,
                        'cliente_id': cliente_id_pago,
                        'cliente_nombre': cliente_pago,
                        'credito_id': credito_id_pago,
                        'fecha_pago': fecha_pago.strftime('%Y-%m-%d'),
                        'medio_pago': medio_pago,
                        'concepto': concepto,
                        'pago_interes': pago_interes,
                        'pago_capital': pago_capital,
                        'valor_pago': valor_pago,
                        'nuevo_cap_pend': nuevo_cap_pend,
                        'nuevo_int_pend': nuevo_int_pend,
                        'nueva_deuda_total': nueva_deuda_total,
                        'nuevo_estado': nuevo_estado,
                        'observaciones': observaciones
                    }

                    pdf_bytes = generar_pdf_comprobante(datos_pago_pdf)

                    st.subheader("📄 Comprobante Digital Generado")
                    col_pdf, col_wa = st.columns(2)

                    with col_pdf:
                        st.download_button(
                            label="📥 Descargar Comprobante PDF",
                            data=pdf_bytes,
                            file_name=f"Comprobante_{proximo_pago_id}_{cliente_pago.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

                    with col_wa:
                        nombre_w = cliente_pago
                        msg_wa = (
                            f"Hola {nombre_w} 😊 ¿Cómo estás?\n\n"
                            f"Te escribo para contarte que ya recibí tu pago de ${valor_pago:,.0f}. 🙌\n\n"
                            f"¡Muchas gracias por ponerte al día y por cumplir con tu pago! 😊"
                        )
                        num_tel = "".join(filter(str.isdigit, str(telefono_cliente)))
                        if len(num_tel) == 10 and not num_tel.startswith("57"):
                            num_tel = "57" + num_tel
                        msg_enc = urllib.parse.quote(msg_wa)
                        
                        st.link_button("💬 Enviar Confirmación por WhatsApp", f"https://wa.me/{num_tel}?text={msg_enc}", use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Descargar Libro de Excel Actualizado")
        excel_bytes = exportar_excel_completo()
        st.download_button(
            label="📥 Descargar Excel Actualizado (.xlsx)",
            data=excel_bytes,
            file_name="proyecto_microcreditos_actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # =========================================================
    # 5. GESTIÓN DE COBRO
    # =========================================================
    elif opcion_menu == "⚖️ Gestión de Cobro":
        st.title("⚖️ Centro de Gestión de Cobro y Alertas")
        st.markdown("Monitorea los créditos que se encuentran en mora y envía recordatorios profesionales de pago de forma inmediata.")
        st.markdown("---")

        if not df_estado_cartera.empty and 'estado' in df_estado_cartera.columns:
            df_mora_detalle = df_estado_cartera.merge(df_clientes[['cliente_id', 'nombre', 'telefono']], on='cliente_id', how='left')
            df_mora_activa = df_mora_detalle[df_mora_detalle['estado'].astype(str).str.contains('mora', case=False, na=False)]

            if not df_mora_activa.empty:
                st.markdown(f"### 📋 Créditos pendientes de pago ({len(df_mora_activa)})")
                st.markdown("<br>", unsafe_allow_html=True)

                for _, row in df_mora_activa.iterrows():
                    nombre_cli = row.get('nombre', 'Desconocido')
                    val_pend = row.get('deuda_vencida', 0)
                    if val_pend == 0:
                        val_pend = row.get('deuda_total_pendiente', 0)
                    cred_id = row.get('credito_id', 'N/A')
                    tel_cli = row.get('telefono', '')
                    estado_actual = row.get('estado', 'En mora')

                    with st.container():
                        col_info, col_btn = st.columns([3, 1])
                        with col_info:
                            st.warning(
                                f"👤 **Cliente:** {nombre_cli}  \n"
                                f"💳 **Crédito N°:** `{cred_id}` | 📌 **Estado:** {estado_actual}  \n"
                                f"💰 **Valor Pendiente / Vencido:** **${val_pend:,.0f} COP**"
                            )
                        with col_btn:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if tel_cli:
                                num_tel = "".join(filter(str.isdigit, str(tel_cli)))
                                if len(num_tel) == 10 and not num_tel.startswith("57"):
                                    num_tel = "57" + num_tel
                                
                                nombre_w = nombre_cli
                                msg_cobro = (
                                    f"Hola {nombre_w} 😊 ¿Cómo estás?\n\n"
                                    f"{nombre_w}, te quería recordar que tienes pendiente un pago de ${val_pend:,.0f}. Cuando puedas, me ayudas con este pendiente para dejar todo al día 🙌\n\n"
                                    f"Cuando lo hagas, me compartes por aquí el comprobante. ¡Muchas gracias! 😊"
                                )
                                msg_enc = urllib.parse.quote(msg_cobro)
                                st.link_button("💬 Enviar Cobro WA", f"https://wa.me/{num_tel}?text={msg_enc}", use_container_width=True)
                            else:
                                st.caption("📞 Teléfono no disponible")
                        st.markdown("---")
            else:
                st.success("🟢 **¡Excelente noticia!** No hay créditos en mora registrados actualmente en el sistema.")
        else:
            st.info("ℹ️ No se encontró la hoja `Estado_Cartera` o la columna `estado` en el archivo cargado.")

    # =========================================================
    # 6. SIMULADOR DE CRÉDITOS
    # =========================================================
    elif opcion_menu == "🧮 Simulador de Créditos":
        st.title("🧮 Simulador de Créditos")
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
    # 7. ASISTENTE INTELIGENTE (IA)
    # =========================================================
    elif opcion_menu == "🤖 Asistente IA":
        st.title("🤖 Asistente Inteligente del Fondo")
        st.markdown("Hazle preguntas en lenguaje natural sobre tus clientes, créditos, pagos o estado de cartera.")
        st.markdown("---")

        groq_api_key = st.text_input("Ingresa tu Clave de API de Groq:", type="password")

        pregunta_usuario = st.text_area(
            "¿Qué te gustaría saber de la cartera?",
            placeholder="Ej: ¿Qué clientes tienen créditos activos y cuál es el total prestado?"
        )

        if st.button("Consultar al Asistente", type="primary"):
            if not groq_api_key:
                st.error("⚠️ Por favor ingresa una clave de API válida para continuar.")
            elif not pregunta_usuario:
                st.error("⚠️ Escribe una pregunta para el asistente.")
            else:
                try:
                    with st.spinner("Analizando datos y generando respuesta..."):
                        client = Groq(api_key=groq_api_key)

                        resumen_clientes = st.session_state['df_clientes'].to_string() if not st.session_state['df_clientes'].empty else "Sin clientes"
                        resumen_estado = st.session_state['df_estado_cartera'].to_string() if not st.session_state['df_estado_cartera'].empty else "Sin cartera"

                        prompt_sistema = f"""
                        Eres el analista financiero experto de 'Entre Amigos Capital', un fondo familiar de microcréditos.
                        Tienes acceso a los siguientes datos actuales de la base de datos:
                        
                        --- CLIENTES ---
                        {resumen_clientes}
                        
                        --- ESTADO DE CARTERA ---
                        {resumen_estado}
                        
                        Responde a la pregunta del usuario de manera clara, profesional, concisa y basada estrictamente en los datos provistos.
                        """

                        chat_completion = client.chat.completions.create(
                            messages=[
                                {"role": "system", "content": prompt_sistema},
                                {"role": "user", "content": pregunta_usuario}
                            ],
                            model="llama3-70b-8192",
                        )

                        respuesta_ia = chat_completion.choices[0].message.content
                        st.markdown("### 💡 Respuesta del Asistente:")
                        st.info(respuesta_ia)

                except Exception as e:
                    st.error(f"Error al conectar con la inteligencia artificial: {e}")

    # =========================================================
    # 8. SOBRE NOSOTROS & POLÍTICAS
    # =========================================================
    elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
        st.title("💎 Sobre Nuestros Microcréditos")
        st.markdown("---")
        st.subheader("💡 Nuestra Propuesta de Valor")
        st.write("Acceso a microcréditos ágiles para familiares y amigos a tasas justas (3% mensual), fomentando la economía colaborativa.")
