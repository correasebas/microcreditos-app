import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import os
from dateutil.relativedelta import relativedelta

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
# CONFIGURACIÓN DE PÁGINA & ESTILO FINTECH
# ---------------------------------------------------------
st.set_page_config(
    page_title="Entre Amigos Capital - Fondo Familiar",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

EXCEL_FILE_DEFAULT = "proyecto_microcreditos_actualizado 8.xlsx"

# ---------------------------------------------------------
# CLASE PDF INSTITUCIONAL
# ---------------------------------------------------------
try:
    from fpdf import FPDF
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
        pdf.cell(0, 4, "Gracias por mantener tu crédito al día. Soporte oficial de recaudo.", ln=True, align="C")
        return bytes(pdf.output())
except:
    def generar_pdf_comprobante(pago_info):
        return b""

# ---------------------------------------------------------
# MOTOR DE CÁLCULO DINÁMICO Y AUTOMÁTICO DE CARTERA
# ---------------------------------------------------------
def calcular_cartera_dinamica(df_creditos, df_pagos):
    if df_creditos.empty:
        return pd.DataFrame()
    
    hoy = datetime.today()
    registros = []
    
    for _, cred in df_creditos.iterrows():
        c_id = cred['credito_id']
        cli_id = cred['cliente_id']
        cap_ini = float(cred.get('capital_inicial', 0))
        tasa = float(cred.get('tasa_mensual', 0.03))
        f_desembolso = pd.to_datetime(cred.get('fecha_desembolso'))
        
        # Pagos del crédito
        pagos_cred = df_pagos[df_pagos['credito_id'] == c_id] if not df_pagos.empty else pd.DataFrame()
        cap_pagado = float(pagos_cred['pago_capital'].sum()) if not pagos_cred.empty and 'pago_capital' in pagos_cred.columns else 0.0
        cap_pend = max(0.0, cap_ini - cap_pagado)
        
        total_interes_pagado = float(pagos_cred['pago_interes'].sum()) if not pagos_cred.empty and 'pago_interes' in pagos_cred.columns else 0.0
        
        # Calcular cuántos meses han transcurrido desde el desembolso hasta hoy
        if pd.notnull(f_desembolso):
            r = relativedelta(hoy, f_desembolso)
            meses_transcurridos = max(0, r.years * 12 + r.months)
            if hoy.day >= f_desembolso.day and meses_transcurridos == 0:
                meses_transcurridos = 1
            elif hoy.day >= f_desembolso.day:
                meses_transcurridos += 1
        else:
            meses_transcurridos = 0
            
        interes_generado_total = meses_transcurridos * (cap_ini * tasa)
        interes_pend = max(0.0, interes_generado_total - total_interes_pagado)
        
        deuda_total = cap_pend + interes_pend
        estado = "En mora" if interes_pend > 0 else "Al día"
        deuda_vencida = interes_pend if estado == "En mora" else 0.0

        registros.append({
            'credito_id': c_id,
            'cliente_id': cli_id,
            'tipo_interes': cred.get('modalidad', 'INTERES_MENSUAL'),
            'capital_inicial': cap_ini,
            'capital_pagado': cap_pagado,
            'capital_pendiente': cap_pend,
            'interes_pendiente': interes_pend,
            'deuda_total_pendiente': deuda_total,
            'deuda_vencida': deuda_vencida,
            'estado': estado
        })
        
    return pd.DataFrame(registros)

# ---------------------------------------------------------
# CARGA DE DATOS DESDE EL EXCEL
# ---------------------------------------------------------
st.sidebar.title("💎 Entre Amigos Capital")
st.sidebar.caption("Fondo de Inversión y Microcréditos Familiares")

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Base de Datos Excel")
uploaded_file = st.sidebar.file_uploader(
    "Carga tu archivo de Excel actualizado:", 
    type=["xlsx"],
    help="Sube tu archivo .xlsx para sincronizar la aplicación."
)

file_to_load = uploaded_file if uploaded_file is not None else EXCEL_FILE_DEFAULT

def cargar_datos_completos(file_source):
    try:
        xls = pd.ExcelFile(file_source)
        df_clientes = pd.read_excel(xls, sheet_name='Clientes')
        df_creditos = pd.read_excel(xls, sheet_name='Creditos')
        df_pagos = pd.read_excel(xls, sheet_name='Pagos')
        df_cal = pd.read_excel(xls, sheet_name='Calendario_Intereses') if 'Calendario_Intereses' in xls.sheet_names else pd.DataFrame()
        
        # Calcular cartera automáticamente con fecha en tiempo real
        df_est_dinamico = calcular_cartera_dinamica(df_creditos, df_pagos)
        return df_clientes, df_creditos, df_pagos, df_est_dinamico, df_cal
    except Exception as e:
        st.error(f"Error al cargar el archivo de Excel: {e}")
        return None, None, None, None, None

if 'current_loaded_file' not in st.session_state or st.session_state['current_loaded_file'] != file_to_load or uploaded_file is not None:
    df_c, df_cr, df_p, df_est, df_cal = cargar_datos_completos(file_to_load)
    st.session_state['df_clientes'] = df_c if df_c is not None else pd.DataFrame()
    st.session_state['df_creditos'] = df_cr if df_cr is not None else pd.DataFrame()
    st.session_state['df_pagos'] = df_p if df_p is not None else pd.DataFrame()
    st.session_state['df_estado_cartera'] = df_est if df_est is not None else pd.DataFrame()
    st.session_state['df_calendario'] = df_cal if df_cal is not None else pd.DataFrame()
    st.session_state['current_loaded_file'] = file_to_load

df_clientes = st.session_state['df_clientes']
df_creditos = st.session_state['df_creditos']
df_pagos = st.session_state['df_pagos']
df_estado_cartera = st.session_state['df_estado_cartera']

st.sidebar.markdown("---")
opcion_menu = st.sidebar.radio(
    "Selecciona una sección:",
    ["📊 Dashboard General", "👤 Ficha por Cliente", "➕ Nuevos Registros", "📝 Registrar Pago", "⚖️ Gestión de Cobro", "🧮 Simulador de Créditos", "🤖 Asistente IA (Groq)", "ℹ️ Sobre Nosotros & Políticas"]
)

def obtener_resumen_general():
    df_ec = st.session_state.get('df_estado_cartera', pd.DataFrame())
    df_cr = st.session_state.get('df_creditos', pd.DataFrame())
    if df_ec.empty:
        return pd.DataFrame()
    
    cap_prestado = df_cr['capital_inicial'].sum() if not df_cr.empty else df_ec['capital_inicial'].sum()
    cap_pagado = df_ec['capital_pagado'].sum() if 'capital_pagado' in df_ec.columns else 0
    cap_pendiente = df_ec['capital_pendiente'].sum() if 'capital_pendiente' in df_ec.columns else 0
    int_pendiente = df_ec['interes_pendiente'].sum() if 'interes_pendiente' in df_ec.columns else 0
    deuda_total = df_ec['deuda_total_pendiente'].sum() if 'deuda_total_pendiente' in df_ec.columns else 0
    deuda_vencida = df_ec['deuda_vencida'].sum() if 'deuda_vencida' in df_ec.columns else 0
    
    creditos_activos = len(df_ec[df_ec['deuda_total_pendiente'] > 1])
    creditos_mora = len(df_ec[df_ec['estado'].astype(str).str.contains('mora', case=False, na=False)])
    creditos_aldia = len(df_ec[df_ec['estado'].astype(str).str.contains('día|salvo', case=False, na=False)])

    return pd.DataFrame({
        'Indicador': ['Capital total prestado', 'Capital pagado', 'Capital pendiente', 'Intereses pendientes', 'Deuda total pendiente', 'Deuda vencida', 'Créditos activos', 'Créditos en mora', 'Créditos al día'],
        'Resultado': [cap_prestado, cap_pagado, cap_pendiente, int_pendiente, deuda_total, deuda_vencida, creditos_activos, creditos_mora, creditos_aldia]
    })

def exportar_excel_completo():
    output = io.BytesIO()
    df_resumen_actualizado = obtener_resumen_general()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not st.session_state['df_clientes'].empty: st.session_state['df_clientes'].to_excel(writer, sheet_name='Clientes', index=False)
        if not st.session_state['df_creditos'].empty: st.session_state['df_creditos'].to_excel(writer, sheet_name='Creditos', index=False)
        if not st.session_state['df_pagos'].empty: st.session_state['df_pagos'].to_excel(writer, sheet_name='Pagos', index=False)
        if not st.session_state['df_estado_cartera'].empty: st.session_state['df_estado_cartera'].to_excel(writer, sheet_name='Estado_Cartera', index=False)
        if not st.session_state['df_calendario'].empty: st.session_state['df_calendario'].to_excel(writer, sheet_name='Calendario_Intereses', index=False)
        if not df_resumen_actualizado.empty: df_resumen_actualizado.to_excel(writer, sheet_name='Resumen_Cartera', index=False)
    return output.getvalue()

# =========================================================
# 1. DASHBOARD GENERAL
# =========================================================
if opcion_menu == "📊 Dashboard General":
    st.title("📊 Control General de Cartera (Tiempo Real)")
    st.markdown("---")
    df_res = obtener_resumen_general()
    dict_res = dict(zip(df_res['Indicador'], df_res['Resultado'])) if not df_res.empty else {}

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Capital Total Prestado", f"${dict_res.get('Capital total prestado', 0):,.0f} COP")
    c2.metric("Capital Pagado", f"${dict_res.get('Capital pagado', 0):,.0f} COP")
    c3.metric("Capital Pendiente", f"${dict_res.get('Capital pendiente', 0):,.0f} COP")
    c4.metric("Deuda Total Pendiente", f"${dict_res.get('Deuda total pendiente', 0):,.0f} COP")

    st.markdown("<br>", unsafe_allow_html=True)
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Intereses Pendientes", f"${dict_res.get('Intereses pendientes', 0):,.0f} COP")
    c6.metric("Deuda Vencida (Mora)", f"${dict_res.get('Deuda vencida', 0):,.0f} COP",
