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

EXCEL_FILE_DEFAULT = "proyecto microcréditos copia 3.xlsx"

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
# CARGA DE DATOS DESDE EL EXCEL
# ---------------------------------------------------------
st.sidebar.title("💎 Entre Amigos Capital")
st.sidebar.caption("Fondo de Inversión y Microcréditos Familiares")

with st.sidebar.expander("📌 Nuestra Misión", expanded=False
