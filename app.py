import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import urllib.parse
from fpdf import FPDF
import os

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

# ---------------------------------------------------------
# CARGA DE DATOS RIGUROSA RESPETANDO EL EXCEL ORIGINAL
# ---------------------------------------------------------
def cargar_datos_excel(file_source):
    try:
        xls = pd.ExcelFile(file_source)
        df_clientes = pd.read_excel(xls, sheet_name='Clientes')
        df_creditos = pd.read_excel(xls, sheet_name='Creditos')
        df_pagos = pd.read_excel(xls, sheet_name='Pagos')
        df_est = pd.read_excel(xls, sheet_name='Estado_Cartera') if 'Estado_Cartera' in xls.sheet_names else pd.DataFrame()
        df_cal = pd.read_excel(xls, sheet_name='Calendario_Intereses') if 'Calendario_Intereses' in xls.sheet_names else pd.DataFrame()
        return df_clientes, df_creditos, df_pagos, df_est, df_cal
    except Exception as e:
        st.error(f"Error al cargar el archivo de Excel: {e}")
        return None, None, None, None, None

# ---------------------------------------------------------
# CARGA INICIAL DESDE EL EXCEL
# ---------------------------------------------------------
st.sidebar.title("💎 Entre Amigos Capital")
st.sidebar.caption("Fondo de Inversión y Microcréditos Familiares")

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Base de Datos Excel")
uploaded_file = st.sidebar.file_uploader(
    "Carga tu archivo de Excel actualizado:", 
    type=["xlsx"],
    help="Si no subes un archivo, se cargará el archivo base por defecto."
)

file_to_load = uploaded_file if uploaded_file is not None else EXCEL_FILE_DEFAULT

if 'current_loaded_file' not in st.session_state or st.session_state['current_loaded_file'] != file_to_load:
    df_c, df_cr, df_p, df_est, df_cal = cargar_datos_excel(file_to_load)
    st.session_state['df_clientes'] = df_c if df_c is not None else pd.DataFrame()
    st.session_state['df_creditos'] = df_cr if df_cr is not None else pd.DataFrame()
    st.session_state['df_pagos'] = df_p if df_p is not None else pd.DataFrame()
    st.session_state['df_estado_cartera'] = df_est if df_est is not None else pd.DataFrame()
    st.session_state['df_calendario'] = df_cal if df_cal is not None else pd.DataFrame()
    st.session_state['current_loaded_file'] = file_to_load

for key, default_val in [('df_clientes', pd.DataFrame()), ('df_creditos', pd.DataFrame()), ('df_pagos', pd.DataFrame()), ('df_estado_cartera', pd.DataFrame())]:
    if key not in st.session_state:
        st.session_state[key] = default_val

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
    st.title("📊 Control General de Cartera")
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
    c6.metric("Deuda Vencida (Mora)", f"${dict_res.get('Deuda vencida', 0):,.0f} COP", delta=f"-{dict_res.get('Créditos en mora', 0)} créditos", delta_color="inverse")
    c7.metric("Créditos Al Día", f"{dict_res.get('Créditos al día', 0)}", delta="Puntuales")
    c8.metric("Total Créditos Activos", f"{dict_res.get('Créditos activos', 0)}")

    st.markdown("---")
    col_left, col_right = st.columns([1, 1])
    with col_left:
        st.subheader("Estado de Créditos Activos")
        df_estado = pd.DataFrame({"Estado": ["Al Día / Paz y Salvo", "En Mora"], "Cantidad": [dict_res.get('Créditos al día', 0), dict_res.get('Créditos en mora', 0)]})
        fig_pie = px.pie(df_estado, names="Estado", values="Cantidad", hole=0.4, color="Estado", color_discrete_map={"Al Día / Paz y Salvo": "#00D26A", "En Mora": "#E74C3C"})
        fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#E6EDF3", legend=dict(font=dict(color="#E6EDF3")))
        st.plotly_chart(fig_pie, use_container_width=True)
    with col_right:
        st.subheader("Estado Oficial de Cartera")
        if not st.session_state['df_estado_cartera'].empty:
            st.dataframe(st.session_state['df_estado_cartera'], use_container_width=True)

# =========================================================
# 2. FICHA POR CLIENTE
# =========================================================
elif opcion_menu == "👤 Ficha por Cliente":
    st.title("👤 Ficha de Cliente e Historial de Deuda")
    st.markdown("---")
    if not df_clientes.empty:
        cliente_sel = st.selectbox("Selecciona un cliente:", df_clientes['nombre'].dropna().unique())
        if cliente_sel:
            info_cliente = df_clientes[df_clientes['nombre'] == cliente_sel].iloc[0]
            cliente_id = info_cliente['cliente_id']
            creditos_cliente = st.session_state['df_creditos'][st.session_state['df_creditos']['cliente_id'] == cliente_id]
            pagos_cliente = st.session_state['df_pagos'][st.session_state['df_pagos']['cliente_id'] == cliente_id]
            cartera_cliente = st.session_state['df_estado_cartera'][st.session_state['df_estado_cartera']['cliente_id'] == cliente_id] if not st.session_state['df_estado_cartera'].empty else pd.DataFrame()

            cap_pend = cartera_cliente['capital_pendiente'].sum() if not cartera_cliente.empty else 0
            int_pend = cartera_cliente['interes_pendiente'].sum() if not cartera_cliente.empty else 0
            deuda_vencida = cartera_cliente['deuda_vencida'].sum() if not cartera_cliente.empty else 0
            deuda_total = cartera_cliente['deuda_total_pendiente'].sum() if not cartera_cliente.empty else 0
            tiene_mora = any(cartera_cliente['estado'].astype(str).str.contains('mora', case=False, na=False)) if not cartera_cliente.empty else False

            if tiene_mora or deuda_vencida > 1:
                st.error(f"⚠️ **Alerta Individual - En Mora:** Este cliente presenta cuotas vencidas por un valor de **${deuda_vencida:,.0f} COP** (Deuda total:${deuda_total:,.0f} COP).")
            elif deuda_total <= 1:
                st.success("🟢 **Paz y Salvo:** El cliente no presenta saldos pendientes.")
            else:
                st.info("🟢 **Al Día:** El cliente cuenta con sus cuotas e intereses al día.")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Capital Pendiente", f"${cap_pend:,.0f} COP")
            m2.metric("Intereses Pendientes", f"${int_pend:,.0f} COP")
            m3.metric("Deuda Vencida", f"${deuda_vencida:,.0f} COP")
            m4.metric("Deuda Total Pendiente", f"${deuda_total:,.0f} COP")

            st.markdown("---")
            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("Información Personal")
                st.write(f"**ID Cliente:** {cliente_id}")
                st.write(f"**Teléfono:** {info_cliente.get('telefono', 'N/A')}")
            with col_b:
                st.subheader("Relación y Registro")
                st.write(f"**Fecha Registro:** {info_cliente.get('fecha_registro', 'N/A')}")
                st.write(f"**Parentesco:** {info_cliente.get('parentesco', 'N/A')}")

            st.markdown("---")
            st.subheader("Créditos del Cliente")
            if not creditos_cliente.empty: st.dataframe(creditos_cliente, use_container_width=True)
            st.subheader("Historial de Pagos")
            if not pagos_cliente.empty: st.dataframe(pagos_cliente, use_container_width=True)

# =========================================================
# 3. NUEVOS REGISTROS
# =========================================================
elif opcion_menu == "➕ Nuevos Registros":
    st.title("➕ Módulo Integrado de Nuevos Registros")
    st.markdown("---")
    tab_cli, tab_cred = st.tabs(["👤 Registrar Nuevo Cliente", "💳 Otorgar Nuevo Préstamo"])

    with tab_cli:
        with st.form("form_nuevo_cliente"):
            col_nc1, col_nc2 = st.columns(2)
            with col_nc1:
                nombre_nuevo = st.text_input("Nombre Completo:")
                telefono_nuevo = st.text_input("Teléfono / WhatsApp:")
            with col_nc2:
                parentesco_nuevo = st.selectbox("Parentesco:", ["Familiar", "Amigo", "Conocido", "Socio"])
                estado_cli_nuevo = st.selectbox("Estado del Cliente:", ["Activo", "Inactivo"])
            if st.form_submit_button("💾 Guardar Nuevo Cliente", type="primary"):
                if nombre_nuevo:
                    ids = st.session_state['df_clientes']['cliente_id'].dropna().tolist() if not st.session_state['df_clientes'].empty else []
                    nums = [int(str(x).replace('CLI', '')) for x in ids if str(x).startswith('CLI') and str(x).replace('CLI', '').isdigit()]
                    nuevo_id = f"CLI{(max(nums) + 1 if nums else 1):03d}"
                    fila = {'cliente_id': nuevo_id, 'nombre': nombre_nuevo, 'telefono': telefono_nuevo, 'fecha_registro': datetime.today().strftime('%Y-%m-%d'), 'estado_cliente': estado_cli_nuevo, 'parentesco': parentesco_nuevo}
                    st.session_state['df_clientes'] = pd.concat([st.session_state['df_clientes'], pd.DataFrame([fila])], ignore_index=True)
                    st.success(f"✅ ¡Cliente **{nombre_nuevo}** registrado como `{nuevo_id}`!")
                else:
                    st.error("⚠️ El nombre es obligatorio.")

    with tab_cred:
        df_cli_act = st.session_state.get('df_clientes', pd.DataFrame())
        if df_cli_act.empty:
            st.warning("⚠️ Registra un cliente primero.")
        else:
            with st.form("form_nuevo_prestamo"):
                cli_sel_cred = st.selectbox("Cliente Beneficiario:", df_cli_act['nombre'].dropna().unique())
                id_cli = df_cli_act[df_cli_act['nombre'] == cli_sel_cred].iloc[0]['cliente_id']
                col_np1, col_np2 = st.columns(2)
                with col_np1:
                    cap_ini = st.number_input("Capital Inicial (COP):", min_value=100000, value=1000000, step=50000, format="%d")
                    tasa = st.number_input("Tasa Mensual (%):", min_value=0.0, value=3.0, step=0.5) / 100.0
                    plazo = st.number_input("Plazo en Meses:", min_value=1, value=6, step=1)
                with col_np2:
                    fecha_des = st.date_input("Fecha Desembolso:", datetime.today())
                    modalidad = st.selectbox("Modalidad:", ["INTERES_MENSUAL", "CUOTA_FIJA"])
                if st.form_submit_button("🚀 Otorgar Préstamo", type="primary"):
                    ids_cr = st.session_state['df_creditos']['credito_id'].dropna().tolist() if not st.session_state['df_creditos'].empty else []
                    nums_cr = [int(str(x).replace('CR', '')) for x in ids_cr if str(x).startswith('CR') and str(x).replace('CR', '').isdigit()]
                    nuevo_cr_id = f"CR{(max(nums_cr) + 1 if nums_cr else 1):03d}"
                    fila_cr = {
                        'credito_id': nuevo_cr_id, 'cliente_id': id_cli, 'fecha_desembolso': pd.to_datetime(fecha_des),
                        'modalidad': modalidad, 'capital_inicial': cap_ini, 'tasa_mensual': tasa, 'plazo_meses': plazo,
                        'dia_pago': pd.to_datetime(fecha_des).day, 'numero_cuotas': plazo, 'valor_cuota': cap_ini * tasa if modalidad == "INTERES_MENSUAL" else (cap_ini * tasa) / (1 - (1 + tasa)**(-plazo)),
                        'saldo_capital': cap_ini, 'estado_credito': 'Al día'
                    }
                    st.session_state['df_creditos'] = pd.concat([st.session_state['df_creditos'], pd.DataFrame([fila_cr])], ignore_index=True)
                    st.success(f"✅ ¡Crédito `{nuevo_cr_id}` creado con éxito!")

    st.markdown("---")
    st.download_button("📥 Descargar Excel Actualizado (.xlsx)", data=exportar_excel_completo(), file_name="proyecto_microcreditos_actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================================================
# 4. REGISTRAR PAGO
# =========================================================
elif opcion_menu == "📝 Registrar Pago":
    st.title("📝 Formulario de Registro de Pagos")
    st.markdown("---")
    col_f1, col_f2 = st.columns([1, 1])
    with col_f1:
        cliente_pago = st.selectbox("Selecciona el cliente:", df_clientes['nombre'].dropna().unique())
        info_cli_pago = df_clientes[df_clientes['nombre'] == cliente_pago].iloc[0]
        creditos_cli = st.session_state['df_creditos'][st.session_state['df_creditos']['cliente_id'] == info_cli_pago['cliente_id']]
        if creditos_cli.empty:
            st.warning("Este cliente no tiene créditos activos.")
        else:
            cred_str = st.selectbox("Crédito:", [f"{r['credito_id']} - Capital: ${r['capital_inicial']:,.0f} ({r['modalidad']})" for _, r in creditos_cli.iterrows()])
            credito_id = cred_str.split(" - ")[0]
            fecha_pago = st.date_input("Fecha:", datetime.today())
            medio = st.selectbox("Medio:", ["Transferencia", "Efectivo"])
            valor = st.number_input("Valor Pagado (COP):", min_value=1000, value=50000, step=1000, format="%d")
            concepto = st.selectbox("Concepto:", ["Intereses", "Abono a Capital", "Intereses y capital"])
    with col_f2:
        if not creditos_cli.empty:
            p_int = valor if concepto == "Intereses" else (valor * 0.3 if concepto == "Intereses y capital" else 0)
            p_cap = valor if concepto == "Abono a Capital" else (valor - p_int if concepto == "Intereses y capital" else 0)
            obs = st.text_input("Observaciones:")
            
            if st.button("💾 Registrar Pago y Generar Comprobante", type="primary"):
                ids_p = st.session_state['df_pagos']['pago_id'].dropna().tolist() if not st.session_state['df_pagos'].empty else []
                nums_p = [int(str(x).replace('PAG', '')) for x in ids_p if str(x).startswith('PAG') and str(x).replace('PAG', '').isdigit()]
                pago_id = f"PAG{(max(nums_p) + 1 if nums_p else 1):03d}"
                
                nueva_p = {'pago_id': pago_id, 'credito_id': credito_id, 'cliente_id': info_cli_pago['cliente_id'], 'fecha_pago': pd.to_datetime(fecha_pago), 'medio_pago': medio, 'valor_pago': valor, 'pago_interes': p_int, 'pago_capital': p_cap, 'concepto': concepto, 'observaciones': obs}
                st.session_state['df_pagos'] = pd.concat([st.session_state['df_pagos'], pd.DataFrame([nueva_p])], ignore_index=True)
                
                row_act = st.session_state['df_estado_cartera'][st.session_state['df_estado_cartera']['credito_id'] == credito_id].iloc[0] if not st.session_state['df_estado_cartera'].empty else {}
                st.success("✅ ¡Pago registrado con éxito!")
                
                pdf_bytes = generar_pdf_comprobante({
                    'pago_id': pago_id, 'cliente_id': info_cli_pago['cliente_id'], 'cliente_nombre': cliente_pago, 'credito_id': credito_id,
                    'fecha_pago': fecha_pago.strftime('%Y-%m-%d'), 'medio_pago': medio, 'concepto': concepto, 'pago_interes': p_int,
                    'pago_capital': p_cap, 'valor_pago': valor, 'nuevo_cap_pend': row_act.get('capital_pendiente', 0), 'nuevo_int_pend': row_act.get('interes_pendiente', 0),
                    'nueva_deuda_total': row_act.get('deuda_total_pendiente', 0), 'nuevo_estado': row_act.get('estado', 'Al día'), 'observaciones': obs
                })
                st.download_button("📥 Descargar Comprobante PDF", data=pdf_bytes, file_name=f"Comprobante_{pago_id}.pdf", mime="application/pdf", use_container_width=True)

# =========================================================
# 5. GESTIÓN DE COBRO
# =========================================================
elif opcion_menu == "⚖️ Gestión de Cobro":
    st.title("⚖️ Centro de Gestión de Cobro")
    st.markdown("---")
    df_ec_act = st.session_state.get('df_estado_cartera', pd.DataFrame())
    if not df_ec_act.empty:
        df_mora = df_ec_act.merge(df_clientes[['cliente_id', 'nombre', 'telefono']], on='cliente_id', how='left')
        df_mora = df_mora[df_mora['estado'].astype(str).str.contains('mora', case=False, na=False)]
        if not df_mora.empty:
            for _, r in df_mora.iterrows():
                val = r.get('deuda_vencida', 0) if r.get('deuda_vencida', 0) > 0 else r.get('deuda_total_pendiente', 0)
                st.warning(f"👤 **{r.get('nombre')}** | Crédito `{r.get('credito_id')}` | Pendiente: **${val:,.0f} COP**")
        else:
            st.success("🟢 ¡No hay créditos en mora actualmente!")

# =========================================================
# 6. SIMULADOR DE CRÉDITOS
# =========================================================
elif opcion_menu == "🧮 Simulador de Créditos":
    st.title("🧮 Simulador de Créditos")
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        monto = st.number_input("Monto:", min_value=100000, value=1000000, step=50000)
        plazo = st.slider("Meses:", 1, 24, 6)
    with c2:
        cuota = (monto * 0.03) / (1 - (1 + 0.03)**(-plazo))
        st.metric("Cuota Fija Mensual Estimada", f"${cuota:,.0f} COP")

# =========================================================
# 7. ASISTENTE IA (GROQ)
# =========================================================
elif opcion_menu == "🤖 Asistente IA (Groq)":
    st.title("🤖 Asistente Inteligente Groq")
    st.markdown("---")
    prompt = st.text_area("Pregunta:", value="¿Cómo está la cartera general?")
    if st.button("Consultar"):
        st.info("Asistente listo. Configura tu GROQ_API_KEY para consultas avanzadas.")

# =========================================================
# 8. SOBRE NOSOTROS
# =========================================================
elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
    st.title("💎 Sobre Nuestros Microcréditos")
    st.markdown("---")
    st.write("Fondo colaborativo familiar y de amigos con tasas justas del 3% mensual.")
