import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import urllib.parse
from fpdf import FPDF

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Entre Amigos Capital - Dashboard",
    page_icon="🤝",
    layout="wide"
)

EXCEL_FILE_DEFAULT = "proyecto microcréditos copia 3.xlsx.xlsx"

# ---------------------------------------------------------
# CLASE Y FUNCIÓN PARA GENERAR EL COMPROBANTE EN PDF (fpdf2)
# ---------------------------------------------------------
class ComprobantePDF(FPDF):
    def header(self):
        self.set_fill_color(27, 54, 93)  # Azul oscuro institucional
        self.rect(0, 0, 210, 32, 'F')
        self.set_font("Arial", "B", 16)
        self.set_text_color(255, 255, 255)
        self.cell(0, 8, "Entre Amigos Capital", ln=True, align="C")
        self.set_font("Arial", "", 10)
        self.set_text_color(209, 216, 224)
        self.cell(0, 5, "Comprobante Oficial de Recaudo de Pago", ln=True, align="C")
        self.ln(10)

def generar_pdf_comprobante(pago_info):
    pdf = ComprobantePDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Badge Recibo N°
    pdf.set_fill_color(39, 174, 96)  # Verde
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 8, f"RECIBO N°: {pago_info['pago_id']}", ln=True, align="C", fill=True)
    pdf.ln(4)

    # 1. Datos Cliente
    pdf.set_text_color(27, 54, 93)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Datos del Cliente y Crédito", ln=True)
    pdf.set_draw_color(52, 152, 219)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(44, 62, 80)
    pdf.cell(50, 5, "Cliente:", 0)
    pdf.cell(0, 5, f"{pago_info['cliente_nombre']} ({pago_info['cliente_id']})", ln=True)
    pdf.cell(50, 5, "Crédito N°:", 0)
    pdf.cell(0, 5, str(pago_info['credito_id']), ln=True)
    pdf.cell(50, 5, "Fecha de Pago:", 0)
    pdf.cell(0, 5, str(pago_info['fecha_pago']), ln=True)
    pdf.cell(50, 5, "Medio de Pago:", 0)
    pdf.cell(0, 5, str(pago_info['medio_pago']), ln=True)
    pdf.ln(4)

    # 2. Desglose
    pdf.set_text_color(27, 54, 93)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Desglose de la Transacción", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(44, 62, 80)
    pdf.cell(120, 5, "Abono a Intereses:", 0)
    pdf.cell(0, 5, f"${pago_info['pago_interes']:,.0f} COP", ln=True, align="R")
    pdf.cell(120, 5, "Abono a Capital:", 0)
    pdf.cell(0, 5, f"${pago_info['pago_capital']:,.0f} COP", ln=True, align="R")
    
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(241, 242, 246)
    pdf.cell(120, 7, f"TOTAL RECIBIDO ({pago_info['concepto']}):", fill=True)
    pdf.cell(0, 7, f"${pago_info['valor_pago']:,.0f} COP", ln=True, align="R", fill=True)
    pdf.ln(4)

    # 3. Estado Deuda
    pdf.set_text_color(27, 54, 93)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Estado Actualizado de la Deuda", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(44, 62, 80)
    pdf.cell(50, 5, "Capital Pendiente:", 0)
    pdf.cell(0, 5, f"${pago_info['nuevo_cap_pend']:,.0f} COP", ln=True)
    pdf.cell(50, 5, "Intereses Pendientes:", 0)
    pdf.cell(0, 5, f"${pago_info['nuevo_int_pend']:,.0f} COP", ln=True)
    
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(192, 57, 43)
    pdf.cell(50, 5, "Deuda Total Pendiente:", 0)
    pdf.cell(0, 5, f"${pago_info['nueva_deuda_total']:,.0f} COP", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(44, 62, 80)
    pdf.cell(50, 5, "Estado del Crédito:", 0)
    pdf.cell(0, 5, str(pago_info['nuevo_estado']), ln=True)

    if pago_info.get('observaciones'):
        pdf.cell(50, 5, "Observaciones:", 0)
        pdf.cell(0, 5, str(pago_info['observaciones']), ln=True)

    pdf.ln(10)
    pdf.set_font("Arial", "I", 8)
    pdf.set_text_color(127, 140, 141)
    pdf.cell(0, 4, "Gracias por mantener tu crédito al día. Este documento sirve como soporte oficial de recaudo.", ln=True, align="C")
    pdf.cell(0, 4, "Entre Amigos Capital - Créditos justos sobre la base de la confianza.", ln=True, align="C")

    return bytes(pdf.output())

# ---------------------------------------------------------
# MENÚ LATERAL & CARGADOR DE ARCHIVO
# ---------------------------------------------------------
st.sidebar.title("🤝 Entre Amigos Capital")
st.sidebar.caption("Créditos justos sobre la base de la confianza")

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
        "📝 Registrar Pago", 
        "🔔 Alertas & Cobros", 
        "🧮 Simulador de Créditos", 
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
    deuda_vencida = df_estado_cartera['deuda_vencida'].sum() if 'deuda_vencida' in df_estado_cartera.columns else 0
    
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
        if not df_clientes.empty:
            df_clientes.to_excel(writer, sheet_name='Clientes', index=False)
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

# ---------------------------------------------------------
# LÓGICA DE NAVEGACIÓN PRINCIPAL
# ---------------------------------------------------------
if not df_creditos.empty:

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
                             color_discrete_map={"Al Día": "#2ecc71", "En Mora": "#e74c3c"})
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
            deuda_vencida = cartera_cliente['deuda_vencida'].sum() if not cartera_cliente.empty and 'deuda_vencida' in cartera_cliente.columns else 0
            deuda_total = cartera_cliente['deuda_total_pendiente'].sum() if not cartera_cliente.empty else 0

            tiene_mora = any(cartera_cliente['estado'].astype(str).str.contains('mora', case=False, na=False)) if not cartera_cliente.empty else False

            if tiene_mora or deuda_vencida > 0:
                st.error(f"⚠️ **Cliente con Cuotas/Intereses Vencidos:** Presenta un saldo pendiente en mora.")
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

    # =========================================================
    # 3. REGISTRAR PAGO CON COMPROBANTE PDF Y WHATSAPP
    # =========================================================
    elif opcion_menu == "📝 Registrar Pago":
        st.title("📝 Formulario de Registro de Pagos")
        st.markdown("Ingresa los datos del pago para actualizar las tablas de Excel y generar el comprobante PDF.")
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
                valor_pago = st.number_input("Valor Pagado (COP):", min_value=1000, value=90000, step=5000, format="%d")
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

                # Generar ID
                existentes_ids = st.session_state['df_pagos']['pago_id'].dropna().tolist()
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

                    # Actualizar df_pagos
                    st.session_state['df_pagos'] = pd.concat([st.session_state['df_pagos'], pd.DataFrame([nueva_fila_pago])], ignore_index=True)

                    # Actualizar Creditos
                    idx_cred = df_creditos.index[df_creditos['credito_id'] == credito_id_pago].tolist()
                    if idx_cred:
                        ic = idx_cred[0]
                        saldo_cap_prev = df_creditos.loc[ic, 'saldo_capital'] if pd.notnull(df_creditos.loc[ic, 'saldo_capital']) else df_creditos.loc[ic, 'capital_inicial']
                        nuevo_saldo_cap = max(0, saldo_cap_prev - pago_capital)
                        st.session_state['df_creditos'].loc[ic, 'saldo_capital'] = nuevo_saldo_cap
                        if nuevo_saldo_cap == 0:
                            st.session_state['df_creditos'].loc[ic, 'estado_credito'] = "Finalizado"

                    # Actualizar Estado_Cartera
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

                        nueva_deuda_total = nuevo_cap_pend + nuevo_int_pend
                        st.session_state['df_estado_cartera'].loc[ie, 'deuda_total_pendiente'] = nueva_deuda_total

                        if nueva_deuda_total == 0:
                            nuevo_estado = "Paz y salvo"
                        else:
                            nuevo_estado = "Al día"
                        st.session_state['df_estado_cartera'].loc[ie, 'estado'] = nuevo_estado

                    st.success(f"✅ ¡Pago **{proximo_pago_id}** registrado exitosamente!")

                    # Generar PDF
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
                        msg_wa = (
                            f"Hola *{cliente_pago}* 👋,\n\n"
                            f"Confirmamos el registro de tu pago en *Entre Amigos Capital* 🤝:\n"
                            f"🔹 *Recibo N°:* {proximo_pago_id}\n"
                            f"🔹 *Monto Recibido:* ${valor_pago:,.0f} COP\n"
                            f"🔹 *Abono a Intereses:* ${pago_interes:,.0f} COP\n"
                            f"🔹 *Abono a Capital:* ${pago_capital:,.0f} COP\n"
                            f"🔹 *Saldo Restante:* ${nueva_deuda_total:,.0f} COP\n\n"
                            f"¡Muchas gracias por mantener tu crédito al día! 🟢"
                        )
                        num_tel = "".join(filter(str.isdigit, str(telefono_cliente)))
                        if len(num_tel) == 10 and not num_tel.startswith("57"):
                            num_tel = "57" + num_tel
                        msg_enc = urllib.parse.quote(msg_wa)
                        
                        st.link_button("💬 Enviar Resumen por WhatsApp", f"https://wa.me/{num_tel}?text={msg_enc}", use_container_width=True)

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
    # 4. ALERTAS DE MORA Y GESTIÓN DE COBROS (SIN INTERESES DE MORA)
    # =========================================================
    elif opcion_menu == "🔔 Alertas & Cobros":
        st.title("🔔 Alertas de Mora y Gestión de Cobros")
        st.markdown("Seguimiento de días de retraso en los pagos y generación de avisos vía WhatsApp sin cobro de intereses por mora.")
        st.markdown("---")

        if st.session_state['df_estado_cartera'].empty:
            st.info("No hay información disponible en la tabla de estado de cartera.")
        else:
            df_mora = st.session_state['df_estado_cartera'].merge(
                df_clientes[['cliente_id', 'nombre', 'telefono']], on='cliente_id', how='left'
            )

            df_mora = df_mora[df_mora['deuda_total_pendiente'] > 0].copy()

            if df_mora.empty:
                st.success("🎉 ¡Excelente! Toda la cartera está al día y en paz y salvo.")
            else:
                if 'dias_mora' not in df_mora.columns:
                    df_mora['dias_mora'] = 0

                df_mora['nivel_riesgo'] = df_mora['dias_mora'].apply(
                    lambda d: "🟢 Al Día / Preventivo" if d <= 0 
                    else ("🟡 Retraso Leve (1-30 días)" if d <= 30 
                    else ("🟠 Retraso Medio (31-60 días)" if d <= 60 
                    else "🔴 Retraso Alto (>60 días)"))
                )

                c_a1, c_a2, c_a3, c_a4 = st.columns(4)
                tot_clientes_mora = len(df_mora[df_mora['dias_mora'] > 0])
                monto_pendiente = df_mora['deuda_total_pendiente'].sum()
                max_dias = df_mora['dias_mora'].max()

                c_a1.metric("Clientes Activos", f"{len(df_mora)}")
                c_a2.metric("Clientes con Retraso", f"{tot_clientes_mora}", delta=f"{tot_clientes_mora} alertas", delta_color="inverse")
                c_a3.metric("Deuda Pendiente Total", f"${monto_pendiente:,.0f} COP")
                c_a4.metric("Máximo Días de Retraso", f"{max_dias} días")

                st.markdown("---")

                filtro_nivel = st.multiselect(
                    "Filtrar por Días de Retraso:",
                    options=["🟢 Al Día / Preventivo", "🟡 Retraso Leve (1-30 días)", "🟠 Retraso Medio (31-60 días)", "🔴 Retraso Alto (>60 días)"],
                    default=["🟡 Retraso Leve (1-30 días)", "🟠 Retraso Medio (31-60 días)", "🔴 Retraso Alto (>60 días)"]
                )

                df_filtrado = df_mora[df_mora['nivel_riesgo'].isin(filtro_nivel)]

                st.subheader("📋 Detalle de Clientes y Gestión de Cobro")

                for _, row in df_filtrado.iterrows():
                    dias = int(row['dias_mora'])
                    with st.expander(f"{row['nivel_riesgo']} | {row['nombre']} - Crédito `{row['credito_id']}` ({dias} días de retraso)"):
                        col_m1, col_m2 = st.columns([1, 1])

                        with col_m1:
                            st.write(f"**Cliente:** {row['nombre']}")
                            st.write(f"**Teléfono:** {row['telefono']}")
                            st.write(f"**Días de Retraso:** {dias} días")
                            st.write(f"**Capital Pendiente:** ${row['capital_pendiente']:,.0f} COP")
                            st.write(f"**Intereses Pendientes:** ${row['interes_pendiente']:,.0f} COP")
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
                                    f"Te escribimos de *Entre Amigos Capital* 🤝 para recordarte que registras *{dias} días de retraso* en tu crédito *{row['credito_id']}*.\n\n"
                                    f"🔹 *Saldo Pendiente:* ${row['deuda_total_pendiente']:,.0f} COP\n\n"
                                    f"Te invitamos a ponerte al día para mantener tus condiciones activas. Por favor confirmanos cuándo podrías realizar el abono. ¡Gracias!"
                                )
                            else:
                                texto_base = (
                                    f"Hola *{row['nombre']}* 👋,\n\n"
                                    f"Nos comunicamos de *Entre Amigos Capital* 🤝 referente a tu crédito *{row['credito_id']}*, el cual presenta un retraso de *{dias} días* con un saldo de *${row['deuda_total_pendiente']:,.0f} COP*.\n\n"
                                    f"Te pedimos ponerte en contacto con nosotros hoy mismo para definir una fecha de pago. ¡Agradecemos tu atención!"
                                )

                            msg_editado = st.text_area("Editar mensaje antes de enviar:", value=texto_base, height=130, key=f"txt_{row['credito_id']}")
                            
                            num_tel = "".join(filter(str.isdigit, str(row['telefono'])))
                            if len(num_tel) == 10 and not num_tel.startswith("57"):
                                num_tel = "57" + num_tel

                            msg_enc = urllib.parse.quote(msg_editado)
                            
                            st.link_button("💬 Enviar Recordatorio por WhatsApp", f"https://wa.me/{num_tel}?text={msg_enc}", use_container_width=True)

    # =========================================================
    # 5. SIMULADOR DE CRÉDITOS
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
    # 6. SOBRE NOSOTROS & POLÍTICAS
    # =========================================================
    elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
        st.title("🤝 Sobre Entre Amigos Capital")
        st.markdown("---")
        st.subheader("💡 Nuestra Propuesta de Valor")
        st.write("Acceso a microcréditos ágiles para familiares y amigos a tasas justas (3% mensual).")
