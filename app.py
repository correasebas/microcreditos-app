import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import os

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
        df_estado_cartera = pd.read_excel(xls, sheet_name='Estado_Cartera') if 'Estado_Cartera' in xls.sheet_names else pd.DataFrame()
        df_resumen = pd.read_excel(xls, sheet_name='Resumen_Cartera') if 'Resumen_Cartera' in xls.sheet_names else pd.DataFrame()
        df_calendario = pd.read_excel(xls, sheet_name='Calendario_Intereses') if 'Calendario_Intereses' in xls.sheet_names else pd.DataFrame()

        return df_clientes, df_creditos, df_pagos, df_estado_cartera, df_resumen, df_calendario
    except Exception as e:
        st.error(f"Error al cargar el archivo de Excel: {e}")
        return None, None, None, None, None, None

df_clientes_init, df_creditos_init, df_pagos_init, df_estado_cartera_init, df_resumen_init, df_calendario_init = load_data()

# Inicializar st.session_state para mantener persistentes las tablas en memoria
if 'df_clientes' not in st.session_state:
    st.session_state['df_clientes'] = df_clientes_init.copy() if df_clientes_init is not None else pd.DataFrame()
if 'df_creditos' not in st.session_state:
    st.session_state['df_creditos'] = df_creditos_init.copy() if df_creditos_init is not None else pd.DataFrame()
if 'df_pagos' not in st.session_state:
    st.session_state['df_pagos'] = df_pagos_init.copy() if df_pagos_init is not None else pd.DataFrame()
if 'df_estado_cartera' not in st.session_state:
    st.session_state['df_estado_cartera'] = df_estado_cartera_init.copy() if df_estado_cartera_init is not None else pd.DataFrame()
if 'df_calendario' not in st.session_state:
    st.session_state['df_calendario'] = df_calendario_init.copy() if df_calendario_init is not None else pd.DataFrame()

df_clientes = st.session_state['df_clientes']
df_creditos = st.session_state['df_creditos']
df_pagos = st.session_state['df_pagos']
df_estado_cartera = st.session_state['df_estado_cartera']
df_calendario = st.session_state['df_calendario']

# ---------------------------------------------------------
# FUNCIÓN DE RECALCULO DE RESUMEN DE CARTERA
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

# Exportar todas las hojas actualizadas a Excel
def exportar_excel_completo():
    output = io.BytesIO()
    df_resumen_actualizado = recalcular_resumen_cartera()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df_clientes.empty:
            df_clientes.to_excel(writer, sheet_name='Clientes', index=False)
        if not df_creditos.empty:
            cols_cred = [c for c in df_creditos.columns if not str(c).startswith('Unnamed')]
            df_creditos[cols_cred].to_excel(writer, sheet_name='Creditos', index=False)
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
# MENÚ LATERAL
# ---------------------------------------------------------
st.sidebar.title("🤝 Entre Amigos Capital")
st.sidebar.caption("Créditos justos sobre la base de la confianza")

opcion_menu = st.sidebar.radio(
    "Selecciona una sección:",
    ["📊 Dashboard General", "👤 Ficha por Cliente", "📝 Registrar Pago", "🧮 Simulador de Créditos", "ℹ️ Sobre Nosotros & Políticas"]
)

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
            deuda_vencida = cartera_cliente['deuda_vencida'].sum() if not cartera_cliente.empty else 0
            deuda_total = cartera_cliente['deuda_total_pendiente'].sum() if not cartera_cliente.empty else 0

            tiene_mora = any(cartera_cliente['estado'].astype(str).str.contains('mora', case=False, na=False)) if not cartera_cliente.empty else False

            if tiene_mora or deuda_vencida > 0:
                st.error(f"⚠️ **Cliente con Cuotas/Intereses Vencidos:** Presenta una mora de **${deuda_vencida:,.0f} COP**.")
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
    # 3. REGISTRAR PAGO
    # =========================================================
    elif opcion_menu == "📝 Registrar Pago":
        st.title("📝 Formulario de Registro de Pagos")
        st.markdown("Ingresa los datos del pago para actualizar todas las tablas de Excel en cadena.")
        st.markdown("---")

        col_f1, col_f2 = st.columns([1, 1])

        with col_f1:
            lista_cli = df_clientes['nombre'].dropna().unique()
            cliente_pago = st.selectbox("Selecciona el cliente:", lista_cli)

            info_cli_pago = df_clientes[df_clientes['nombre'] == cliente_pago].iloc[0]
            cliente_id_pago = info_cli_pago['cliente_id']

            creditos_cli = df_creditos[df_creditos['cliente_id'] == cliente_id_pago]

            if creditos_cli.empty:
                st.warning("Este cliente no tiene créditos activos registrados.")
            else:
                opciones_credito = [f"{row['credito_id']} - Capital: ${row['capital_inicial']:,.0f} ({row['modalidad']})" for _, row in creditos_cli.iterrows()]
                credito_sel_str = st.selectbox("Selecciona el Crédito:", opciones_credito)
                credito_id_pago = credito_sel_str.split(" - ")[0]

                fecha_pago = st.date_input("Fecha del Pago:", datetime.today())
                
                # MODIFICACIÓN: Solamente Transferencia y Efectivo
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

                # CALCULAR PRÓXIMO ID DE PAGO (PAG069, etc.)
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

                if st.button("💾 Registrar Pago y Actualizar Excel Completo", type="primary"):
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

                    # 1. ACTUALIZAR 'df_pagos' en la sesión
                    st.session_state['df_pagos'] = pd.concat([st.session_state['df_pagos'], pd.DataFrame([nueva_fila_pago])], ignore_index=True)

                    # 2. ACTUALIZAR 'Creditos'
                    idx_cred = df_creditos.index[df_creditos['credito_id'] == credito_id_pago].tolist()
                    if idx_cred:
                        ic = idx_cred[0]
                        saldo_cap_prev = df_creditos.loc[ic, 'saldo_capital'] if pd.notnull(df_creditos.loc[ic, 'saldo_capital']) else df_creditos.loc[ic, 'capital_inicial']
                        nuevo_saldo_cap = max(0, saldo_cap_prev - pago_capital)
                        st.session_state['df_creditos'].loc[ic, 'saldo_capital'] = nuevo_saldo_cap
                        if nuevo_saldo_cap == 0:
                            st.session_state['df_creditos'].loc[ic, 'estado_credito'] = "Finalizado"

                    # 3. ACTUALIZAR 'Calendario_Intereses'
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

                    # 4. ACTUALIZAR 'Estado_Cartera'
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
                            st.session_state['df_estado_cartera'].loc[ie, 'estado'] = "Paz y salvo"
                        elif nuevo_vencido == 0:
                            st.session_state['df_estado_cartera'].loc[ie, 'estado'] = "Al día"

                    st.success(f"✅ ¡Pago **{proximo_pago_id}** registrado exitosamente!")

        st.markdown("---")

        st.subheader("📥 Descargar Libro de Excel Actualizado")
        st.markdown("Para mantener la copia de respaldo sincronizada, descarga aquí el archivo con todas las hojas actualizadas:")

        excel_bytes = exportar_excel_completo()
        st.download_button(
            label="📥 Descargar Excel Actualizado (.xlsx)",
            data=excel_bytes,
            file_name="proyecto_microcreditos_actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.subheader("📋 Últimos Pagos Registrados")
        cols_p = [c for c in st.session_state['df_pagos'].columns if not str(c).startswith('Unnamed')]
        st.dataframe(st.session_state['df_pagos'][cols_p].tail(10), use_container_width=True)

    # =========================================================
    # 4. SIMULADOR DE CRÉDITOS
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
    # 5. SOBRE NOSOTROS & POLÍTICAS
    # =========================================================
    elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
        st.title("🤝 Sobre Entre Amigos Capital")
        st.markdown("---")
        st.subheader("💡 Nuestra Propuesta de Valor")
        st.write("Acceso a microcréditos ágiles para familiares y amigos a tasas justas (3% mensual).")
