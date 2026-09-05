import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from groq import Groq

# Configuración inicial de la página
st.set_page_config(
    page_title="Entre Amigos Capital - Fondo Familiar",
    page_icon="💰",
    layout="wide"
)

# Estilo corporativo (Azul Marino / Verde Esmeralda)
st.markdown("""
    <style>
    .stApp {
        background-color: #0d1b2a;
        color: #e0e1dd;
    }
    sidebar .stSidebar {
        background-color: #1b263b;
    }
    .stTextInput input, .stTextArea textarea, .stNumberInput input {
        background-color: #1b263b;
        color: #ffffff;
        border: 1px solid #1b4332;
    }
    .stButton button {
        background-color: #2d6a4f;
        color: white;
        border-radius: 5px;
    }
    .stButton button:hover {
        background-color: #40916c;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# Inicializar DataFrames en st.session_state si no existen
if 'df_clientes' not in st.session_state:
    st.session_state['df_clientes'] = pd.DataFrame(columns=["ID", "Nombre", "Teléfono", "Crédito Activo", "Monto"])

if 'df_estado_cartera' not in st.session_state:
    st.session_state['df_estado_cartera'] = pd.DataFrame(columns=["ID Crédito", "Cliente", "Saldo Pendiente", "Estado"])

# Menú lateral
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
        "ℹ️ Sobre Nosotros & Políticas",
    ],
)

# =========================================================
# 1. DASHBOARD GENERAL
# =========================================================
if opcion_menu == "📊 Dashboard General":
    st.title("📊 Dashboard General - Entre Amigos Capital")
    st.markdown("Resumen general del estado financiero del fondo familiar.")
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Clientes Registrados", len(st.session_state['df_clientes']))
    with col2:
        total_monto = st.session_state['df_clientes']['Monto'].sum() if not st.session_state['df_clientes'].empty else 0
        st.metric("Capital Total Prestado", f"${total_monto:,.2f}")

    st.markdown("### 📋 Listado Actual de Clientes")
    if not st.session_state['df_clientes'].empty:
        st.dataframe(st.session_state['df_clientes'], use_container_width=True)
    else:
        st.info("No hay clientes registrados todavía. Ve a la sección 'Nuevos Registros' para agregar el primero.")


# =========================================================
# 2. FICHA POR CLIENTE
# =========================================================
elif opcion_menu == "👤 Ficha por Cliente":
    st.title("👤 Ficha Detallada por Cliente")
    st.markdown("Consulta el historial y estado actual de cada integrante o prestatario.")
    st.markdown("---")
    if not st.session_state['df_clientes'].empty:
        cliente_seleccionado = st.selectbox("Selecciona un cliente:", st.session_state['df_clientes']['Nombre'].unique())
        datos_cliente = st.session_state['df_clientes'][st.session_state['df_clientes']['Nombre'] == cliente_seleccionado]
        st.dataframe(datos_cliente, use_container_width=True)
    else:
        st.info("Primero debes registrar clientes en la sección 'Nuevos Registros'.")


# =========================================================
# 3. NUEVOS REGISTROS
# =========================================================
elif opcion_menu == "➕ Nuevos Registros":
    st.title("➕ Nuevos Registros de Clientes y Créditos")
    st.markdown("Ingresa la información para dar de alta un nuevo crédito en el fondo.")
    st.markdown("---")

    with st.form("form_nuevo_cliente"):
        nombre = st.text_input("Nombre Completo del Cliente")
        telefono = st.text_input("Teléfono de Contacto")
        id_credito = st.text_input("ID o Código del Crédito (ej. CR-001)")
        monto = st.number_input("Monto del Crédito", min_value=0.0, step=1000.0)
        
        submitted = st.form_submit_button("Guardar Registro")
        
        if submitted:
            if nombre and id_credito and monto > 0:
                # Agregar a clientes
                nuevo_cliente = pd.DataFrame([[id_credito, nombre, telefono, "Activo", monto]], 
                                             columns=["ID", "Nombre", "Teléfono", "Crédito Activo", "Monto"])
                st.session_state['df_clientes'] = pd.concat([st.session_state['df_clientes'], nuevo_cliente], ignore_index=True)
                
                # Agregar a estado de cartera
                nueva_cartera = pd.DataFrame([[id_credito, nombre, monto, "Al Día"]], 
                                             columns=["ID Crédito", "Cliente", "Saldo Pendiente", "Estado"])
                st.session_state['df_estado_cartera'] = pd.concat([st.session_state['df_estado_cartera'], nueva_cartera], ignore_index=True)
                
                st.success(f"¡Cliente {nombre} y crédito {id_credito} guardados con éxito!")
            else:
                st.error("Por favor completa los campos obligatorios y asegúrate de que el monto sea mayor a 0.")


# =========================================================
# 4. REGISTRAR PAGO
# =========================================================
elif opcion_menu == "📝 Registrar Pago":
    st.title("📝 Registro de Abonos y Pagos")
    st.markdown("Actualiza las cuotas o saldos pendientes.")
    st.markdown("---")
    st.write("Módulo en preparación. Pronto podrás registrar abonos aquí.")


# =========================================================
# 5. GESTIÓN DE COBRO
# =========================================================
elif opcion_menu == "⚖️ Gestión de Cobro":
    st.title("⚖️ Gestión de Cartera y Cobros")
    st.markdown("Visualiza alertas y cuotas próximas a vencer o atrasadas.")
    st.markdown("---")
    if not st.session_state['df_estado_cartera'].empty:
        st.dataframe(st.session_state['df_estado_cartera'], use_container_width=True)
    else:
        st.info("No hay registros de cartera activos.")


# =========================================================
# 6. SIMULADOR DE CRÉDITOS
# =========================================================
elif opcion_menu == "🧮 Simulador de Créditos":
    st.title("🧮 Simulador de Créditos")
    st.markdown("Calcula cuotas, intereses y tablas de amortización.")
    st.markdown("---")


# =========================================================
# 7. ASISTENTE INTELIGENTE (IA)
# =========================================================
elif opcion_menu == "🤖 Asistente IA":
    st.title("🤖 Asistente Inteligente del Fondo")
    st.markdown("Hazle preguntas en lenguaje natural sobre tus clientes, créditos, pagos o estado de cartera.")
    st.markdown("---")

    groq_api_key = st.text_input("Ingresa tu Clave de API de Groq:", type="password", help="Pega aquí tu clave de Groq (gsk_...)")

    pregunta_usuario = st.text_area(
        "¿Qué te gustaría saber de la cartera?",
        placeholder="Ej: ¿Qué clientes tienen créditos activos y cuál es el total prestado?"
    )

    if st.button("Consultar al Asistente", type="primary"):
        if not groq_api_key:
            st.error("⚠️ Por favor ingresa tu clave de API de Groq para continuar.")
        elif not pregunta_usuario:
            st.error("⚠️ Escribe una pregunta para el asistente.")
        else:
            try:
                with st.spinner("Analizando datos y generando respuesta..."):
                    client = Groq(api_key=groq_api_key)

                    resumen_clientes = st.session_state['df_clientes'].to_string() if not st.session_state['df_clientes'].empty else "Sin clientes registrados"
                    resumen_estado = st.session_state['df_estado_cartera'].to_string() if not st.session_state['df_estado_cartera'].empty else "Sin cartera registrada"

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
                        model="llama-3.1-8b-instant",
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
    st.title("ℹ️ Sobre el Fondo & Políticas")
    st.markdown("Información corporativa, misión y normativas de Entre Amigos Capital.")
    st.markdown("---")
