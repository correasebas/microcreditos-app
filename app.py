import streamlit as st
import pandas as pd
import plotly.express as pdx
from datetime import datetime
from groq import Groq

# Configuración inicial de la página
st.set_page_config(
    page_title="Entre Amigos Capital - Fondo Familiar",
    page_icon="💰",
    layout="wide"
)

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
    st.metric("Total Clientes Registrados", len(st.session_state['df_clientes']))


# =========================================================
# 2. FICHA POR CLIENTE
# =========================================================
elif opcion_menu == "👤 Ficha por Cliente":
    st.title("👤 Ficha Detallada por Cliente")
    st.markdown("Consulta el historial y estado actual de cada integrante o prestatario.")
    st.markdown("---")
    st.write("Selecciona o busca un cliente para ver sus detalles.")


# =========================================================
# 3. NUEVOS REGISTROS
# =========================================================
elif opcion_menu == "➕ Nuevos Registros":
    st.title("➕ Nuevos Registros")
    st.markdown("Registra nuevos clientes o créditos otorgados.")
    st.markdown("---")


# =========================================================
# 4. REGISTRAR PAGO
# =========================================================
elif opcion_menu == "📝 Registrar Pago":
    st.title("📝 Registro de Abonos y Pagos")
    st.markdown("Actualiza las cuotas o saldos pendientes.")
    st.markdown("---")


# =========================================================
# 5. GESTIÓN DE COBRO
# =========================================================
elif opcion_menu == "⚖️ Gestión de Cobro":
    st.title("⚖️ Gestión de Cartera y Cobros")
    st.markdown("Visualiza alertas y cuotas próximas a vencer o atrasadas.")
    st.markdown("---")


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

    # Campo para ingresar la clave de Groq manualmente en la pantalla
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
                        model="llama-3.1-8b-instant",  # Modelo estable y disponible universalmente en Groq
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
