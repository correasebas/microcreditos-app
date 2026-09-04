# ---------------------------------------------------------
# OPCIONES DEL MENÚ
# ---------------------------------------------------------
if opcion_menu == "📊 Dashboard General":
    st.title("📊 Control General de Cartera")
    st.markdown("---")
    # ... (código del dashboard) ...

elif opcion_menu == "👤 Ficha por Cliente":
    st.title("👤 Ficha de Cliente e Historial de Deuda")
    st.markdown("---")
    # ... (código de ficha de cliente) ...

elif opcion_menu == "📝 Registrar Pago":
    st.title("📝 Formulario de Registro de Pagos")
    st.markdown("---")
    # ... (código de registrar pago) ...

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

elif opcion_menu == "🧮 Simulador de Créditos":
    st.title("🧮 Simulador de Créditos")
    st.markdown("---")
    # ... (código del simulador) ...

elif opcion_menu == "ℹ️ Sobre Nosotros & Políticas":
    st.title("🤝 Sobre Entre Amigos Capital")
    st.markdown("---")
    # ... (código de sobre nosotros) ...
