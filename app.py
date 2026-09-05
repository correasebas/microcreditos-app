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

                        # Actualizar session_state de forma inmediata
                        st.session_state['df_clientes'] = pd.concat([st.session_state['df_clientes'], pd.DataFrame([nueva_fila_cliente])], ignore_index=True)
                        st.success(f"✅ ¡Cliente **{nombre_nuevo}** registrado con éxito bajo el ID `{proximo_cli_id}`! Ya puedes seleccionarlo en la pestaña de préstamos.")
                    else:
                        st.error("⚠️ El campo de nombre completo es obligatorio.")

        with tab_cred:
            st.subheader("Otorgar Crédito y Sincronizar Hojas de Excel")
            
            # Usar directamente el session_state actualizado para reflejar los nuevos registros al instante
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
