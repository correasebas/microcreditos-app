# =========================================================
        # GRÁFICO DE PROYECCIÓN DE INGRESOS (FLUJO DE CAJA) - CORREGIDO
        # =========================================================
        st.subheader("📈 Proyección de Flujo de Caja Esperado (Año 2026)")
        st.markdown("Ingresos programados mes a mes por concepto de **Capital** e **Intereses** según el calendario de pagos.")

        df_cal_actual = st.session_state.get('df_calendario', pd.DataFrame())
        
        if not df_cal_actual.empty and 'fecha_programada' in df_cal_actual.columns:
            df_cal_cp = df_cal_actual.copy()
            df_cal_cp['fecha_programada'] = pd.to_datetime(df_cal_cp['fecha_programada'], errors='coerce')
            
            # Limpiar posibles nulos y asegurar formato numérico para los valores
            df_cal_cp['capital_programado'] = pd.to_numeric(df_cal_cp['capital_programado'], errors='fillna').fillna(0)
            df_cal_cp['interes_programado'] = pd.to_numeric(df_cal_cp['interes_programado'], errors='fillna').fillna(0)

            # Filtrar el año 2026 (o mostrar todo el calendario si prefieres ver la data real disponible)
            df_cal_2026 = df_cal_cp[df_cal_cp['fecha_programada'].dt.year == 2026].copy()

            if df_cal_2026.empty:
                # Si no hay datos específicos de 2026, usamos toda la data disponible para depurar
                df_cal_2026 = df_cal_cp.copy()

            if not df_cal_2026.empty:
                df_cal_2026['Mes_Num'] = df_cal_2026['fecha_programada'].dt.month

                # Mapeo completo de meses en español
                meses_es = {
                    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 
                    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto', 
                    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
                }
                
                df_cal_2026['Mes_Texto'] = df_cal_2026['Mes_Num'].map(meses_es)

                # Agrupar por mes sumando los valores reales
                df_proyeccion = df_cal_2026.groupby(['Mes_Num', 'Mes_Texto'])[['capital_programado', 'interes_programado']].sum().reset_index()
                df_proyeccion = df_proyeccion.sort_values('Mes_Num')

                # Reestructurar para gráfico de barras apiladas con Plotly
                df_melted = df_proyeccion.melt(
                    id_vars=['Mes_Texto', 'Mes_Num'], 
                    value_vars=['capital_programado', 'interes_programado'],
                    var_name='Concepto', 
                    value_name='Valor'
                )
                df_melted['Concepto'] = df_melted['Concepto'].replace({
                    'capital_programado': 'Abono a Capital', 
                    'interes_programado': 'Intereses'
                })

                fig_proy = px.bar(
                    df_melted, 
                    x='Mes_Texto', 
                    y='Valor', 
                    color='Concepto',
                    barmode='stack',
                    color_discrete_map={'Intereses': '#00D26A', 'Abono a Capital': '#3498DB'},
                    labels={'Mes_Texto': 'Mes', 'Valor': 'Valor Esperado (COP)', 'Concepto': 'Concepto'}
                )
                fig_proy.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", 
                    plot_bgcolor="rgba(0,0,0,0)", 
                    font_color="#E6EDF3",
                    xaxis=dict(categoryorder='array', categoryarray=list(meses_es.values())),
                    legend=dict(font=dict(color="#E6EDF3"))
                )
                st.plotly_chart(fig_proy, use_container_width=True)
            else:
                st.info("ℹ️ No se encontraron registros con fechas válidas en el calendario de intereses.")
        else:
            st.info("ℹ️ El calendario de intereses no está cargado o le falta la columna 'fecha_programada'.")
