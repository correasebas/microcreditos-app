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
        modalidad = str(cred.get('modalidad', 'INTERES_MENSUAL'))
        
        pagos_cred = df_pagos[df_pagos['credito_id'] == c_id] if not df_pagos.empty else pd.DataFrame()
        cap_pagado = float(pagos_cred['pago_capital'].sum()) if not pagos_cred.empty and 'pago_capital' in pagos_cred.columns else 0.0
        cap_pend = max(0.0, cap_ini - cap_pagado)
        
        total_interes_pagado = float(pagos_cred['pago_interes'].sum()) if not pagos_cred.empty and 'pago_interes' in pagos_cred.columns else 0.0
        
        # 1. LÓGICA PARA INTERÉS MENSUAL
        if "INTERES_MENSUAL" in modalidad.upper():
            if pd.notnull(f_desembolso):
                r = relativedelta(hoy, f_desembolso)
                meses_transcurridos = r.years * 12 + r.months
                if hoy.day < f_desembolso.day:
                    meses_transcurridos = max(0, meses_transcurridos - 1)
            else:
                meses_transcurridos = 0
                
            interes_generado_total = meses_transcurridos * (cap_ini * tasa)
            interes_pend = max(0.0, interes_generado_total - total_interes_pagado)
            
            deuda_total = cap_pend + interes_pend
            estado = "En mora" if interes_pend > 0 else "Al día"
            deuda_vencida = interes_pend if estado == "En mora" else 0.0

        # 2. LÓGICA PARA CUOTAS FIJAS / OTROS CRÉDITOS
        else:
            # Para cuotas fijas, evaluamos saldo pendiente y si la fecha de pago del mes ya venció sin cubrirse
            valor_cuota_est = float(cred.get('valor_cuota', 0))
            dia_pago = int(cred.get('dia_pago', f_desembolso.day if pd.notnull(f_desembolso) else 1))
            
            # Verificamos si este mes ya pasó la fecha límite de pago y no hay pago registrado en el mes corriente
            en_mora = False
            if pd.notnull(f_desembolso) and cap_pend > 1:
                # Construimos la fecha de pago del mes actual
                try:
                    limite_mes_actual = datetime(hoy.year, hoy.month, dia_pago)
                except ValueError:
                    limite_mes_actual = datetime(hoy.year, hoy.month, 28) # Ajuste por meses cortos
                
                # Si hoy ya pasó la fecha de pago de este mes, revisamos si se pagó en el mes corriente
                if hoy > limite_mes_actual:
                    pagos_mes_actual = pagos_cred[
                        (pd.to_datetime(pagos_cred['fecha_pago']).dt.year == hoy.year) & 
                        (pd.to_datetime(pagos_cred['fecha_pago']).dt.month == hoy.month)
                    ] if not pagos_cred.empty else pd.DataFrame()
                    
                    if pagos_mes_actual.empty:
                        en_mora = True

            interes_pend = 0.0 # En cuotas fijas el interés está diluido en la cuota
            deuda_total = cap_pend
            estado = "En mora" if en_mora else "Al día"
            deuda_vencida = valor_cuota_est if en_mora else 0.0

        registros.append({
            'credito_id': c_id,
            'cliente_id': cli_id,
            'tipo_interes': modalidad,
            'capital_inicial': cap_ini,
            'capital_pagado': cap_pagado,
            'capital_pendiente': cap_pend,
            'interes_pendiente': interes_pend,
            'deuda_total_pendiente': deuda_total,
            'deuda_vencida': deuda_vencida,
            'estado': estado
        })
        
    return pd.DataFrame(registros)
