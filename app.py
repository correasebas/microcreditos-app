import io
import urllib.parse
import streamlit as st
from weasyprint import HTML

def generar_pdf_comprobante(pago_info):
    """
    Genera un comprobante de pago en formato PDF utilizando HTML + CSS paged media via WeasyPrint.
    """
    html_content = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{
                size: A4;
                margin: 15mm 12mm;
                background-color: #faf8f5;
            }}
            body {{
                font-family: 'Helvetica Neue', Arial, sans-serif;
                margin: 0;
                padding: 0;
                color: #2c3e50;
            }}
            .header {{
                background-color: #1b365d;
                color: #ffffff;
                padding: 22px 20px;
                margin: -15mm -12mm 20px -12mm;
                text-align: center;
            }}
            .header h1 {{
                margin: 0;
                font-size: 20pt;
                letter-spacing: 1px;
            }}
            .header p {{
                margin: 4px 0 0 0;
                font-size: 10pt;
                color: #d1d8e0;
            }}
            .badge {{
                display: inline-block;
                background-color: #27ae60;
                color: white;
                padding: 5px 14px;
                border-radius: 15px;
                font-size: 9.5pt;
                font-weight: bold;
                margin-top: 8px;
            }}
            .container {{
                padding: 5px;
            }}
            .box {{
                background-color: #ffffff;
                border: 1px solid #e1e8ed;
                border-radius: 8px;
                padding: 14px 18px;
                margin-bottom: 18px;
            }}
            .box-title {{
                font-size: 11pt;
                font-weight: bold;
                color: #1b365d;
                border-bottom: 2px solid #3498db;
                padding-bottom: 4px;
                margin-bottom: 10px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
            }}
            td {{
                padding: 6px 4px;
                font-size: 10pt;
            }}
            .label {{
                font-weight: bold;
                color: #576574;
                width: 40%;
            }}
            .value {{
                color: #2c3e50;
                width: 60%;
            }}
            .table-summary {{
                margin-top: 5px;
            }}
            .table-summary th {{
                background-color: #1b365d;
                color: white;
                padding: 8px 10px;
                font-size: 9.5pt;
                text-align: left;
            }}
            .table-summary td {{
                border-bottom: 1px solid #e1e8ed;
                padding: 8px 10px;
            }}
            .total-row {{
                background-color: #f1f2f6;
                font-weight: bold;
            }}
            .footer {{
                margin-top: 25px;
                text-align: center;
                font-size: 8.5pt;
                color: #7f8c8d;
                border-top: 1px solid #dcdde1;
                padding-top: 12px;
            }}
            .highlight-green {{
                color: #27ae60;
                font-weight: bold;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>🤝 Entre Amigos Capital</h1>
            <p>Comprobante Oficial de Recaudo de Pago</p>
            <div class="badge">RECIBO N°: {pago_info['pago_id']}</div>
        </div>

        <div class="container">
            <div class="box">
                <div class="box-title">👤 Datos del Cliente y Crédito</div>
                <table>
                    <tr>
                        <td class="label">Cliente:</td>
                        <td class="value"><strong>{pago_info['cliente_nombre']}</strong> ({pago_info['cliente_id']})</td>
                    </tr>
                    <tr>
                        <td class="label">Número de Crédito:</td>
                        <td class="value">{pago_info['credito_id']}</td>
                    </tr>
                    <tr>
                        <td class="label">Fecha y Hora:</td>
                        <td class="value">{pago_info['fecha_pago']}</td>
                    </tr>
                    <tr>
                        <td class="label">Medio de Pago:</td>
                        <td class="value">{pago_info['medio_pago']}</td>
                    </tr>
                </table>
            </div>

            <div class="box">
                <div class="box-title">💰 Desglose de la Transacción</div>
                <table class="table-summary">
                    <thead>
                        <tr>
                            <th>Concepto</th>
                            <th style="text-align: right;">Monto (COP)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>Abono a Intereses</td>
                            <td style="text-align: right;">${pago_info['pago_interes']:,.0f}</td>
                        </tr>
                        <tr>
                            <td>Abono a Capital</td>
                            <td style="text-align: right;">${pago_info['pago_capital']:,.0f}</td>
                        </tr>
                        <tr class="total-row">
                            <td>TOTAL RECIBIDO ({pago_info['concepto']})</td>
                            <td style="text-align: right; color: #1b365d; font-size: 11pt;">${pago_info['valor_pago']:,.0f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="box">
                <div class="box-title">📊 Estado Actualizado de la Deuda</div>
                <table>
                    <tr>
                        <td class="label">Capital Pendiente Actual:</td>
                        <td class="value">${pago_info['nuevo_cap_pend']:,.0f} COP</td>
                    </tr>
                    <tr>
                        <td class="label">Intereses Pendientes Actuales:</td>
                        <td class="value">${pago_info['nuevo_int_pend']:,.0f} COP</td>
                    </tr>
                    <tr>
                        <td class="label">Deuda Total Pendiente:</td>
                        <td class="value"><strong style="color: #c0392b;">${pago_info['nueva_deuda_total']:,.0f} COP</strong></td>
                    </tr>
                    <tr>
                        <td class="label">Estado del Crédito:</td>
                        <td class="value"><span class="highlight-green">{pago_info['nuevo_estado']}</span></td>
                    </tr>
                    {f"<tr><td class='label'>Observaciones:</td><td class='value'>{pago_info['observaciones']}</td></tr>" if pago_info.get('observaciones') else ""}
                </table>
            </div>

            <div class="footer">
                <p>Gracias por mantener tu crédito al día. Este documento sirve como soporte oficial de recaudo.</p>
                <p><strong>Entre Amigos Capital</strong> — Créditos justos sobre la base de la confianza.</p>
            </div>
        </div>
    </body>
    </html>
    """
    return HTML(string=html_content).write_pdf()

def construir_link_whatsapp(telefono, mensaje):
    """Limpia el número de teléfono y genera el enlace directo a WhatsApp Web/App."""
    num_limpio = "".join(filter(str.isdigit, str(telefono)))
    if len(num_limpio) == 10 and not num_limpio.startswith("57"):
        num_limpio = "57" + num_limpio  # Prefijo de Colombia
    msg_encoded = urllib.parse.quote(mensaje)
    return f"https://wa.me/{num_limpio}?text={msg_encoded}"
