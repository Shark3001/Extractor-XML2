import os
import io
import xml.etree.ElementTree as ET
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
import openpyxl
from openpyxl.styles import PatternFill

app = Flask(__name__)

app.secret_key = os.getenv("SECRET_KEY_APP_XML", "CAMBIA_ESTA_CLAVE_EN_RENDER")
CORRECT_PASSWORD = os.getenv("APP_PASSWORD", "AFC2024*")

# --- Funciones de conversión ---
def formatear_numero(valor):
    if valor is None:
        return ""
    return str(valor).replace(".", ",")

def formatear_fecha(fecha_str):
    if fecha_str:
        try:
            return datetime.fromisoformat(fecha_str.replace('Z', '+00:00')).strftime('%d-%m-%Y')
        except ValueError:
            return fecha_str
    return ""

def convertir_numero(valor):
    if valor is None or valor == "":
        return 0
    # Reemplaza comas por punto para float
    return float(str(valor).replace(".", "").replace(",", "."))

def convertir_fecha_excel(fecha_str):
    if fecha_str:
        try:
            return datetime.strptime(fecha_str, "%d-%m-%Y")
        except ValueError:
            pass
    return None

def extraer_datos_xml_en_memoria(xml_files, numero_receptor_filtro):
    wb = openpyxl.Workbook()

    # --- HOJA facturas_detalladas ---
    ws_detalladas = wb.active
    ws_detalladas.title = "facturas_detalladas"
    headers_detalladas = [
        "Clave","Consecutivo","Fecha","Nombre Emisor","Número Emisor","Nombre Receptor","Número Receptor",
        "Código Cabys","Detalle","Cantidad","Precio Unitario","Monto Total","Monto Descuento","Subtotal",
        "Tarifa (%)","Monto Impuesto","Impuesto Neto","Código Moneda","Tipo Cambio",
        "Total Gravado","Total Exento","Total Exonerado","Total Venta","Total Descuentos",
        "Total Venta Neta","Total Impuesto","Total Comprobante","Otros Cargos","Archivo","Tipo de Documento"
    ]
    ws_detalladas.append(headers_detalladas)

    # --- HOJA facturas_resumidas ---
    ws_resumidas = wb.create_sheet(title="facturas_resumidas")
    headers_resumidas = [
        "Clave","Consecutivo","Fecha","Nombre Emisor","Número Emisor","Número Receptor",
        "Total Exento","Total Exonerado","Total Venta","Total Descuentos","Total Venta Neta",
        "Total Impuesto","Total Comprobante","Otros Cargos","Archivo","Tipo de Documento"
    ]
    ws_resumidas.append(headers_resumidas)

    for uploaded_file in xml_files:
        filename = uploaded_file.filename
        try:
            tree = ET.parse(uploaded_file)
            root = tree.getroot()
            for elem in root.iter():
                elem.tag = elem.tag.split('}', 1)[-1]

            tipo_documento = root.tag.split(' ')[0]
            if tipo_documento == "MensajeHacienda":
                continue

            clave = root.find('Clave').text if root.find('Clave') is not None else ""
            consecutivo = root.find('NumeroConsecutivo').text if root.find('NumeroConsecutivo') is not None else ""
            fecha = formatear_fecha(root.find('FechaEmision').text) if root.find('FechaEmision') is not None else ""
            nombre_emisor = root.find('Emisor/Nombre').text if root.find('Emisor/Nombre') is not None else ""
            numero_emisor = root.find('Emisor/Identificacion/Numero').text if root.find('Emisor/Identificacion/Numero') is not None else ""
            nombre_receptor = root.find('Receptor/Nombre').text if root.find('Receptor/Nombre') is not None else ""
            numero_receptor = root.find('Receptor/Identificacion/Numero').text if root.find('Receptor/Identificacion/Numero') is not None else ""

            resumen_factura = root.find('ResumenFactura')
            total_venta = formatear_numero(resumen_factura.find('TotalVenta').text) if resumen_factura is not None and resumen_factura.find('TotalVenta') is not None else ""
            total_descuentos = formatear_numero(resumen_factura.find('TotalDescuentos').text) if resumen_factura is not None and resumen_factura.find('TotalDescuentos') is not None else ""
            total_venta_neta = formatear_numero(resumen_factura.find('TotalVentaNeta').text) if resumen_factura is not None and resumen_factura.find('TotalVentaNeta') is not None else ""
            total_exento = formatear_numero(resumen_factura.find('TotalExento').text) if resumen_factura is not None and resumen_factura.find('TotalExento') is not None else ""
            total_exonerado = formatear_numero(resumen_factura.find('TotalExonerado').text) if resumen_factura is not None and resumen_factura.find('TotalExonerado') is not None else ""
            total_impuesto = formatear_numero(resumen_factura.find('TotalImpuesto').text) if resumen_factura is not None and resumen_factura.find('TotalImpuesto') is not None else ""
            total_comprobante = formatear_numero(resumen_factura.find('TotalComprobante').text) if resumen_factura is not None and resumen_factura.find('TotalComprobante') is not None else ""
            otros_cargos = formatear_numero(root.find('OtrosCargos/MontoCargo').text) if root.find('OtrosCargos/MontoCargo') is not None else "0,00"

            detalles_servicio = root.find('DetalleServicio')
            detalle_texto = ""
            if detalles_servicio is not None:
                lineas_detalle = detalles_servicio.findall('LineaDetalle')
                detalle_texto = "; ".join([linea.find('Detalle').text if linea.find('Detalle') is not None else "" for linea in lineas_detalle])

            # --- facturas_detalladas ---
            if detalles_servicio is not None:
                for linea in detalles_servicio.findall('LineaDetalle'):
                    codigo_cabys = linea.find('Codigo').text if linea.find('Codigo') is not None else ""
                    detalle = linea.find('Detalle').text if linea.find('Detalle') is not None else ""
                    cantidad = formatear_numero(linea.find('Cantidad').text) if linea.find('Cantidad') is not None else ""
                    precio_unitario = formatear_numero(linea.find('PrecioUnitario').text) if linea.find('PrecioUnitario') is not None else ""
                    monto_total_linea = formatear_numero(linea.find('MontoTotal').text) if linea.find('MontoTotal') is not None else ""
                    monto_descuento_linea = formatear_numero(linea.find('Descuento/MontoDescuento').text) if linea.find('Descuento/MontoDescuento') is not None else "0,00"
                    subtotal_linea = formatear_numero(linea.find('SubTotal').text) if linea.find('SubTotal') is not None else ""
                    impuesto = linea.find('Impuesto')
                    tarifa_linea = formatear_numero(impuesto.find('Tarifa').text) if impuesto is not None and impuesto.find('Tarifa') is not None else "0,00"
                    monto_impuesto_linea = formatear_numero(impuesto.find('Monto').text) if impuesto is not None and impuesto.find('Monto') is not None else "0,00"
                    impuesto_neto_linea = formatear_numero(linea.find('ImpuestoNeto').text) if linea.find('ImpuestoNeto') is not None else ""
                    codigo_moneda = root.find('ResumenFactura/CodigoTipoMoneda/CodigoMoneda').text if root.find('ResumenFactura/CodigoTipoMoneda/CodigoMoneda') is not None else ""
                    tipo_cambio = formatear_numero(root.find('ResumenFactura/CodigoTipoMoneda/TipoCambio').text) if root.find('ResumenFactura/CodigoTipoMoneda/TipoCambio') is not None else ""
                    total_gravado = formatear_numero(root.find('ResumenFactura/TotalGravado').text) if root.find('ResumenFactura/TotalGravado') is not None else ""
                    total_comprobante_linea = total_comprobante
                    otros_cargos_linea = otros_cargos

                    fila_detallada = [
                        clave,
                        consecutivo,
                        convertir_fecha_excel(fecha),
                        nombre_emisor,
                        numero_emisor,
                        nombre_receptor,
                        numero_receptor,
                        codigo_cabys,
                        detalle,
                        convertir_numero(cantidad),
                        convertir_numero(precio_unitario),
                        convertir_numero(monto_total_linea),
                        convertir_numero(monto_descuento_linea),
                        convertir_numero(subtotal_linea),
                        convertir_numero(tarifa_linea),
                        convertir_numero(monto_impuesto_linea),
                        convertir_numero(impuesto_neto_linea),
                        codigo_moneda,
                        convertir_numero(tipo_cambio),
                        convertir_numero(total_gravado),
                        convertir_numero(total_exento),
                        convertir_numero(total_exonerado),
                        convertir_numero(total_venta),
                        convertir_numero(total_descuentos),
                        convertir_numero(total_venta_neta),
                        convertir_numero(total_impuesto),
                        convertir_numero(total_comprobante_linea),
                        convertir_numero(otros_cargos_linea),
                        filename,
                        tipo_documento
                    ]
                    ws_detalladas.append(fila_detallada)

            # --- facturas_resumidas ---
            fila_resumida = [
                clave,
                consecutivo,
                convertir_fecha_excel(fecha),
                nombre_emisor,
                numero_emisor,
                numero_receptor,
                convertir_numero(total_exento),
                convertir_numero(total_exonerado),
                convertir_numero(total_venta),
                convertir_numero(total_descuentos),
                convertir_numero(total_venta_neta),
                convertir_numero(total_impuesto),
                convertir_numero(total_comprobante),
                convertir_numero(otros_cargos),
                filename,
                tipo_documento
            ]
            ws_resumidas.append(fila_resumida)

        except Exception as e:
            flash(f"Error al procesar '{filename}': {e}", 'error')

    # --- Formato colores facturas_detalladas ---
    fill_celeste = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    fill_rojo = PatternFill(start_color="FFAAAA", end_color="FFAAAA", fill_type="solid")
    col_azul = ["B","C","I","P","V","X","Y","AC"]
    for col in col_azul:
        for cell in ws_detalladas[col]:
            cell.fill = fill_celeste
    for cell in ws_detalladas["G"][1:]:
        if cell.value and numero_receptor_filtro and str(cell.value) != str(numero_receptor_filtro):
            cell.fill = fill_rojo

    # --- Formato colores facturas_resumidas ---
    col_azul_res = [2,3,12,10,11,8,9,13]
    for col_idx in col_azul_res:
        for cell in list(ws_resumidas.columns)[col_idx-1]:
            cell.fill = fill_celeste
    for fila in ws_resumidas.iter_rows(min_row=2):
        cell = fila[5]
        if cell.value and numero_receptor_filtro and str(cell.value) != str(numero_receptor_filtro):
            cell.fill = fill_rojo

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    # Limpiar archivos en memoria
    xml_files.clear()
    return out

# --------- Rutas ---------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        password = request.form.get('password')
        if password == CORRECT_PASSWORD:
            session['logged_in'] = True
            flash('Inicio de sesión exitoso.', 'success')
            return redirect(url_for('index'))
        else:
            flash('Contraseña incorrecta. Inténtalo de nuevo.', 'error')
            return redirect(url_for('login'))
    return render_template('login.html')

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    flash('Has cerrado sesión correctamente.', 'success')
    return redirect(url_for('login'))

@app.route('/upload', methods=['POST'])
def upload_files():
    if not session.get('logged_in'):
        flash('Por favor, inicia sesión para acceder a esta función.', 'error')
        return redirect(url_for('login'))

    if 'xml_files' not in request.files:
        flash('No se subieron archivos.', 'error')
        return redirect(url_for('index'))

    files = request.files.getlist('xml_files')
    if not files or files[0].filename == '':
        flash('No se seleccionó ningún archivo.', 'error')
        return redirect(url_for('index'))

    numero_receptor = request.form.get('numero_receptor')
    if not numero_receptor:
        flash('El número de identificación del receptor es obligatorio.', 'error')
        return redirect(url_for('index'))

    excel_stream = extraer_datos_xml_en_memoria(files, numero_receptor)

    return send_file(
        excel_stream,
        download_name='datos_facturas.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=False)

