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
    try:
        if isinstance(valor, (int, float)):
            return f"{valor:.2f}".replace('.', ',')
        return str(valor).replace(".", ",")
    except:
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
        return 0.0
    limpio = str(valor).replace(" ", "").replace("\xa0", "")
    # Convierte formatos tipo "1.234.567,89" o "1 234 567,89" -> "1234567.89"
    limpio = limpio.replace(".", "").replace(",", ".")
    try:
        return float(limpio)
    except:
        try:
            return float(str(valor).replace(",", ".").replace(" ", ""))
        except:
            return 0.0

def convertir_fecha_excel(fecha_str):
    if fecha_str:
        try:
            return datetime.strptime(fecha_str, "%d-%m-%Y")
        except ValueError:
            pass
    return None

def extraer_datos_xml_en_memoria(xml_files, numero_receptor_filtro):
    wb = openpyxl.Workbook()

    # HOJA facturas_detalladas
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

    # HOJA facturas_resumidas
    ws_resumidas = wb.create_sheet(title="facturas_resumidas")
    headers_resumidas = [
        "Consecutivo","Detalle","Fecha","Código Moneda","Subtotal",
        "Total Descuentos","Total Impuesto","Otros Cargos","Número Receptor","Total Comprobante"
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
            otros_cargos = formatear_numero(root.find('OtrosCargos/MontoCargo').text) if root.find('OtrosCargos/MontoCargo') is not None else 0
            codigo_moneda = root.find('ResumenFactura/CodigoTipoMoneda/CodigoMoneda').text if root.find('ResumenFactura/CodigoTipoMoneda/CodigoMoneda') is not None else ""

            detalles_servicio = root.find('DetalleServicio')
            detalle_texto = ""
            subtotal_factura = 0.0

            # ----------- NUEVA LÓGICA: procesar líneas y acumular subtotal aquí mismo -----------
            if detalles_servicio is not None:
                lineas_detalle = detalles_servicio.findall('LineaDetalle')
                detalle_texto = "; ".join([linea.find('Detalle').text if linea.find('Detalle') is not None else "" for linea in lineas_detalle])

                for linea in lineas_detalle:
                    codigo_cabys = linea.find('Codigo').text if linea.find('Codigo') is not None else ""
                    detalle = linea.find('Detalle').text if linea.find('Detalle') is not None else ""
                    cantidad = convertir_numero(linea.find('Cantidad').text) if linea.find('Cantidad') is not None else 0
                    precio_unitario = convertir_numero(linea.find('PrecioUnitario').text) if linea.find('PrecioUnitario') is not None else 0
                    monto_total_linea = convertir_numero(linea.find('MontoTotal').text) if linea.find('MontoTotal') is not None else 0
                    monto_descuento_linea = convertir_numero(linea.find('Descuento/MontoDescuento').text) if linea.find('Descuento/MontoDescuento') is not None else 0
                    # convertimos el SubTotal UNA única vez y lo usamos tanto para la fila detallada como para la suma
                    subtotal_linea = convertir_numero(linea.find('SubTotal').text) if linea.find('SubTotal') is not None else 0
                    subtotal_factura += subtotal_linea  # <-- acumulación aquí
                    impuesto = linea.find('Impuesto')
                    tarifa_linea = convertir_numero(impuesto.find('Tarifa').text) if impuesto is not None and impuesto.find('Tarifa') is not None else 0
                    monto_impuesto_linea = convertir_numero(impuesto.find('Monto').text) if impuesto is not None and impuesto.find('Monto') is not None else 0
                    impuesto_neto_linea = convertir_numero(linea.find('ImpuestoNeto').text) if linea.find('ImpuestoNeto') is not None else 0
                    tipo_cambio = convertir_numero(root.find('ResumenFactura/CodigoTipoMoneda/TipoCambio').text) if root.find('ResumenFactura/CodigoTipoMoneda/TipoCambio') is not None else 0
                    total_gravado = convertir_numero(root.find('ResumenFactura/TotalGravado').text) if root.find('ResumenFactura/TotalGravado') is not None else 0

                    fila_detallada = [
                        clave, consecutivo, convertir_fecha_excel(fecha), nombre_emisor, numero_emisor,
                        nombre_receptor, numero_receptor, codigo_cabys, detalle,
                        cantidad, precio_unitario,
                        monto_total_linea, monto_descuento_linea,
                        subtotal_linea, tarifa_linea,
                        monto_impuesto_linea, impuesto_neto_linea,
                        codigo_moneda, tipo_cambio,
                        total_gravado, convertir_numero(total_exento),
                        convertir_numero(total_exonerado), convertir_numero(total_venta),
                        convertir_numero(total_descuentos), convertir_numero(total_venta_neta),
                        convertir_numero(total_impuesto), convertir_numero(total_comprobante),
                        convertir_numero(otros_cargos), filename, tipo_documento
                    ]
                    ws_detalladas.append(fila_detallada)

            # Ahora subtotal_factura es la suma exacta de los subtotal_linea (numéricos)
            try:
                subtotal_factura = float(subtotal_factura)
            except:
                subtotal_factura = 0.0
            subtotal_factura = round(subtotal_factura, 2)

            # facturas_resumidas (usamos el subtotal ya calculado)
            fila_resumida = [
                consecutivo,
                detalle_texto,
                convertir_fecha_excel(fecha),
                codigo_moneda,
                subtotal_factura,
                convertir_numero(total_descuentos),
                convertir_numero(total_impuesto),
                convertir_numero(otros_cargos),
                numero_receptor,
                convertir_numero(total_comprobante)
            ]
            ws_resumidas.append(fila_resumida)

            # Forzamos valor y formato en Subtotal (columna 5)
            fila_actual = ws_resumidas.max_row
            cel_subtotal = ws_resumidas.cell(row=fila_actual, column=5)
            cel_subtotal.value = float(subtotal_factura)
            cel_subtotal.number_format = '#,##0.00'

            # Formateo adicional por fila
            for col_idx in [6, 7, 8, 10]:
                cel = ws_resumidas.cell(row=fila_actual, column=col_idx)
                if isinstance(cel.value, (int, float)):
                    cel.number_format = '#,##0.00'

        except Exception as e:
            flash(f"Error al procesar '{filename}': {e}", 'error')

    # Post-procesamiento robusto: limpieza y formateo de la columna Subtotal y otras numéricas
    if ws_resumidas.max_row >= 2:
        for row_idx in range(2, ws_resumidas.max_row + 1):
            cel = ws_resumidas.cell(row=row_idx, column=5)
            val = cel.value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue
            if isinstance(val, str):
                limpio = val.replace(" ", "").replace("\xa0", "")
                limpio = limpio.replace(".", "").replace(",", ".")
                try:
                    num = float(limpio)
                except:
                    try:
                        num = float(val.replace(",", ".").replace(" ", ""))
                    except:
                        num = 0.0
            elif isinstance(val, (int, float)):
                num = float(val)
            else:
                try:
                    num = float(str(val).replace(",", "."))
                except:
                    num = 0.0
            cel.value = round(num, 2)
            cel.number_format = '#,##0.00'

            for col_idx in [6, 7, 8, 10]:
                c = ws_resumidas.cell(row=row_idx, column=col_idx)
                v = c.value
                if v is None or (isinstance(v, str) and v.strip() == ""):
                    continue
                if isinstance(v, str):
                    lv = v.replace(" ", "").replace("\xa0", "")
                    lv = lv.replace(".", "").replace(",", ".")
                    try:
                        nv = float(lv)
                    except:
                        try:
                            nv = float(v.replace(",", ".").replace(" ", ""))
                        except:
                            nv = None
                elif isinstance(v, (int, float)):
                    nv = float(v)
                else:
                    try:
                        nv = float(str(v).replace(",", "."))
                    except:
                        nv = None
                if nv is not None:
                    c.value = round(nv, 2)
                    c.number_format = '#,##0.00'

    # Formato facturas_detalladas (colores y chequeo receptor)
    fill_celeste = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    fill_rojo = PatternFill(start_color="FFAAAA", end_color="FFAAAA", fill_type="solid")
    col_azul = ["B","C","I","P","V","X","Y","AC"]
    for col in col_azul:
        for cell in ws_detalladas[col]:
            cell.fill = fill_celeste
    for cell in ws_detalladas["G"][1:]:
        if cell.value and numero_receptor_filtro and str(cell.value) != str(numero_receptor_filtro):
            cell.fill = fill_rojo

    # Formato facturas_resumidas (colores)
    col_azul_res = [1, 3, 5, 6, 7, 8, 10]
    for col_idx in col_azul_res:
        for cell in list(ws_resumidas.columns)[col_idx-1]:
            cell.fill = fill_celeste

    # Respaldo: aplicar formato numérico fila por fila (por si Excel ignora alguno)
    for fila in ws_resumidas.iter_rows(min_row=2):
        column_indices_to_format = [4, 5, 6, 7, 9]  # 0-based indices dentro de 'fila'
        for col_index_to_format in column_indices_to_format:
            try:
                cell_to_format = fila[col_index_to_format]
            except IndexError:
                continue
            if isinstance(cell_to_format.value, (int, float)):
                cell_to_format.number_format = '#,##0.00'

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

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
