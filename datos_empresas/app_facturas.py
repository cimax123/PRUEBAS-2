import streamlit as st
import pandas as pd
import openpyxl
import io
import re

# --- Funciones de Limpieza y Utilidad ---

def clean_text(text):
    """Normaliza texto: mayúsculas y strip."""
    if text:
        return str(text).strip().upper()
    return ""

def get_float(s):
    """Convierte a float limpiando símbolos no numéricos."""
    if s is None: return 0.0
    try:
        s_clean = re.sub(r'[^\d.,-]', '', str(s))
        if not s_clean: return 0.0
        s_clean = s_clean.replace(',', '')
        return float(s_clean)
    except ValueError:
        return 0.0

def parse_month(text):
    """Convierte meses texto a número."""
    text = clean_text(text)
    months = {
        'ENERO': '01', 'JANUARY': '01', 'JAN': '01',
        'FEBRERO': '02', 'FEBRUARY': '02', 'FEB': '02',
        'MARZO': '03', 'MARCH': '03', 'MAR': '03',
        'ABRIL': '04', 'APRIL': '04', 'APR': '04',
        'MAYO': '05', 'MAY': '05',
        'JUNIO': '06', 'JUNE': '06', 'JUN': '06',
        'JULIO': '07', 'JULY': '07', 'JUL': '07',
        'AGOSTO': '08', 'AUGUST': '08', 'AUG': '08',
        'SEPTIEMBRE': '09', 'SEPTEMBER': '09', 'SEP': '09',
        'OCTUBRE': '10', 'OCTOBER': '10', 'OCT': '10',
        'NOVIEMBRE': '11', 'NOVEMBER': '11', 'NOV': '11',
        'DICIEMBRE': '12', 'DECEMBER': '12', 'DEC': '12'
    }
    for k, v in months.items():
        if k in text: return v
    if text.isdigit() and len(text) == 1: return f"0{text}"
    return text if text.isdigit() else None

# --- Lógica de Búsqueda Estructural ---

def find_label_cell(sheet, keywords):
    """
    Busca la coordenada de una etiqueta en la hoja.
    Retorna (row, col) de la primera coincidencia.
    """
    # Limitamos búsqueda a primeras 150 filas para cabeceras/footers
    for row in sheet.iter_rows(min_row=1, max_row=150, max_col=50):
        for cell in row:
            val = clean_text(cell.value)
            if not val: continue
            
            for k in keywords:
                # Búsqueda flexible: la celda contiene la keyword
                if k in val:
                    return cell.row, cell.column
    return None

def extract_value_near_label(sheet, label_row, label_col, look_down=True, look_right=True, max_steps=10):
    """
    Dada la posición de una etiqueta, busca el valor asociado saltando vacíos.
    Prioriza búsqueda según argumentos.
    """
    directions = []
    if look_down: directions.append('below')
    if look_right: directions.append('right')
    
    for direction in directions:
        for step in range(1, max_steps + 1):
            r, c = label_row, label_col
            if direction == 'below': r += step
            elif direction == 'right': c += step
            
            try:
                cell = sheet.cell(row=r, column=c)
                val = cell.value
                val_clean = clean_text(val)
                
                if val_clean:
                    # Evitar tomar otra etiqueta como valor (ej: encontrar "FECHA:" buscando "CLIENTE:")
                    # Si el valor termina en dos puntos o parece una etiqueta común, ignorar o detener
                    # Heurística simple: Si contiene otra keyword estructural, saltar?
                    # Por ahora, asumimos que es el dato.
                    return val
            except:
                pass
    return None

def extract_multiline_text(sheet, start_row, col, max_lines=6):
    """
    Concatena texto de varias filas consecutivas (para Observaciones largas).
    Se detiene si encuentra una celda vacía o una nueva etiqueta evidente.
    """
    lines = []
    empty_count = 0
    
    for r in range(start_row, start_row + max_lines):
        try:
            # Ampliamos un poco el ancho (col y col+1) por si está combinado o desplazado
            # Priorizamos la columna alineada, luego la siguiente
            val = sheet.cell(row=r, column=col).value
            if not val:
                 val = sheet.cell(row=r, column=col+1).value # Intento columna vecina derecha
            
            val_clean = clean_text(val)
            
            if not val_clean:
                empty_count += 1
                if empty_count >= 2: break # 2 líneas vacías = fin del bloque
                continue
            else:
                empty_count = 0
                
            # Filtro de parada: Si parece un footer o nueva sección
            if any(x in val_clean for x in ["TOTAL", "PAGE", "FIRMA", "SIGNATURE", "GRACIAS"]):
                break
                
            lines.append(str(val).strip())
        except:
            continue
            
    return " ".join(lines) if lines else None

# --- Procesamiento de Archivo ---

def process_file(uploaded_file):
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active
    except Exception as e:
        return [], f"Error leyendo Excel: {str(e)}"

    header_data = {}

    # 1. CLIENTE
    lbl_coords = find_label_cell(sheet, ['CLIENTE', 'CUSTOMER', 'CONSIGNEE', 'SOLD TO'])
    if lbl_coords:
        header_data['Cliente'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1])
    else:
        header_data['Cliente'] = None

    # 2. EXP / FACTURA
    lbl_coords = find_label_cell(sheet, ['EXP', 'INVOICE NO', 'FACTURA N']) # Keywords
    if lbl_coords:
        header_data['Num_Exp'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1])
    else:
        header_data['Num_Exp'] = None

    # 3. FECHA (Buscamos componentes o fecha completa)
    # Intento 1: Año/Mes/Dia separados
    cy = find_label_cell(sheet, ['AÑO', 'YEAR'])
    cm = find_label_cell(sheet, ['MES', 'MONTH'])
    cd = find_label_cell(sheet, ['DIA', 'DAY'])
    
    val_y = extract_value_near_label(sheet, cy[0], cy[1]) if cy else None
    val_m = extract_value_near_label(sheet, cm[0], cm[1]) if cm else None
    val_d = extract_value_near_label(sheet, cd[0], cd[1]) if cd else None
    
    if val_y and val_m and val_d:
        m_num = parse_month(val_m)
        header_data['Fecha'] = f"{val_d}/{m_num}/{val_y}"
    else:
        # Intento 2: Etiqueta "FECHA" única
        c_date = find_label_cell(sheet, ['FECHA', 'DATE'])
        if c_date:
            raw_date = extract_value_near_label(sheet, c_date[0], c_date[1])
            if hasattr(raw_date, 'strftime'):
                header_data['Fecha'] = raw_date.strftime('%d/%m/%Y')
            else:
                header_data['Fecha'] = raw_date
        else:
            header_data['Fecha'] = None

    # 4. CONDICIÓN DE VENTA (Texto libre bajo etiqueta)
    # Busca la etiqueta y toma lo que haya. No busca palabras específicas.
    lbl_coords = find_label_cell(sheet, ['CONDICION DE VENTA', 'TERMS OF SALE', 'DELIVERY TERMS', 'INCOTERM'])
    if lbl_coords:
        # Priorizamos mirar ABAJO para textos largos, luego a la DERECHA
        header_data['Condicion_Venta'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1], look_down=True, look_right=True)
    else:
        header_data['Condicion_Venta'] = None

    # 5. FORMA DE PAGO
    lbl_coords = find_label_cell(sheet, ['FORMA DE PAGO', 'PAYMENT TERMS'])
    if lbl_coords:
        header_data['Forma_Pago'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1])
    else:
        header_data['Forma_Pago'] = None

    # 6. PUERTOS
    c_emb = find_label_cell(sheet, ['PUERTO DE EMBARQUE', 'LOADING PORT', 'POL'])
    header_data['Puerto_Emb'] = extract_value_near_label(sheet, c_emb[0], c_emb[1]) if c_emb else None
    
    c_dest = find_label_cell(sheet, ['PUERTO DESTINO', 'DISCHARGING PORT', 'DESTINATION', 'POD'])
    header_data['Puerto_Dest'] = extract_value_near_label(sheet, c_dest[0], c_dest[1]) if c_dest else None

    # 7. MONEDA e INCOTERM (Códigos)
    # Estos sí son estandarizados (USD, FOB), así que los escaneamos globalmente por seguridad
    detected_codes = {'Moneda': None, 'Incoterm': None}
    incoterm_list = ['FOB', 'CIF', 'CFR', 'EXW', 'FCA', 'DDP', 'DAP']
    currency_map = {'DÓLAR': 'USD', 'DOLAR': 'USD', 'USD': 'USD', 'EURO': 'EUR'}
    
    for row in sheet.iter_rows(min_row=1, max_row=100):
        for cell in row:
            v = clean_text(cell.value)
            if not v: continue
            
            if not detected_codes['Incoterm']:
                for inc in incoterm_list:
                    if re.search(rf"\b{inc}\b", v): detected_codes['Incoterm'] = inc
            
            if not detected_codes['Moneda']:
                for cur, code in currency_map.items():
                    if cur in v: detected_codes['Moneda'] = code

    header_data['Moneda'] = detected_codes['Moneda']
    header_data['Incoterm'] = detected_codes['Incoterm']

    # --- EXTRACCIÓN DE PRODUCTOS ---
    products = []
    col_map = {}
    header_row = None
    last_row_table = 20
    
    # Keywords de columnas
    kw_desc = ['DESCRIPCION', 'DESCRIPTION', 'MERCADERIA']
    kw_qty = ['CANTIDAD', 'QUANTITY', 'QTTY', 'PCS', 'BOXES']
    kw_price = ['PRECIO', 'PRICE', 'UNIT']
    kw_total = ['TOTAL']

    # 1. Encontrar cabecera de tabla
    for row in sheet.iter_rows(min_row=1, max_row=100):
        row_txt = [clean_text(c.value) for c in row]
        # Debe tener descripcion Y cantidad
        if any(any(k in t for k in kw_desc) for t in row_txt) and any(any(k in t for k in kw_qty) for t in row_txt):
            header_row = row[0].row
            # Mapear columnas
            for cell in row:
                v = clean_text(cell.value)
                if not v: continue
                if any(k in v for k in kw_desc): col_map['Descripcion'] = cell.column
                elif any(k in v for k in kw_qty): col_map['Cantidad'] = cell.column
                elif any(k in v for k in kw_price): col_map['Precio Unitario'] = cell.column
                elif any(k in v for k in kw_total) and 'TOTAL CASES' not in v and 'FOB' not in v: 
                    col_map['Total Linea'] = cell.column
            break

    # 2. Leer filas
    if header_row and 'Descripcion' in col_map:
        curr = header_row + 1
        empty_streak = 0
        while curr <= sheet.max_row:
            desc_val = sheet.cell(row=curr, column=col_map['Descripcion']).value
            desc_clean = clean_text(desc_val)
            
            # Si encontramos palabra TOTAL al inicio de la descripción, asumimos fin de tabla
            if desc_clean.startswith("TOTAL") or " TOTAL" in desc_clean:
                last_row_table = curr
                break
            
            # Datos
            qty = sheet.cell(row=curr, column=col_map['Cantidad']).value if 'Cantidad' in col_map else 0
            price = sheet.cell(row=curr, column=col_map['Precio Unitario']).value if 'Precio Unitario' in col_map else 0
            total = sheet.cell(row=curr, column=col_map['Total Linea']).value if 'Total Linea' in col_map else 0
            
            q_num = get_float(qty)
            p_num = get_float(price)
            t_num = get_float(total)
            
            # Filtro: Debe tener precio > 0 y descripción para ser producto
            if p_num > 0 and desc_clean:
                prod = {
                    'Descripcion': desc_val,
                    'Cantidad': q_num,
                    'Precio Unitario': p_num,
                    'Total Linea': t_num if t_num > 0 else (q_num * p_num)
                }
                products.append(prod)
                empty_streak = 0
                last_row_table = curr # Actualizamos última fila válida
            else:
                if not desc_clean:
                    empty_streak += 1
                    if empty_streak > 15: # Fin de tabla por vacíos
                        last_row_table = curr - 15
                        break
            curr += 1
    else:
        last_row_table = 20 # Fallback

    # --- OBSERVACIONES (Lógica Multilínea + Ubicación Dinámica) ---
    # Buscamos la etiqueta 'OBSERVACIONES' en TODA la hoja, pero enfocándonos en el footer si es posible
    # Si la encontramos, usamos extract_multiline_text para tomar todo el bloque
    
    obs_label_pos = find_label_cell(sheet, ['OBSERVACIONES', 'REMARKS', 'NOTAS', 'GLOSA', 'COMENTARIOS'])
    
    if obs_label_pos:
        r_obs, c_obs = obs_label_pos
        # Intentamos obtener texto multilínea DEBAJO de la etiqueta
        obs_text = extract_multiline_text(sheet, start_row=r_obs + 1, col=c_obs)
        
        # Si no hubo nada abajo, miramos a la derecha (una sola celda o bloque)
        if not obs_text:
             obs_text = extract_value_near_label(sheet, r_obs, c_obs, look_down=False, look_right=True)
             
        header_data['Observaciones'] = obs_text
    else:
        header_data['Observaciones'] = None

    # --- Unificación ---
    final_rows = []
    if products:
        for p in products:
            final_rows.append({'Archivo': uploaded_file.name, **header_data, **p})
    else:
        final_rows.append({'Archivo': uploaded_file.name, **header_data})
        
    return final_rows, None

# --- Interfaz Gráfica ---

st.set_page_config(page_title="Extractor V11 Genérico", layout="wide")
st.title("📄 Extractor de Facturas V11 (Estructural)")
st.info("Búsqueda 100% basada en estructura: Etiquetas + Posición. Detecta observaciones multilínea y textos variables.")

uploaded_files = st.file_uploader("Archivos Excel (.xlsx)", type=['xlsx'], accept_multiple_files=True)

if uploaded_files and st.button("Procesar Archivos"):
    all_data = []
    for file in uploaded_files:
        rows, err = process_file(file)
        if rows: all_data.extend(rows)
        if err: st.error(f"{file.name}: {err}")
        
    if all_data:
        df = pd.DataFrame(all_data)
        
        # Ordenar columnas
        cols_order = ['Archivo', 'Cliente', 'Num_Exp', 'Fecha', 'Condicion_Venta', 'Incoterm', 'Forma_Pago', 
                      'Moneda', 'Puerto_Emb', 'Puerto_Dest', 'Cantidad', 'Descripcion', 
                      'Precio Unitario', 'Total Linea', 'Observaciones']
        
        final_cols = [c for c in cols_order if c in df.columns] + [c for c in df.columns if c not in cols_order]
        df = df[final_cols]
        
        st.success("Procesamiento completado.")
        st.dataframe(df, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button("Descargar Excel", buffer.getvalue(), "facturas_v11.xlsx")
