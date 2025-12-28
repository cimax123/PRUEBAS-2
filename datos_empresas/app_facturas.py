import streamlit as st
import pandas as pd
import openpyxl
import io
import re

# --- Funciones de Utilidad ---

def clean_text(text):
    """Normaliza texto: mayúsculas y strip."""
    if text:
        return str(text).strip().upper()
    return ""

def get_float(s):
    """Intenta convertir a float, retorna 0.0 si falla."""
    if s is None: return 0.0
    try:
        # Limpiar caracteres no numéricos excepto punto y coma
        # Asumiendo formato decimal con punto o coma, esto puede requerir ajuste según región
        # Aquí intentamos simple: quitar comas de miles y parsear
        val_str = str(s).replace(',', '').strip()
        return float(val_str)
    except ValueError:
        return 0.0

def parse_month(text):
    """Convierte meses en texto (ES/EN) a número."""
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
        if k in text:
            return v
    if text.isdigit() and len(text) == 1:
        return f"0{text}"
    return text if text.isdigit() else None

def scan_sheet_for_specific_values(sheet):
    """
    Busca valores específicos (Moneda, Incoterm) escaneando toda la hoja.
    """
    detected = {
        'Moneda': None,
        'Incoterm': None
    }
    
    currency_map = {'DÓLAR': 'USD', 'DOLAR': 'USD', 'USD': 'USD', 'EURO': 'EUR', 'EUR': 'EUR'}
    # Incoterms de 3 letras
    incoterm_list = ['FOB', 'CIF', 'CFR', 'EXW', 'FCA', 'DDP', 'DAP', 'CPT', 'CIP']

    for row in sheet.iter_rows(min_row=1, max_row=100):
        for cell in row:
            val = clean_text(cell.value)
            if not val: continue
            
            # 1. Moneda
            if not detected['Moneda']:
                for k, v in currency_map.items():
                    if k == val or f" {k}" in val or f"{k} " in val:
                        detected['Moneda'] = v
                        break
            
            # 2. Incoterm (Código puro)
            if not detected['Incoterm']:
                for inc in incoterm_list:
                    # Buscamos palabra exacta "FOB" o "TOTAL FOB"
                    if re.search(rf"\b{inc}\b", val):
                        detected['Incoterm'] = inc
                        break
    return detected

def get_data_near_label(sheet, start_row, start_col, search_directions=['below', 'right'], max_steps=10):
    """
    Busca dato no vacío cerca de una coordenada.
    Aumentado max_steps a 10 para saltar columnas vacías ocultas.
    """
    for direction in search_directions:
        for step in range(1, max_steps + 1):
            target_row = start_row
            target_col = start_col
            
            if direction == 'below':
                target_row += step
            elif direction == 'right':
                target_col += step
            
            try:
                cell = sheet.cell(row=target_row, column=target_col)
                val = cell.value
                val_clean = clean_text(val)
                # Ignorar si encontramos otra etiqueta conocida (ej: encontrar 'FECHA' buscando 'CLIENTE')
                if val is not None and val_clean != "":
                    return val
            except:
                pass
    return None

def find_coords(sheet, keywords, exact_match=False):
    """Devuelve (row, col) de la primera coincidencia de keyword."""
    for row in sheet.iter_rows(min_row=1, max_row=100, max_col=50):
        for cell in row:
            val = clean_text(cell.value)
            if not val: continue
            
            for k in keywords:
                if exact_match:
                    if k == val or re.search(rf"\b{re.escape(k)}\b", val):
                        return cell.row, cell.column
                else:
                    if k in val:
                        return cell.row, cell.column
    return None

# --- Procesamiento Principal ---

def process_file(uploaded_file):
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active
    except Exception as e:
        return [], f"Error leyendo Excel: {str(e)}"

    header_data = {}
    
    # 1. Búsquedas Estructurales
    
    # -- CLIENTE --
    coord = find_coords(sheet, ['CLIENTE', 'CUSTOMER', 'CONSIGNEE', 'SOLD TO'])
    header_data['Cliente'] = get_data_near_label(sheet, coord[0], coord[1], ['below', 'right']) if coord else None
    
    # -- EXP (Búsqueda estricta) --
    coord = find_coords(sheet, ['EXP', 'INVOICE NO', 'FACTURA N'], exact_match=True)
    header_data['Num_Exp'] = get_data_near_label(sheet, coord[0], coord[1], ['below', 'right']) if coord else None

    # -- FECHA --
    coord_year = find_coords(sheet, ['AÑO', 'YEAR'], exact_match=True)
    coord_month = find_coords(sheet, ['MES', 'MONTH'], exact_match=True)
    coord_day = find_coords(sheet, ['DIA', 'DAY'], exact_match=True)
    
    y = get_data_near_label(sheet, coord_year[0], coord_year[1], ['below', 'right']) if coord_year else None
    m = get_data_near_label(sheet, coord_month[0], coord_month[1], ['below', 'right']) if coord_month else None
    d = get_data_near_label(sheet, coord_day[0], coord_day[1], ['below', 'right']) if coord_day else None
    
    if y and m and d:
        m_num = parse_month(m)
        header_data['Fecha'] = f"{d}/{m_num}/{y}"
    else:
        coord_date = find_coords(sheet, ['FECHA', 'DATE'])
        if coord_date:
            val = get_data_near_label(sheet, coord_date[0], coord_date[1], ['right', 'below'])
            if hasattr(val, 'strftime'): 
                header_data['Fecha'] = val.strftime('%d/%m/%Y')
            else:
                header_data['Fecha'] = val
        else:
            header_data['Fecha'] = None

    # -- OBSERVACIONES --
    # Ampliamos keywords y aumentamos radio de búsqueda
    obs_keywords = ['OBSERVACIONES', 'REMARKS', 'NOTAS', 'GLOSA', 'COMENTARIOS', 'SPECIAL INSTRUCTIONS']
    coord = find_coords(sheet, obs_keywords)
    header_data['Observaciones'] = get_data_near_label(sheet, coord[0], coord[1], ['below', 'right'], max_steps=6) if coord else None

    # -- PUERTOS --
    coord = find_coords(sheet, ['PUERTO DE EMBARQUE', 'LOADING PORT', 'POL'])
    header_data['Puerto_Emb'] = get_data_near_label(sheet, coord[0], coord[1], ['below', 'right']) if coord else None
    
    coord = find_coords(sheet, ['PUERTO DESTINO', 'DISCHARGING PORT', 'DESTINATION', 'POD'])
    header_data['Puerto_Dest'] = get_data_near_label(sheet, coord[0], coord[1], ['below', 'right']) if coord else None

    # -- CONDICION DE VENTA (Texto descriptivo) --
    # Buscamos "CONDICION DE VENTA" pero NO "FORMA DE PAGO"
    coord = find_coords(sheet, ['CONDICION DE VENTA', 'TERMS OF SALE', 'DELIVERY TERMS', 'INCOTERM'])
    header_data['Condicion_Venta'] = get_data_near_label(sheet, coord[0], coord[1], ['below', 'right']) if coord else None

    # -- FORMA DE PAGO --
    coord = find_coords(sheet, ['PAYMENT TERMS', 'FORMA DE PAGO'])
    header_data['Forma_Pago'] = get_data_near_label(sheet, coord[0], coord[1], ['below', 'right']) if coord else None

    # 2. Escaneo Directo (Moneda e Incoterm Código)
    scanned = scan_sheet_for_specific_values(sheet)
    header_data['Moneda'] = scanned['Moneda']
    header_data['Incoterm'] = scanned['Incoterm'] 

    # 3. Extracción de Productos (Lógica "Solo Precios Reales")
    products = []
    
    col_map = {}
    header_row = None
    
    desc_keywords = ['DESCRIPCION', 'DESCRIPTION', 'MERCADERIA']
    qty_keywords = ['CANTIDAD', 'QUANTITY', 'QTTY', 'PCS', 'BOXES']
    price_keywords = ['PRECIO', 'PRICE', 'UNIT']
    total_keywords = ['TOTAL']

    # Buscar cabecera de tabla
    for row in sheet.iter_rows(min_row=1, max_row=100):
        row_txt = [clean_text(c.value) for c in row]
        has_desc = any(any(k in txt for k in desc_keywords) for txt in row_txt if txt)
        has_qty = any(any(k in txt for k in qty_keywords) for txt in row_txt if txt)
        
        if has_desc and has_qty:
            header_row = row[0].row
            for cell in row:
                val = clean_text(cell.value)
                if not val: continue
                if any(k in val for k in desc_keywords): col_map['Descripcion'] = cell.column
                elif any(k in val for k in qty_keywords): col_map['Cantidad'] = cell.column
                elif any(k in val for k in price_keywords): col_map['Precio Unitario'] = cell.column
                elif any(k in val for k in total_keywords) and 'TOTAL CASES' not in val and 'TOTAL FOB' not in val: 
                    col_map['Total Linea'] = cell.column
            break
            
    if header_row and 'Descripcion' in col_map:
        curr = header_row + 1
        empty_streak = 0
        while curr <= sheet.max_row:
            desc_val = sheet.cell(row=curr, column=col_map['Descripcion']).value
            desc_clean = clean_text(desc_val)
            
            # Criterio de parada: Totales explícitos en descripción
            if desc_clean.startswith("TOTAL") or " TOTAL" in desc_clean:
                # Ojo: A veces "TOTAL CASES" aparece. Cortamos aquí.
                break
            
            # Extracción
            qty = sheet.cell(row=curr, column=col_map['Cantidad']).value if 'Cantidad' in col_map else None
            price = sheet.cell(row=curr, column=col_map['Precio Unitario']).value if 'Precio Unitario' in col_map else None
            total = sheet.cell(row=curr, column=col_map['Total Linea']).value if 'Total Linea' in col_map else None
            
            # CONVERSIÓN A NÚMEROS (Crucial para filtrar basura)
            qty_num = get_float(qty)
            price_num = get_float(price)
            total_num = get_float(total)

            # FILTRO ESTRICTO:
            # Una fila es válida SOLO si tiene PRECIO > 0.
            # (Muchas veces hay filas con cantidades totales de cajas pero sin precio, o textos largos).
            is_valid_product = (price_num > 0)
            
            if not desc_clean:
                empty_streak += 1
                if empty_streak > 15: break # Margen amplio por si hay saltos de página
            else:
                empty_streak = 0
                if is_valid_product:
                    prod = {
                        'Descripcion': desc_val,
                        'Cantidad': qty_num,
                        'Precio Unitario': price_num,
                        'Total Linea': total_num if total_num > 0 else (qty_num * price_num)
                    }
                    products.append(prod)
                
            curr += 1

    # 4. Unificación
    final_rows = []
    if products:
        for p in products:
            final_rows.append({'Archivo': uploaded_file.name, **header_data, **p})
    else:
        # Si no detectó productos válidos, devolvemos solo la cabecera
        final_rows.append({'Archivo': uploaded_file.name, **header_data})
        
    return final_rows, None

# --- UI ---
st.set_page_config(page_title="Extractor V7", layout="wide")
st.title("📄 Extractor de Facturas V7 (Validación de Precios)")
st.info("Ahora filtra estrictamente filas que no tengan un precio unitario > 0 para evitar filas fantasma.")

uploaded_files = st.file_uploader("Archivos Excel", type=['xlsx'], accept_multiple_files=True)

if uploaded_files and st.button("Procesar"):
    all_data = []
    for file in uploaded_files:
        rows, err = process_file(file)
        if rows: all_data.extend(rows)
        if err: st.error(f"{file.name}: {err}")
        
    if all_data:
        df = pd.DataFrame(all_data)
        
        cols = ['Archivo', 'Cliente', 'Num_Exp', 'Fecha', 'Condicion_Venta', 'Incoterm', 'Forma_Pago', 'Moneda', 
                'Puerto_Emb', 'Puerto_Dest', 'Cantidad', 'Descripcion', 
                'Precio Unitario', 'Total Linea', 'Observaciones']
        
        existing_cols = [c for c in cols if c in df.columns]
        remaining_cols = [c for c in df.columns if c not in existing_cols]
        df = df[existing_cols + remaining_cols]
        
        st.dataframe(df, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button("Descargar Excel Unificado", buffer.getvalue(), "facturas_procesadas_v7.xlsx")
