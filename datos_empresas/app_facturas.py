import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

# --- Funciones de Limpieza y Utilidad ---

def normalize_text(text):
    """Normaliza texto para comparaciones."""
    if not text: return ""
    text = str(text).upper().strip()
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text

def clean_text(text):
    """Limpieza básica."""
    if text:
        return str(text).strip().upper()
    return ""

def get_float(s):
    """Convierte a float limpiando símbolos."""
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
    text = normalize_text(text)
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

# --- Lógica Forense (XML Parsing con Librería Estándar) ---

def extract_text_from_drawings(uploaded_file):
    """
    Descomprime el XLSX y busca texto dentro de los archivos XML de dibujos (shapes)
    usando xml.etree.ElementTree (nativo de Python).
    """
    drawings_text = []
    try:
        with zipfile.ZipFile(uploaded_file, 'r') as z:
            # Buscar archivos de dibujos (xl/drawings/drawing*.xml)
            drawing_files = [f for f in z.namelist() if 'xl/drawings/drawing' in f]
            
            for df in drawing_files:
                with z.open(df) as f:
                    content = f.read()
                    # Parsear XML nativamente
                    root = ET.fromstring(content)
                    # Iterar recursivamente sobre todos los elementos
                    for elem in root.iter():
                        # El texto en OpenXML suele estar en etiquetas que terminan en 't' (ej: <a:t>)
                        if elem.tag.endswith('}t'):
                            if elem.text and elem.text.strip():
                                drawings_text.append(elem.text.strip())
    except Exception as e:
        # Silencioso en caso de error para no romper la app principal
        print(f"Nota: No se pudo realizar extracción forense completa: {e}")
        return []
    
    return drawings_text

# --- Lógica de Búsqueda Estructural (Celdas) ---

def find_label_cell(sheet, keywords):
    keywords_norm = [normalize_text(k) for k in keywords]
    for row in sheet.iter_rows(min_row=1, max_row=150, max_col=50):
        for cell in row:
            val = normalize_text(cell.value)
            if not val: continue
            for k in keywords_norm:
                if k in val: return cell.row, cell.column
    return None

def extract_value_near_label(sheet, label_row, label_col, look_down=True, look_right=True, max_steps=10):
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
                val_norm = normalize_text(val)
                if val_norm:
                    if "CLIENTE" in val_norm or "FECHA" in val_norm: continue
                    return val
            except: pass
    return None

def scan_headers_footers(sheet):
    texts = []
    try:
        props = [sheet.oddHeader.left, sheet.oddHeader.center, sheet.oddHeader.right, sheet.oddFooter.left, sheet.oddFooter.center, sheet.oddFooter.right]
        for p in props:
            if p.text: texts.append(p.text)
    except: pass
    return " | ".join(texts) if texts else None

def find_longest_text_in_footer(sheet, start_row, min_length=20):
    longest_text = None
    max_len = 0
    for r in range(start_row, min(start_row + 50, sheet.max_row + 1)):
        for c in range(1, 25): 
            try:
                cell = sheet.cell(row=r, column=c)
                val = str(cell.value).strip() if cell.value else ""
                val_norm = normalize_text(val)
                if any(x in val_norm for x in ["TOTAL", "PAGE", "FIRMA", "SIGNATURE"]): continue
                if len(val) > min_length and len(val) > max_len:
                    max_len = len(val)
                    longest_text = val
            except: continue
    return longest_text

# --- Debug Helper ---
def get_all_sheet_text(sheet):
    data = []
    for r in range(1, min(sheet.max_row + 1, 200)):
        for c in range(1, min(sheet.max_column + 1, 30)):
            cell = sheet.cell(row=r, column=c)
            if cell.value:
                data.append({"Ubicación": f"Fila {r}, Col {c}", "Tipo": "Celda", "Valor": str(cell.value)})
    hf_text = scan_headers_footers(sheet)
    if hf_text: data.append({"Ubicación": "Header/Footer", "Tipo": "Impresión", "Valor": hf_text})
    return pd.DataFrame(data)

# --- Procesamiento de Archivo ---

def process_file(uploaded_file):
    # 1. Extracción Forense (Drawings/Text Boxes)
    forensic_texts = []
    try:
        # Necesitamos leer el archivo dos veces, una como ZIP y otra como Excel
        # Guardamos la posición actual
        uploaded_file.seek(0)
        forensic_texts = extract_text_from_drawings(uploaded_file)
        uploaded_file.seek(0) # Reset pointer para openpyxl
    except Exception as e:
        print(f"Error en forense: {e}")

    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active
    except Exception as e:
        return [], f"Error leyendo Excel: {str(e)}", None, forensic_texts

    header_data = {}
    debug_df = get_all_sheet_text(sheet)

    # --- EXTRACCIÓN DE DATOS ---

    # 1. CLIENTE
    lbl_coords = find_label_cell(sheet, ['CLIENTE', 'CUSTOMER', 'CONSIGNEE'])
    header_data['Cliente'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1]) if lbl_coords else None

    # 2. EXP
    lbl_coords = find_label_cell(sheet, ['EXP', 'INVOICE NO', 'FACTURA N'])
    header_data['Num_Exp'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1]) if lbl_coords else None

    # 3. FECHA
    cy = find_label_cell(sheet, ['AÑO', 'YEAR'])
    cm = find_label_cell(sheet, ['MES', 'MONTH'])
    cd = find_label_cell(sheet, ['DIA', 'DAY'])
    if cy and cm and cd:
        val_m = extract_value_near_label(sheet, cm[0], cm[1])
        val_d = extract_value_near_label(sheet, cd[0], cd[1])
        val_y = extract_value_near_label(sheet, cy[0], cy[1])
        m_num = parse_month(str(val_m))
        header_data['Fecha'] = f"{val_d}/{m_num}/{val_y}"
    else:
        c_date = find_label_cell(sheet, ['FECHA', 'DATE'])
        if c_date:
            raw_date = extract_value_near_label(sheet, c_date[0], c_date[1])
            header_data['Fecha'] = raw_date.strftime('%d/%m/%Y') if hasattr(raw_date, 'strftime') else raw_date
        else:
            header_data['Fecha'] = None

    # 4. CONDICIÓN DE VENTA (Lógica V16 con Forense)
    condicion_found = None
    
    # A) Buscar en Textos Forenses (Cuadros de Texto)
    for txt in forensic_texts:
        txt_norm = normalize_text(txt)
        if "BAJO CONDICION" in txt_norm or "UNDER CONDITION" in txt_norm or "CONSIGNACION" in txt_norm:
            condicion_found = txt 
            break
            
    # B) Si no, buscar en celdas normales
    if not condicion_found:
        for idx, row in debug_df.iterrows():
            txt = normalize_text(row['Valor'])
            if "BAJO CONDICION" in txt or "UNDER CONDITION" in txt:
                condicion_found = row['Valor']
                break

    if condicion_found:
        header_data['Condicion_Venta'] = condicion_found
    else:
        lbl_coords = find_label_cell(sheet, ['CONDICION DE VENTA', 'TERMS OF SALE', 'DELIVERY TERMS', 'INCOTERM'])
        if lbl_coords:
             header_data['Condicion_Venta'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1])
        else:
            # Fallback a Flete solo si no se encontró nada más
            lbl_flete = find_label_cell(sheet, ['TIPO FLETE', 'FREIGHT TYPE'])
            if lbl_flete:
                 header_data['Condicion_Venta'] = extract_value_near_label(sheet, lbl_flete[0], lbl_flete[1])
            else:
                 header_data['Condicion_Venta'] = None

    # 5. FORMA DE PAGO
    lbl_coords = find_label_cell(sheet, ['FORMA DE PAGO', 'PAYMENT TERMS'])
    header_data['Forma_Pago'] = extract_value_near_label(sheet, lbl_coords[0], lbl_coords[1]) if lbl_coords else None

    # 6. PUERTOS
    c_emb = find_label_cell(sheet, ['PUERTO DE EMBARQUE', 'LOADING PORT', 'POL'])
    header_data['Puerto_Emb'] = extract_value_near_label(sheet, c_emb[0], c_emb[1]) if c_emb else None
    c_dest = find_label_cell(sheet, ['PUERTO DESTINO', 'DISCHARGING PORT', 'DESTINATION', 'POD'])
    header_data['Puerto_Dest'] = extract_value_near_label(sheet, c_dest[0], c_dest[1]) if c_dest else None

    # 7. MONEDA, INCOTERM, TIPO DE CAMBIO
    detected = {'Moneda': None, 'Incoterm': None}
    incoterms = ['FOB', 'CIF', 'CFR', 'EXW', 'FCA']
    currencies = {'DÓLAR': 'USD', 'DOLAR': 'USD', 'USD': 'USD', 'EURO': 'EUR'}
    
    # Escaneo híbrido: Celdas + Forense
    all_texts_to_scan = [row['Valor'] for _, row in debug_df.iterrows()] + forensic_texts
    
    for txt in all_texts_to_scan:
        v_norm = normalize_text(txt)
        if not detected['Incoterm']:
            for inc in incoterms:
                if re.search(rf"\b{inc}\b", v_norm): detected['Incoterm'] = inc
        if not detected['Moneda']:
            for cur, code in currencies.items():
                if cur in v_norm: detected['Moneda'] = code
                
    header_data['Moneda'] = detected['Moneda']
    header_data['Incoterm'] = detected['Incoterm']
    
    lbl_tc = find_label_cell(sheet, ['TIPO DE CAMBIO', 'TIPO CAMBIO', 'T.C.', 'EXCHANGE RATE'])
    header_data['Tipo_Cambio'] = extract_value_near_label(sheet, lbl_tc[0], lbl_tc[1]) if lbl_tc else None

    # --- PRODUCTOS ---
    products = []
    col_map = {}
    header_row = None
    last_row_table = 20
    kw_desc = ['DESCRIPCION', 'DESCRIPTION', 'MERCADERIA']
    kw_qty = ['CANTIDAD', 'QUANTITY']
    kw_price = ['PRECIO', 'PRICE']
    kw_total = ['TOTAL']

    for r in range(1, 100):
        row_cells = [(c, normalize_text(sheet.cell(row=r, column=c).value)) for c in range(1, 30)]
        row_txt = [x[1] for x in row_cells]
        if any(any(k in t for k in kw_desc) for t in row_txt) and any(any(k in t for k in kw_qty) for t in row_txt):
            header_row = r
            for c, v in row_cells:
                if not v: continue
                if any(k in v for k in kw_desc): col_map['Descripcion'] = c
                elif any(k in v for k in kw_qty): col_map['Cantidad'] = c
                elif any(k in v for k in kw_price): col_map['Precio Unitario'] = c
                elif any(k in v for k in kw_total) and 'TOTAL' not in v: col_map['Total Linea'] = c
            break

    if header_row and 'Descripcion' in col_map:
        curr = header_row + 1
        empty_streak = 0
        while curr <= sheet.max_row:
            desc_val = sheet.cell(row=curr, column=col_map['Descripcion']).value
            desc_norm = normalize_text(desc_val)
            if desc_norm.startswith("TOTAL"):
                last_row_table = curr
                break
            
            qty = sheet.cell(row=curr, column=col_map['Cantidad']).value if 'Cantidad' in col_map else 0
            price = sheet.cell(row=curr, column=col_map['Precio Unitario']).value if 'Precio Unitario' in col_map else 0
            total = sheet.cell(row=curr, column=col_map['Total Linea']).value if 'Total Linea' in col_map else 0
            q = get_float(qty); p = get_float(price); t = get_float(total)
            
            if p > 0 and desc_norm and "TOTAL" not in desc_norm:
                products.append({
                    'Descripcion': desc_val, 'Cantidad': q, 'Precio Unitario': p, 
                    'Total Linea': t if t > 0 else (q*p)
                })
                empty_streak = 0; last_row_table = curr
            else:
                if not desc_norm:
                    empty_streak += 1
                    if empty_streak > 15: last_row_table = curr - 15; break
            curr += 1
    else: last_row_table = 20

    # --- OBSERVACIONES (Celdas + Forense) ---
    obs_text = None
    
    # 1. Etiquetas en celdas
    obs_lbl = find_label_cell(sheet, ['OBSERVACIONES', 'REMARKS', 'NOTAS'])
    if obs_lbl: obs_text = extract_value_near_label(sheet, obs_lbl[0], obs_lbl[1])
    
    # 2. Forense: Buscar texto largo en dibujos (Text Boxes)
    if not obs_text and forensic_texts:
        longest_drawing = max(forensic_texts, key=len) if forensic_texts else None
        if longest_drawing and len(longest_drawing) > 20:
            obs_text = longest_drawing 

    # 3. Footer Celdas
    if not obs_text: obs_text = find_longest_text_in_footer(sheet, last_row_table + 1)
    
    header_data['Observaciones'] = obs_text

    # --- Unificación ---
    final_rows = []
    if products:
        for p in products: final_rows.append({'Archivo': uploaded_file.name, **header_data, **p})
    else: final_rows.append({'Archivo': uploaded_file.name, **header_data})
        
    return final_rows, None, debug_df, forensic_texts

# --- Interfaz Gráfica ---

st.set_page_config(page_title="Extractor V17 Forense", layout="wide")
st.title("📄 Extractor V17: Modo Forense")
st.info("Versión optimizada que usa librerías estándar de Python para leer datos ocultos en cuadros de texto.")

uploaded_files = st.file_uploader("Archivos Excel (.xlsx)", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("Procesar Archivos"):
        all_data = []
        forensic_info = {}
        
        for file in uploaded_files:
            rows, err, debug_df, forensic_txt = process_file(file)
            if rows: 
                all_data.extend(rows)
                forensic_info[file.name] = forensic_txt
            if err: st.error(f"{file.name}: {err}")
        
        tab1, tab2 = st.tabs(["📊 Resultados", "🕵️ Datos Ocultos Encontrados"])
        
        with tab1:
            if all_data:
                df = pd.DataFrame(all_data)
                cols_order = ['Archivo', 'Cliente', 'Num_Exp', 'Fecha', 'Condicion_Venta', 'Incoterm', 'Forma_Pago', 
                              'Moneda', 'Tipo_Cambio', 'Puerto_Emb', 'Puerto_Dest', 'Cantidad', 'Descripcion', 
                              'Precio Unitario', 'Total Linea', 'Observaciones']
                final_cols = [c for c in cols_order if c in df.columns] + [c for c in df.columns if c not in cols_order]
                df = df[final_cols]
                st.dataframe(df, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer: df.to_excel(writer, index=False)
                st.download_button("Descargar Excel", buffer.getvalue(), "facturas_v17.xlsx")
            else: st.warning("No se extrajeron datos.")

        with tab2:
            st.info("Texto encontrado en Cuadros de Texto (Shapes):")
            for fname, texts in forensic_info.items():
                with st.expander(f"Archivo: {fname}"):
                    if texts:
                        st.write(texts)
                    else:
                        st.write("No se encontraron cuadros de texto ocultos.")
