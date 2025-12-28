import streamlit as st
import pandas as pd
import openpyxl
import io

# --- Funciones de Lógica de Negocio ---

def clean_text(text):
    """Normaliza texto para búsquedas (mayúsculas y sin espacios extra)."""
    if text:
        return str(text).strip().upper()
    return ""

def find_keywords_in_sheet(sheet, keywords_list):
    """
    Escanea la hoja buscando coordenadas de palabras clave.
    Retorna un diccionario {keyword_encontrada: (fila, columna)}.
    """
    found_coords = {}
    # Escaneamos un área razonable (primeras 100 filas) para encontrar cabeceras
    for row in sheet.iter_rows(min_row=1, max_row=100, max_col=50):
        for cell in row:
            val = clean_text(cell.value)
            if not val:
                continue
            
            for key in keywords_list:
                # Si la palabra clave es parte del texto de la celda
                if key in val:
                    # Guardamos la coordenada si no la hemos encontrado antes
                    if key not in found_coords:
                        found_coords[key] = (cell.row, cell.column)
    return found_coords

def get_data_near_label(sheet, start_row, start_col, search_directions=['below', 'right'], max_steps=5):
    """
    Busca un valor no vacío cerca de una coordenada.
    search_directions: lista de direcciones a probar ('below', 'right').
    max_steps: cuántas celdas avanzar ignorando vacíos.
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
                if val is not None and str(val).strip() != "":
                    # Encontró dato no vacío
                    return val
            except:
                pass # Fuera de rango
    return None

def extract_products(sheet):
    """
    Localiza la tabla de productos dinámicamente y extrae filas.
    """
    products = []
    
    # 1. Buscar cabeceras de columnas de tabla
    headers_map = {}
    header_row = None
    
    # Palabras clave para identificar columnas
    col_keywords = {
        'CANTIDAD': ['CANTIDAD', 'QUANTITY', 'QTTY', 'PCS', 'BOXES'],
        'DESCRIPCION': ['DESCRIPCION', 'DESCRIPTION', 'MERCADERIA', 'COMMODITY', 'GOODS'],
        'PRECIO': ['PRECIO', 'PRICE', 'UNIT PRICE', 'UNITARIO'],
        'TOTAL': ['TOTAL'] # Cuidado, "TOTAL" también aparece al final
    }

    # Escanear para encontrar la fila de cabeceras de tabla
    # Buscamos la fila que tenga AL MENOS 'DESCRIPCION' y 'CANTIDAD'
    for row in sheet.iter_rows(min_row=1, max_row=100):
        row_values = [clean_text(c.value) for c in row]
        
        # Chequeo difuso: ¿Están las keywords en esta fila?
        has_desc = any(any(k in val for k in col_keywords['DESCRIPCION']) for val in row_values if val)
        has_qty = any(any(k in val for k in col_keywords['CANTIDAD']) for val in row_values if val)
        
        if has_desc and has_qty:
            header_row = row[0].row
            # Mapear columnas exactas
            for cell in row:
                val = clean_text(cell.value)
                if not val: continue
                
                # Asignar índice de columna
                if any(k in val for k in col_keywords['CANTIDAD']):
                    headers_map['Cantidad'] = cell.column
                elif any(k in val for k in col_keywords['DESCRIPCION']):
                    headers_map['Descripcion'] = cell.column
                elif any(k in val for k in col_keywords['PRECIO']):
                    headers_map['Precio Unitario'] = cell.column
                elif any(k in val for k in col_keywords['TOTAL']) and 'TOTAL CASES' not in val:
                    headers_map['Total Linea'] = cell.column
            break
    
    if not header_row or 'Descripcion' not in headers_map:
        return []

    # 2. Iterar filas debajo de la cabecera
    current_row = header_row + 1
    empty_streak = 0 # Para detener si hay muchas líneas vacías seguidas
    
    while current_row <= sheet.max_row:
        # Obtener descripción para decidir si seguimos
        desc_col = headers_map['Descripcion']
        desc_val = sheet.cell(row=current_row, column=desc_col).value
        desc_clean = clean_text(desc_val)
        
        # Criterios de parada
        if "TOTAL" in desc_clean and "CAJAS" not in desc_clean: 
            # Detecta fila de totales finales (ej: "TOTAL FOB")
            break
            
        if not desc_clean:
            empty_streak += 1
            if empty_streak > 5: # Si hay 5 filas vacías seguidas, asumimos fin de tabla
                break
        else:
            empty_streak = 0
            # Extraer datos de la fila
            row_data = {}
            row_data['Descripcion'] = desc_val
            
            if 'Cantidad' in headers_map:
                row_data['Cantidad'] = sheet.cell(row=current_row, column=headers_map['Cantidad']).value
            if 'Precio Unitario' in headers_map:
                row_data['Precio Unitario'] = sheet.cell(row=current_row, column=headers_map['Precio Unitario']).value
            if 'Total Linea' in headers_map:
                row_data['Total Linea'] = sheet.cell(row=current_row, column=headers_map['Total Linea']).value
            
            products.append(row_data)
        
        current_row += 1
        
    return products

def process_file(uploaded_file):
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active
    except Exception as e:
        return [], f"Error leyendo Excel: {str(e)}"

    # --- 1. Definición de Etiquetas a Buscar (Cabecera) ---
    # Diccionario: Clave Interna -> Lista de posibles textos en el Excel
    labels_config = {
        'Cliente': ['CLIENTE', 'CUSTOMER', 'CONSIGNEE', 'SOLD TO'],
        'Num_Exp': ['EXP', 'INVOICE NO', 'FACTURA N'],
        'Fecha_Anio': ['AÑO', 'YEAR'],
        'Fecha_Mes': ['MES', 'MONTH'],
        'Fecha_Dia': ['DIA', 'DAY'],
        'Incoterm': ['CONDICION DE VENTA', 'TERMS OF SALE', 'DELIVERY TERMS', 'INCOTERM', 'FLETE', 'FREIGHT TYPE'],
        'Puerto_Emb': ['PUERTO DE EMBARQUE', 'LOADING PORT', 'PORT OF LOADING'],
        'Puerto_Dest': ['PUERTO DESTINO', 'DISCHARGING PORT', 'DESTINATION', 'PORT OF DESTINATION', 'PUERTO DE DESTINO'],
        'Moneda': ['MONEDA', 'CURRENCY'],
        'Observaciones': ['OBSERVACIONES', 'REMARKS', 'NOTAS']
    }
    
    # Aplanamos la lista de keywords para buscar todo de una vez
    all_keywords = []
    for k, v in labels_config.items():
        all_keywords.extend(v)
        
    found_coords = find_keywords_in_sheet(sheet, all_keywords)
    
    header_data = {}
    
    # --- 2. Extracción de Valores de Cabecera ---
    for field, keywords in labels_config.items():
        # Buscar cuál de las variantes de keywords se encontró
        found_key = next((k for k in keywords if k in found_coords), None)
        
        if found_key:
            row, col = found_coords[found_key]
            # Buscamos: 1o Abajo (hasta 5 celdas), 2o Derecha (hasta 5 celdas)
            val = get_data_near_label(sheet, row, col, search_directions=['below', 'right'], max_steps=4)
            header_data[field] = val
        else:
            header_data[field] = None

    # Lógica especial para Fecha
    fecha_full = None
    if header_data['Fecha_Anio'] and header_data['Fecha_Mes'] and header_data['Fecha_Dia']:
        fecha_full = f"{header_data['Fecha_Dia']}/{header_data['Fecha_Mes']}/{header_data['Fecha_Anio']}"
    header_data['Fecha'] = fecha_full

    # Lógica especial para Incoterm si no se encuentra etiqueta
    # A veces el incoterm está pegado a un total, ej "TOTAL FOB"
    if not header_data['Incoterm']:
        # Busqueda de emergencia en toda la hoja
        # Buscamos celdas que contengan "FOB", "CIF", "EXW"
        common_incoterms = ['FOB', 'CIF', 'CFR', 'EXW', 'FCA']
        for row in sheet.iter_rows(min_row=10, max_row=sheet.max_row): # Generalmente abajo
            for cell in row:
                val = clean_text(cell.value)
                if val:
                    for inc in common_incoterms:
                        # Si encuentra "TOTAL FOB" extrae "FOB"
                        if f"TOTAL {inc}" in val or val == inc:
                            header_data['Incoterm'] = inc
                            break
            if header_data['Incoterm']: break

    # --- 3. Extracción de Productos y Unión ---
    products_list = extract_products(sheet)
    
    final_rows = []
    
    # Si hay productos, creamos una fila por producto con la cabecera repetida
    if products_list:
        for prod in products_list:
            row = {
                'Archivo': uploaded_file.name,
                **header_data, # Expande datos de cabecera
                **prod         # Expande datos de producto
            }
            # Limpiamos columnas auxiliares de fecha
            del row['Fecha_Anio']
            del row['Fecha_Mes']
            del row['Fecha_Dia']
            final_rows.append(row)
    else:
        # Si no hay productos, devolvemos al menos la cabecera
        row = {
            'Archivo': uploaded_file.name,
            **header_data
        }
        del row['Fecha_Anio']
        del row['Fecha_Mes']
        del row['Fecha_Dia']
        final_rows.append(row)
        
    return final_rows, None

# --- UI Streamlit ---

st.set_page_config(page_title="Extractor Horizontal de Facturas", layout="wide")

st.title("📄 Extractor de Facturas a Tabla Plana")
st.markdown("""
Sube tus facturas Excel. El sistema generará una **única tabla maestra** donde cada fila es un producto 
acompañado de sus datos de exportación (Cliente, Incoterm, Puerto, etc.).
""")

uploaded_files = st.file_uploader("Arrastra archivos .xlsx aquí", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("Procesar y Unificar"):
        all_data = []
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            rows, error = process_file(file)
            if rows:
                all_data.extend(rows)
            if error:
                st.toast(f"Alerta en {file.name}: {error}", icon="⚠️")
            
            progress_bar.progress((i + 1) / len(uploaded_files))
            
        if all_data:
            df = pd.DataFrame(all_data)
            
            # Ordenar columnas lógicamente
            desired_order = [
                'Archivo', 'Cliente', 'Num_Exp', 'Fecha', 'Incoterm', 
                'Puerto_Emb', 'Puerto_Dest', 'Moneda', 
                'Cantidad', 'Descripcion', 'Precio Unitario', 'Total Linea', 'Observaciones'
            ]
            # Filtrar solo las que existen
            cols = [c for c in desired_order if c in df.columns]
            # Agregar el resto que no esté en la lista deseada al final
            remaining = [c for c in df.columns if c not in cols]
            df = df[cols + remaining]

            st.success("¡Procesamiento completo!")
            st.dataframe(df, use_container_width=True)
            
            # Descarga
            buffer = io.BytesIO()
            # Usamos openpyxl explícitamente para evitar el error anterior
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Master Data")
                
            st.download_button(
                label="📥 Descargar Tabla Unificada (Excel)",
                data=buffer.getvalue(),
                file_name="facturas_procesadas_master.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se pudieron extraer datos. Verifica el formato de los archivos.")
