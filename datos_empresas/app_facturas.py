import streamlit as st
import pandas as pd
import openpyxl
import io

def clean_text(text):
    """Limpia espacios y normaliza texto para comparaciones."""
    if text:
        return str(text).strip().upper()
    return ""

def find_cell_coordinates(sheet, keywords):
    """
    Busca en toda la hoja las coordenadas (fila, columna) de las palabras clave.
    Devuelve un diccionario {keyword: (row, col)}.
    """
    found_coords = {}
    # Recorremos la hoja. Limitamos la búsqueda a las primeras 100 filas/columnas por eficiencia
    # Ajusta max_row/max_col si las facturas son muy extensas.
    for row in sheet.iter_rows(min_row=1, max_row=100, max_col=50):
        for cell in row:
            cell_text = clean_text(cell.value)
            if not cell_text:
                continue
            
            for key in keywords:
                # Búsqueda flexible: si la palabra clave está contenida en la celda
                if key in cell_text:
                    # Guardamos la primera ocurrencia encontrada
                    if key not in found_coords:
                        found_coords[key] = (cell.row, cell.column)
    return found_coords

def get_value_near(sheet, row, col, search_order=['below', 'right']):
    """
    Intenta obtener un valor en celdas adyacentes (abajo o derecha).
    Retorna el primer valor no nulo encontrado.
    """
    offsets = {
        'below': (1, 0),
        'right': (0, 1),
        'below_2': (2, 0), # A veces hay una celda vacía en medio
        'right_2': (0, 2)
    }
    
    for direction in search_order:
        dr, dc = offsets.get(direction, (0,0))
        try:
            target_cell = sheet.cell(row=row + dr, column=col + dc)
            val = target_cell.value
            if val is not None and str(val).strip() != "":
                return val
        except:
            continue
    return None

def process_excel_file(uploaded_file):
    """Procesa un archivo Excel individual y extrae los datos."""
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active # Asumimos que la data está en la primera hoja
    except Exception as e:
        return None, f"Error al leer archivo: {e}"

    data_extracted = {
        "Archivo": uploaded_file.name,
        "Cliente": None,
        "Num_Exp": None,
        "Fecha": None,
        "Incoterm": None,
        "Puerto_Emb": None,
        "Puerto_Dest": None,
        "Moneda": None,
        "Observaciones": None,
        "Productos": [] # Lista de diccionarios para los productos
    }

    # 1. Mapeo de Etiquetas (Keywords) a buscar
    # Claves = Texto a buscar en el Excel, Valores = Campo interno
    keywords_header = {
        'CLIENTE / CUSTOMER': 'Cliente',
        'EXP': 'Num_Exp',
        'AÑO / YEAR': 'Anio',
        'MES / MONTH': 'Mes',
        'DIA / DAY': 'Dia',
        'CONDICION DE VENTA': 'Incoterm', # O 'TERMS OF SALE'
        'PUERTO DE EMBARQUE': 'Puerto_Emb',
        'PUERTO DESTINO': 'Puerto_Dest',
        'OBSERVACIONES': 'Observaciones',
        'MONEDA': 'Moneda',
        'DÓLAR': 'Moneda_Detectada' # Caso especial si la moneda está explícita
    }

    # Buscamos dónde están las etiquetas
    coords = find_cell_coordinates(sheet, keywords_header.keys())

    # 2. Extracción de Cabecera (Lógica difusa)
    if 'CLIENTE / CUSTOMER' in coords:
        r, c = coords['CLIENTE / CUSTOMER']
        data_extracted['Cliente'] = get_value_near(sheet, r, c, ['below'])

    if 'EXP' in coords:
        r, c = coords['EXP']
        # A veces está a la derecha, a veces abajo
        data_extracted['Num_Exp'] = get_value_near(sheet, r, c, ['right', 'below', 'right_2'])
    
    # Construcción de Fecha
    year = get_value_near(sheet, coords['AÑO / YEAR'][0], coords['AÑO / YEAR'][1], ['below']) if 'AÑO / YEAR' in coords else ""
    month = get_value_near(sheet, coords['MES / MONTH'][0], coords['MES / MONTH'][1], ['below']) if 'MES / MONTH' in coords else ""
    day = get_value_near(sheet, coords['DIA / DAY'][0], coords['DIA / DAY'][1], ['below']) if 'DIA / DAY' in coords else ""
    data_extracted['Fecha'] = f"{day}/{month}/{year}" if (year and month and day) else None

    # Logística
    if 'PUERTO DESTINO' in coords: # Búsqueda parcial permitida por find_cell_coordinates
        r, c = coords['PUERTO DESTINO']
        data_extracted['Puerto_Dest'] = get_value_near(sheet, r, c, ['below', 'right'])
    
    # Moneda (Lógica simple)
    if 'DÓLAR' in coords:
         data_extracted['Moneda'] = 'USD'
    elif 'MONEDA' in coords:
        r, c = coords['MONEDA']
        data_extracted['Moneda'] = get_value_near(sheet, r, c, ['right', 'below'])

    # 3. Extracción de Tabla de Productos
    # Estrategia: Buscar fila de cabecera de productos y barrer hacia abajo
    product_headers_map = {
        'CANTIDAD': None, # Guardaremos la columna (int)
        'DESCRIPCION': None,
        'PRECIO UNIT': None,
        'TOTAL': None
    }
    
    # Localizar columnas de la tabla
    header_row = None
    
    # Buscamos la fila que contenga 'DESCRIPCION' y 'CANTIDAD'
    # Usamos las coordenadas ya encontradas o buscamos de nuevo específicamente para headers de tabla
    table_keywords = ['CANTIDAD', 'DESCRIPCION', 'PRECIO', 'TOTAL']
    table_coords = find_cell_coordinates(sheet, table_keywords)

    # Identificar la fila principal de la tabla (usando DESCRIPCION como ancla)
    if 'DESCRIPCION' in table_coords:
        header_row = table_coords['DESCRIPCION'][0] # Fila
        
        # Mapear columnas en esa fila específica
        for col in range(1, sheet.max_column + 1):
            val = clean_text(sheet.cell(row=header_row, column=col).value)
            if 'CANTIDAD' in val or 'QUANTITY' in val:
                product_headers_map['CANTIDAD'] = col
            elif 'DESCRIPCION' in val or 'DESCRIPTION' in val:
                product_headers_map['DESCRIPCION'] = col
            elif 'PRECIO' in val or 'PRICE' in val:
                product_headers_map['PRECIO UNIT'] = col
            elif 'TOTAL' in val and 'TOTAL CASES' not in val: # Cuidado con el footer
                product_headers_map['TOTAL'] = col

        # Iterar filas debajo del header
        if product_headers_map['DESCRIPCION']:
            current_row = header_row + 1
            while current_row <= sheet.max_row:
                desc_cell = sheet.cell(row=current_row, column=product_headers_map['DESCRIPCION'])
                desc_val = desc_cell.value
                
                # Criterio de parada: Si encontramos "TOTAL" en la descripción o celda vacía repetitiva
                # Nota: A veces hay filas vacías estéticas, permitiremos 1 o 2 vacías antes de cortar, 
                # pero el criterio fuerte es encontrar la palabra "TOTAL"
                desc_text = clean_text(desc_val)
                
                if "TOTAL" in desc_text:
                    break
                
                # Si hay descripción, extraemos la fila
                if desc_val:
                    qty = sheet.cell(row=current_row, column=product_headers_map['CANTIDAD']).value if product_headers_map['CANTIDAD'] else 0
                    price = sheet.cell(row=current_row, column=product_headers_map['PRECIO UNIT']).value if product_headers_map['PRECIO UNIT'] else 0
                    total = sheet.cell(row=current_row, column=product_headers_map['TOTAL']).value if product_headers_map['TOTAL'] else 0
                    
                    data_extracted['Productos'].append({
                        'Cantidad': qty,
                        'Descripcion': desc_val,
                        'Precio Unitario': price,
                        'Total Linea': total
                    })
                
                current_row += 1

    return data_extracted, None

# --- UI Streamlit ---

st.set_page_config(page_title="Extractor de Facturas de Exportación", layout="wide")

st.title("📂 Procesador Inteligente de Facturas de Exportación")
st.markdown("""
Esta herramienta procesa archivos Excel `.xlsx` de facturas buscando etiquetas clave 
(Cliente, EXP, Cantidad, Descripción) sin depender de posiciones fijas.
""")

uploaded_files = st.file_uploader("Sube tus archivos Excel (.xlsx)", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("Procesar Archivos"):
        all_invoices = []
        all_products = []
        
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            data, error = process_excel_file(file)
            
            if error:
                st.error(f"Error en {file.name}: {error}")
                continue
            
            # Separar datos de cabecera y productos para dos vistas diferentes
            # 1. Cabecera (1 fila por factura)
            header_info = data.copy()
            del header_info['Productos'] # Removemos la lista anidada para el DF plano
            all_invoices.append(header_info)
            
            # 2. Productos (N filas por factura)
            for prod in data['Productos']:
                prod_row = prod.copy()
                prod_row['Archivo_Origen'] = data['Archivo']
                prod_row['Num_Exp'] = data['Num_Exp'] # Relacionar con la factura
                all_products.append(prod_row)
                
            progress_bar.progress((i + 1) / len(uploaded_files))
            
        st.success(f"Procesados {len(uploaded_files)} archivos exitosamente.")
        
        # --- Visualización de Resultados ---
        
        tab1, tab2 = st.tabs(["📋 Resumen de Facturas (Cabeceras)", "📦 Detalle de Productos"])
        
        with tab1:
            if all_invoices:
                df_headers = pd.DataFrame(all_invoices)
                st.dataframe(df_headers, use_container_width=True)
                
                # Descargar
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_headers.to_excel(writer, sheet_name='Facturas', index=False)
                
                st.download_button(
                    label="Descargar Resumen Facturas (Excel)",
                    data=buffer.getvalue(),
                    file_name="resumen_facturas.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.info("No se encontraron datos de cabecera.")

        with tab2:
            if all_products:
                df_products = pd.DataFrame(all_products)
                # Reordenar columnas para mejor lectura
                cols = ['Archivo_Origen', 'Num_Exp', 'Cantidad', 'Descripcion', 'Precio Unitario', 'Total Linea']
                # Filtrar solo columnas que existan (por seguridad)
                cols = [c for c in cols if c in df_products.columns]
                df_products = df_products[cols]
                
                st.dataframe(df_products, use_container_width=True)
                
                # Descargar
                buffer_prod = io.BytesIO()
                with pd.ExcelWriter(buffer_prod, engine='xlsxwriter') as writer:
                    df_products.to_excel(writer, sheet_name='Productos', index=False)
                    
                st.download_button(
                    label="Descargar Detalle Productos (Excel)",
                    data=buffer_prod.getvalue(),
                    file_name="detalle_productos.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.warning("No se pudieron extraer productos. Revisa si las columnas 'CANTIDAD' y 'DESCRIPCION' existen en el archivo.")
