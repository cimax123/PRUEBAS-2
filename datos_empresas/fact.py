import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import datetime
import re

# Configuración de la página
st.set_page_config(page_title="Extractor de Facturas Pro", layout="wide")

def clean_text(text):
    if not text or not isinstance(text, str):
        return ""
    return text.strip().upper()

def find_cell_by_keywords(sheet, keywords):
    """
    Busca una celda que contenga cualquiera de las palabras clave.
    """
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.upper()
                if any(k.upper() in val for k in keywords):
                    return cell
    return None

def extract_value_smart(sheet, keywords):
    """
    Busca una etiqueta y extrae el valor con lógica mejorada:
    1. Misma celda: Limpia la etiqueta del valor (ej. "RUT 123" -> "123").
    2. Celda derecha.
    3. Celda inferior.
    """
    cell = find_cell_by_keywords(sheet, keywords)
    if not cell:
        return None
    
    val_cell = str(cell.value).strip()
    val_upper = val_cell.upper()
    
    # 1. Intentar extraer de la misma celda (Limpieza agresiva)
    # Si hay dos puntos, separamos
    if ":" in val_cell:
        parts = val_cell.split(":", 1)
        if len(parts) > 1 and parts[1].strip():
            return parts[1].strip()
    
    # Si no hay dos puntos, pero hay texto además de la keyword
    # Ej: "RUT 123456" -> Quitamos "RUT" y devolvemos "123456"
    for k in keywords:
        if k.upper() in val_upper:
            # Creamos una regex para reemplazar la keyword (case insensitive)
            cleaned = re.sub(f"(?i){re.escape(k)}", "", val_cell).strip()
            # Quitamos caracteres comunes de separación que puedan haber quedado
            cleaned = cleaned.lstrip(".:- ").strip()
            if len(cleaned) > 1: # Asumimos que un valor real tiene al menos 2 chars
                return cleaned

    # 2. Intentar a la derecha
    val_right = sheet.cell(row=cell.row, column=cell.column + 1).value
    if val_right is not None and str(val_right).strip() != "":
        return val_right
    
    # 3. Intentar abajo
    val_below = sheet.cell(row=cell.row + 1, column=cell.column).value
    return val_below

def find_incoterm_advanced(sheet):
    """
    Busca Incoterms estándar (FOB, CIF, EXW, etc.) si la búsqueda normal falla.
    """
    # 1. Búsqueda normal por etiqueta
    val = extract_value_smart(sheet, ["INCOTERM", "CONDICION VENTA", "TERMS", "DELIVERY"])
    if val: return val

    # 2. Búsqueda por códigos de Incoterms 2020/2010
    incoterms_list = [
        "EXW", "FCA", "CPT", "CIP", "DAP", "DPU", "DAT", "DDP", 
        "FAS", "FOB", "CFR", "CIF"
    ]
    
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val_upper = cell.value.upper()
                # Buscamos la palabra exacta (ej. "FOB" pero no "FOOBAR")
                for code in incoterms_list:
                    # Regex busca la palabra completa \bCODE\b
                    if re.search(r'\b' + code + r'\b', val_upper):
                        return cell.value # Devolvemos el texto completo (ej: "FOB VALPARAISO")
    return None

def extract_product_table(sheet):
    """
    Localiza la tabla buscando encabezados de columna.
    """
    # Buscamos la fila que contenga 'CANTIDAD' o 'DESCRIPCIÓN'
    header_cell = find_cell_by_keywords(sheet, ["CANTIDAD", "QUANTITY", "DESCRIP", "ITEM", "PRODUCTO"])
    if not header_cell:
        return pd.DataFrame()

    headers_row = header_cell.row
    col_map = {}
    
    # Mapeo dinámico de columnas por contenido
    for col in range(1, sheet.max_column + 1):
        val = clean_text(sheet.cell(row=headers_row, column=col).value)
        if not val: continue
        
        if any(k in val for k in ["CANT", "QTY", "CANTIDAD", "UNIDADES"]): col_map['cantidad'] = col
        elif any(k in val for k in ["DESC", "ITEM", "DETALLE", "PRODUCTO", "GOODS"]): col_map['descripcion'] = col
        elif any(k in val for k in ["UNIT", "PRECIO", "P.U", "PRICE"]): col_map['precio'] = col
        elif any(k in val for k in ["TOTAL", "SUBTOTAL", "IMPORTE", "AMOUNT"]): col_map['total'] = col

    # Si no encontramos columnas críticas, abortamos
    if 'descripcion' not in col_map:
        return pd.DataFrame()

    data = []
    # Escanear filas hacia abajo
    for r in range(headers_row + 1, sheet.max_row + 1):
        desc = sheet.cell(row=r, column=col_map.get('descripcion')).value
        
        # Criterios de parada
        if not desc:
            continue # Saltamos filas vacías intermedias, pero no paramos
            
        desc_str = str(desc).upper()
        if "TOTAL" in desc_str or "SON:" in desc_str or "PAGE" in desc_str:
            break
            
        row_data = {
            "Cantidad": sheet.cell(row=r, column=col_map.get('cantidad', 0)).value if 'cantidad' in col_map else 0,
            "Descripción": desc,
            "Precio Unitario": sheet.cell(row=r, column=col_map.get('precio', 0)).value if 'precio' in col_map else 0,
            "Total": sheet.cell(row=r, column=col_map.get('total', 0)).value if 'total' in col_map else 0
        }
        
        # Verificar si la fila tiene datos reales (no solo formato)
        if row_data["Descripción"] or row_data["Total"]:
            data.append(row_data)
        
    return pd.DataFrame(data)

def run_app():
    st.title("🚀 Extractor de Facturas de Exportación (Versión Mejorada)")
    
    uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            sheet = wb.active
            
            # --- Extracción ---
            # Usamos lógica avanzada para Incoterms y RUT
            data = {
                "Cliente": extract_value_smart(sheet, ["CLIENTE", "CUSTOMER", "SOLD TO", "BUYER"]),
                "RUT": extract_value_smart(sheet, ["RUT", "R.U.T", "TAX ID", "VAT"]),
                "Expediente": extract_value_smart(sheet, ["EXP", "NRO EXP", "EXPORT NR", "REF"]),
                "Incoterm": find_incoterm_advanced(sheet), # Nueva función específica
                "Puerto Emb.": extract_value_smart(sheet, ["EMBARQUE", "LOADING", "POL", "FROM"]),
                "Puerto Dest.": extract_value_smart(sheet, ["DESTINO", "DESTINATION", "POD", "TO"]),
                "Moneda": extract_value_smart(sheet, ["MONEDA", "CURRENCY", "DIVISA"]),
                "Fecha": extract_value_smart(sheet, ["FECHA", "DATE"])
            }
            
            # Mostrar Resumen de Cabecera
            st.subheader("Datos de Cabecera Detectados")
            
            # Crear métricas en filas de 4
            keys = list(data.keys())
            for i in range(0, len(keys), 4):
                cols = st.columns(4)
                for j, col in enumerate(cols):
                    if i + j < len(keys):
                        key = keys[i + j]
                        val = data[key]
                        col.metric(key, str(val) if val else "No encontrado")
            
            # Extraer Tabla
            df_items = extract_product_table(sheet)
            
            if not df_items.empty:
                st.subheader("Detalle de Productos")
                st.dataframe(df_items, use_container_width=True)
                
                # Consolidar para descarga
                df_export = df_items.copy()
                for k, v in data.items():
                    df_export[k] = v
                
                # Descargas
                st.divider()
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False)
                
                st.download_button(
                    "📥 Descargar Resultado consolidado (Excel)",
                    data=output.getvalue(),
                    file_name="factura_procesada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se detectó la tabla de productos automáticamente. Verifica que existan encabezados como 'Cantidad' y 'Descripción'.")
                
        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")

if __name__ == "__main__":
    run_app()
