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
    Busca una etiqueta y extrae el valor:
    1. En la misma celda (si hay un ':')
    2. En la celda de la derecha.
    3. En la celda de abajo.
    """
    cell = find_cell_by_keywords(sheet, keywords)
    if not cell:
        return None
    
    val_cell = str(cell.value)
    
    # 1. Intentar extraer de la misma celda tras ':'
    if ":" in val_cell:
        parts = val_cell.split(":", 1)
        if len(parts) > 1 and parts[1].strip():
            return parts[1].strip()
            
    # 2. Intentar a la derecha
    val_right = sheet.cell(row=cell.row, column=cell.column + 1).value
    if val_right is not None and str(val_right).strip() != "":
        return val_right
    
    # 3. Intentar abajo
    val_below = sheet.cell(row=cell.row + 1, column=cell.column).value
    return val_below

def extract_product_table(sheet):
    """
    Localiza la tabla buscando encabezados de columna.
    """
    # Buscamos la fila que contenga 'CANTIDAD' o 'DESCRIPCIÓN'
    header_cell = find_cell_by_keywords(sheet, ["CANTIDAD", "QUANTITY", "DESCRIP", "ITEM"])
    if not header_cell:
        return pd.DataFrame()

    headers_row = header_cell.row
    col_map = {}
    
    # Mapeo dinámico de columnas por contenido
    for col in range(1, sheet.max_column + 1):
        val = clean_text(sheet.cell(row=headers_row, column=col).value)
        if not val: continue
        
        if any(k in val for k in ["CANT", "QTY", "CANTIDAD"]): col_map['cantidad'] = col
        elif any(k in val for k in ["DESC", "ITEM", "DETALLE"]): col_map['descripcion'] = col
        elif any(k in val for k in ["UNIT", "PRECIO", "P.U"]): col_map['precio'] = col
        elif any(k in val for k in ["TOTAL", "SUBTOTAL", "IMPORTE"]): col_map['total'] = col

    # Si no encontramos columnas críticas, abortamos
    if 'descripcion' not in col_map:
        return pd.DataFrame()

    data = []
    for r in range(headers_row + 1, sheet.max_row + 1):
        desc = sheet.cell(row=r, column=col_map.get('descripcion')).value
        if not desc or "TOTAL" in str(desc).upper():
            # Si llegamos a una celda vacía o al resumen de totales, paramos
            break
            
        row_data = {
            "Cantidad": sheet.cell(row=r, column=col_map.get('cantidad', 0)).value if 'cantidad' in col_map else 0,
            "Descripción": desc,
            "Precio Unitario": sheet.cell(row=r, column=col_map.get('precio', 0)).value if 'precio' in col_map else 0,
            "Total": sheet.cell(row=r, column=col_map.get('total', 0)).value if 'total' in col_map else 0
        }
        data.append(row_data)
        
    return pd.DataFrame(data)

def run_app():
    st.title("🚀 Extractor de Facturas de Exportación (Versión Mejorada)")
    
    uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            sheet = wb.active
            
            # --- Extracción con Sinónimos ---
            data = {
                "Cliente": extract_value_smart(sheet, ["CLIENTE", "CUSTOMER", "SOLD TO"]),
                "Expediente": extract_value_smart(sheet, ["EXP", "NRO EXP", "EXPORT NR"]),
                "Incoterm": extract_value_smart(sheet, ["INCOTERM", "CONDICION VENTA", "TERMS"]),
                "Puerto Emb.": extract_value_smart(sheet, ["EMBARQUE", "LOADING", "POL"]),
                "Puerto Dest.": extract_value_smart(sheet, ["DESTINO", "DESTINATION", "POD"]),
                "Moneda": extract_value_smart(sheet, ["MONEDA", "CURRENCY", "DIVISA"]),
                "Fecha": extract_value_smart(sheet, ["FECHA", "DATE"])
            }
            
            # Mostrar Resumen de Cabecera
            st.subheader("Datos de Cabecera Detectados")
            cols = st.columns(4)
            for i, (k, v) in enumerate(data.items()):
                cols[i % 4].metric(k, str(v) if v else "No encontrado")
            
            # Extraer Tabla
            df_items = extract_product_table(sheet)
            
            if not df_items.empty:
                st.subheader("Detalle de Productos")
                st.dataframe(df_items, use_container_width=True)
                
                # Consolidar para descarga
                for k, v in data.items():
                    df_items[k] = v
                
                # Descargas
                st.divider()
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_items.to_excel(writer, index=False)
                
                st.download_button(
                    "📥 Descargar Resultado consolidado (Excel)",
                    data=output.getvalue(),
                    file_name="factura_procesada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No se detectó la tabla de productos. Asegúrate de que las columnas tengan nombres estándar como 'Cantidad', 'Descripción', etc.")
                
        except Exception as e:
            st.error(f"Error crítico: {e}")

if __name__ == "__main__":
    run_app()
