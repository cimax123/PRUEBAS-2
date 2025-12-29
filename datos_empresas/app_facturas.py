import pandas as pd
import numpy as np
import re
from datetime import datetime
import io
import streamlit as st

class InvoiceParser:
    def __init__(self, df):
        # Convertimos el dataframe a una matriz de cadenas para facilitar la búsqueda
        # fillna('') asegura que no fallen las comparaciones de texto
        self.df = df.astype(str).replace('nan', '')
        self.raw_data = self.df.values # CORREGIDO: Usamos self.df (texto) en lugar de df (original mixto)

    def _find_coordinates(self, keywords):
        """
        Busca las coordenadas (fila, columna) de la primera celda 
        que contenga alguna de las palabras clave.
        """
        if isinstance(keywords, str):
            keywords = [keywords]
        
        keywords = [k.upper() for k in keywords]
        
        for r_idx, row in enumerate(self.raw_data):
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).upper().strip()
                # Coincidencia exacta o parcial segura
                if any(k == cell_str or f" {k} " in f" {cell_str} " for k in keywords):
                    return r_idx, c_idx
        return None, None

    def _get_value_relative(self, anchor_keywords, row_offset=1, col_offset=0):
        """Busca un ancla y devuelve el valor en la posición relativa indicada."""
        r, c = self._find_coordinates(anchor_keywords)
        if r is not None:
            try:
                # Verificamos límites
                target_r, target_c = r + row_offset, c + col_offset
                if target_r < len(self.raw_data) and target_c < len(self.raw_data[0]):
                    val = self.raw_data[target_r][target_c]
                    # Aseguramos conversión a string por seguridad extra
                    return str(val).strip() if val else "N/A"
            except IndexError:
                pass
        return "N/A"

    def extract_date(self):
        """Lógica unificada para FECHA (Día, Mes, Año)."""
        day = self._get_value_relative(["DIA", "DIA / DAY"])
        month = self._get_value_relative(["MES", "MES / MONTH"])
        year = self._get_value_relative(["AÑO", "AÑO / YEAR", "YEAR"])
        
        if "N/A" in [day, month, year]:
            return "N/A"
        
        # Mapeo simple de meses numéricos si vienen como texto
        month_map = {
            '1': 'Enero', '01': 'Enero', 'JAN': 'Enero',
            '2': 'Febrero', '02': 'Febrero', 'FEB': 'Febrero',
            '3': 'Marzo', '03': 'Marzo', 'MAR': 'Marzo',
            '4': 'Abril', '04': 'Abril', 'APR': 'Abril',
            '5': 'Mayo', '05': 'Mayo', 'MAY': 'Mayo',
            '6': 'Junio', '06': 'Junio', 'JUN': 'Junio',
            '7': 'Julio', '07': 'Julio', 'JUL': 'Julio',
            '8': 'Agosto', '08': 'Agosto', 'AUG': 'Agosto',
            '9': 'Septiembre', '09': 'Septiembre', 'SEP': 'Septiembre',
            '10': 'Octubre', 'OCT': 'Octubre',
            '11': 'Noviembre', 'NOV': 'Noviembre',
            '12': 'Diciembre', 'DEC': 'Diciembre'
        }
        
        month_norm = month_map.get(month.upper(), month)
        return f"{day} de {month_norm} de {year}"

    def extract_sales_condition(self):
        """Divide la condición de venta en Tipo e Incoterm."""
        raw_cond = self._get_value_relative(["CONDICION VENTA", "CONDICION DE VENTA"])
        
        if raw_cond == "N/A":
            return "N/A", "N/A"
            
        # Separar por guiones, guiones largos o espacios grandes
        parts = re.split(r'\s*[-–]\s*', raw_cond)
        
        tipo_venta = parts[0].strip() if len(parts) > 0 else "N/A"
        incoterm = parts[1].strip() if len(parts) > 1 else "N/A"
        
        return tipo_venta, incoterm

    def extract_products_table(self):
        """
        Identifica la tabla de productos buscando 'CANTIDAD' y 'DESCRIPCION'.
        Extrae filas hasta encontrar un vacío o totales.
        """
        r_qty, c_qty = self._find_coordinates(["CANTIDAD", "QTY"])
        r_desc, c_desc = self._find_coordinates(["DESCRIPCION", "DESCRIPTION"])
        r_price, c_price = self._find_coordinates(["PRECIO UNIT", "UNIT PRICE"])
        r_total, c_total = self._find_coordinates(["TOTAL", "TOTAL LINEA"])

        if r_desc is None:
            return []

        products = []
        # Asumimos que los datos empiezan en la fila siguiente al encabezado encontrado
        current_r = max(r for r in [r_qty, r_desc] if r is not None) + 1
        
        while current_r < len(self.raw_data):
            # Obtener descripción como pivote para saber si la fila es válida
            desc_val = self.raw_data[current_r][c_desc] if c_desc is not None else ""
            
            # Condición de parada: Fila vacía o palabra clave de cierre (como 'TOTAL FOB')
            if not desc_val.strip() or "TOTAL" in desc_val.upper() or "OBSERVACIONES" in desc_val.upper():
                break
                
            qty_val = self.raw_data[current_r][c_qty] if c_qty is not None else "0"
            price_val = self.raw_data[current_r][c_price] if c_price is not None else "0"
            total_val = self.raw_data[current_r][c_total] if c_total is not None else "0"
            
            products.append({
                "CANTIDAD": qty_val,
                "DESCRIPCION": desc_val,
                "PRECIO UNITARIO": price_val,
                "TOTAL LINEA": total_val
            })
            current_r += 1
            
        return products

    def process(self):
        """Ejecuta todo el flujo y devuelve una lista de diccionarios (flat table)."""
        
        # 1. Extracción de Cabecera
        cliente = self._get_value_relative(["CLIENTE", "CLIENTE / CUSTOMER", "CUSTOMER"])
        exp = self._get_value_relative(["EXP", "EXP N°", "REF EXP"])
        fecha_unificada = self.extract_date()
        
        tipo_venta, incoterm = self.extract_sales_condition()
        
        puerto_emb = self._get_value_relative(["PUERTO EMBARQUE", "PORT OF LOADING"])
        puerto_dest = self._get_value_relative(["PUERTO DESTINO", "PORT OF DESTINATION"])
        
        moneda = self._get_value_relative(["MONEDA", "CURRENCY"])
        if moneda == "N/A": 
            # Intento secundario: buscar al lado de TOTAL FOB
            moneda = self._get_value_relative(["TOTAL FOB"], col_offset=1)

        # 2. Extracción de Productos (Líneas)
        products = self.extract_products_table()
        
        # 3. Construcción de filas planas (Flat Table)
        output_rows = []
        
        # Datos comunes para todas las filas
        header_data = {
            "CLIENTE": cliente,
            "EXP": exp,
            "FECHA": fecha_unificada,
            "TIPO DE VENTA": tipo_venta,
            "INCOTERM": incoterm,
            "PUERTO EMBARQUE": puerto_emb,
            "PUERTO DESTINO": puerto_dest,
            "MONEDA": moneda
        }
        
        if not products:
            # Si no hay productos, devolvemos al menos la cabecera
            row = header_data.copy()
            row.update({"CANTIDAD": "", "DESCRIPCION": "", "PRECIO UNITARIO": "", "TOTAL LINEA": ""})
            output_rows.append(row)
        else:
            for prod in products:
                row = header_data.copy()
                row.update(prod)
                output_rows.append(row)
                
        return pd.DataFrame(output_rows)

# ==========================================
# INTERFAZ DE USUARIO STREAMLIT
# ==========================================
def main():
    st.set_page_config(page_title="Extractor de Facturas", page_icon="📄", layout="wide")
    
    st.title("🤖 Extractor Inteligente de Facturas de Exportación")
    st.markdown("""
    Esta herramienta extrae datos de facturas Excel basándose en **etiquetas de texto (anchors)**, 
    ignorando colores o posiciones fijas. Ideal para formatos desordenados.
    
    **Datos que extrae:** Cliente, Fechas, Incoterms, Puertos y Tablas de Productos.
    """)

    # Carga de archivos múltiple
    uploaded_files = st.file_uploader(
        "Arrastra y suelta tus facturas Excel aquí (uno o varios)", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True
    )

    if uploaded_files:
        all_data = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, file in enumerate(uploaded_files):
            try:
                status_text.text(f"Procesando archivo: {file.name}...")
                
                # Leemos el archivo (sin header para usar coordenadas puras)
                df_raw = pd.read_excel(file, header=None)
                
                # Instanciamos la lógica de extracción
                parser = InvoiceParser(df_raw)
                df_result = parser.process()
                
                # Añadimos el nombre del archivo para trazabilidad
                df_result.insert(0, "ARCHIVO_ORIGEN", file.name)
                
                all_data.append(df_result)
                
            except Exception as e:
                st.error(f"❌ Error al procesar {file.name}: {str(e)}")
            
            # Actualizar barra de progreso
            progress_bar.progress((i + 1) / len(uploaded_files))

        status_text.text("¡Procesamiento completado!")
        
        if all_data:
            # Consolidar todo en un solo DataFrame
            final_df = pd.concat(all_data, ignore_index=True)
            
            st.success(f"✅ Se han extraído {len(final_df)} líneas de productos de {len(uploaded_files)} facturas.")
            
            # Mostrar vista previa
            st.subheader("Vista Previa de Resultados")
            st.dataframe(final_df, use_container_width=True)
            
            # Botón de descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Consolidado')
            
            st.download_button(
                label="📥 Descargar Excel Consolidado",
                data=output.getvalue(),
                file_name="reporte_facturas_exportacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("👆 Sube tus archivos Excel para comenzar.")
        
        # Opción para descargar plantilla de prueba (opcional)
        st.divider()
        st.caption("¿No tienes un archivo a mano? Genera uno de prueba:")
        if st.button("Generar Factura de Prueba (Mock)"):
            # Lógica simple para generar el mock al vuelo solo si se pide
            mock_data = [
                ["", "FACTURA DE EXPORTACIÓN", "", "", "", "", ""],
                ["", "CLIENTE / CUSTOMER", "", "EXP", "", "FECHA", ""],
                ["", "EJEMPLO CORP S.A.", "", "REF-999", "", "01/JAN/2025", ""],
                ["", "CONDICION VENTA", "", "PUERTO DESTINO", "", "MONEDA", ""],
                ["", "A FIRME - CIF - ROTTERDAM", "", "ROTTERDAM", "", "USD", ""],
                ["", "DESCRIPCION", "CANTIDAD", "PRECIO UNIT", "TOTAL", "", ""],
                ["", "CAJAS DE ARÁNDANOS", "1000", "15.00", "15000.00", "", ""],
                ["", "CAJAS DE CEREZAS", "500", "20.00", "10000.00", "", ""]
            ]
            df_mock = pd.DataFrame(mock_data)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_mock.to_excel(writer, index=False, header=False)
            
            st.download_button("Descargar Mock.xlsx", buffer.getvalue(), "mock_factura.xlsx")

if __name__ == "__main__":
    main()
