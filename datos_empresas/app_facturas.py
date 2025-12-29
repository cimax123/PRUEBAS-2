import pandas as pd
import numpy as np
import re
from datetime import datetime
import io
import streamlit as st

class InvoiceParser:
    def __init__(self, df):
        # Convertimos el dataframe a una matriz de cadenas para facilitar la búsqueda
        # Aseguramos que todo sea texto para evitar errores con floats
        self.df = df.astype(str).replace('nan', '')
        self.raw_data = self.df.values 

    def _find_coordinates(self, keywords):
        """Busca las coordenadas (fila, columna) de una palabra clave."""
        if isinstance(keywords, str):
            keywords = [keywords]
        keywords = [k.upper() for k in keywords]
        
        for r_idx, row in enumerate(self.raw_data):
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).upper().strip()
                if any(k == cell_str or f" {k} " in f" {cell_str} " for k in keywords):
                    return r_idx, c_idx
        return None, None

    def _scan_neighborhood(self, r, c, direction='down', max_steps=5):
        """
        Escanea celdas vecinas (abajo o derecha) saltando vacíos hasta encontrar un valor.
        """
        if r is None or c is None:
            return "N/A"
            
        for i in range(1, max_steps + 1):
            try:
                if direction == 'down':
                    target_r, target_c = r + i, c
                elif direction == 'right':
                    target_r, target_c = r, c + i
                else:
                    return "N/A"

                # Verificar límites
                if target_r >= len(self.raw_data) or target_c >= len(self.raw_data[0]):
                    continue

                val = str(self.raw_data[target_r][target_c]).strip()
                # Si encontramos algo que no sea vacío, lo devolvemos
                if val: 
                    return val
            except IndexError:
                pass
        return "N/A"

    def extract_date(self):
        """Extrae y formatea la fecha a formato DD/MM/AAAA."""
        # Buscamos valores saltando posibles celdas vacías debajo de los headers
        r_d, c_d = self._find_coordinates(["DIA", "DIA / DAY"])
        day = self._scan_neighborhood(r_d, c_d, direction='down')
        
        r_m, c_m = self._find_coordinates(["MES", "MES / MONTH"])
        month = self._scan_neighborhood(r_m, c_m, direction='down')
        
        r_y, c_y = self._find_coordinates(["AÑO", "AÑO / YEAR", "YEAR"])
        year = self._scan_neighborhood(r_y, c_y, direction='down')
        
        if "N/A" in [day, month, year]:
            # Intento alternativo: Buscar "FECHA" y tomar el valor completo
            r_f, c_f = self._find_coordinates(["FECHA", "DATE", "FECHA DOCUMENTO"])
            full_date = self._scan_neighborhood(r_f, c_f, direction='down')
            return full_date if full_date != "N/A" else "N/A"
        
        # Mapeo de meses a números para formato estándar
        month_map = {
            'JANUARY': '01', 'JAN': '01', 'ENERO': '01', 'ENE': '01', '1': '01', '01': '01',
            'FEBRUARY': '02', 'FEB': '02', 'FEBRERO': '02', '2': '02', '02': '02',
            'MARCH': '03', 'MAR': '03', 'MARZO': '03', '3': '03', '03': '03',
            'APRIL': '04', 'APR': '04', 'ABRIL': '04', 'ABR': '04', '4': '04', '04': '04',
            'MAY': '05', 'MAYO': '05', '5': '05', '05': '05',
            'JUNE': '06', 'JUN': '06', 'JUNIO': '06', '6': '06', '06': '06',
            'JULY': '07', 'JUL': '07', 'JULIO': '07', '7': '07', '07': '07',
            'AUGUST': '08', 'AUG': '08', 'AGOSTO': '08', 'AGO': '08', '8': '08', '08': '08',
            'SEPTEMBER': '09', 'SEP': '09', 'SEPTIEMBRE': '09', 'SEPT': '09', '9': '09', '09': '09',
            'OCTOBER': '10', 'OCT': '10', 'OCTUBRE': '10', '10': '10',
            'NOVEMBER': '11', 'NOV': '11', 'NOVIEMBRE': '11', '11': '11',
            'DECEMBER': '12', 'DEC': '12', 'DICIEMBRE': '12', 'DIC': '12', '12': '12'
        }
        
        m_num = month_map.get(month.upper(), month)
        # Asegurar ceros a la izquierda para día
        d_num = day.zfill(2) if day.isdigit() else day
        
        return f"{d_num}/{m_num}/{year}"

    def extract_currency(self):
        """Busca la moneda cerca de 'TOTAL FOB' o etiquetas similares, escaneando a la derecha."""
        # Estrategia 1: Buscar etiqueta "MONEDA"
        r, c = self._find_coordinates(["MONEDA", "CURRENCY"])
        if r is not None:
            val = self._scan_neighborhood(r, c, direction='down') # A veces está abajo
            if val == "N/A": 
                val = self._scan_neighborhood(r, c, direction='right') # A veces a la derecha
            if val != "N/A": return val

        # Estrategia 2: Buscar al lado de "TOTAL FOB" o "TOTAL" (derecha)
        r, c = self._find_coordinates(["TOTAL FOB", "TOTAL VALUE"])
        if r is not None:
            # Escanear hasta 5 celdas a la derecha buscando texto (USD, EUR, DÓLAR)
            val = self._scan_neighborhood(r, c, direction='right', max_steps=8)
            return val
            
        return "N/A"

    def extract_products_table(self):
        """
        Extrae productos con tolerancia a filas vacías intermedias.
        """
        r_qty, c_qty = self._find_coordinates(["CANTIDAD", "QTY", "QUANTITY"])
        r_desc, c_desc = self._find_coordinates(["DESCRIPCION", "DESCRIPTION", "MERCHANDISE DESCRIPTION"])
        r_price, c_price = self._find_coordinates(["PRECIO UNIT", "UNIT PRICE", "PRECIO"])
        r_total, c_total = self._find_coordinates(["TOTAL", "TOTAL LINEA"])

        if r_desc is None:
            return []

        products = []
        # Empezamos una fila debajo del encabezado más "profundo" encontrado
        start_r = max(r for r in [r_qty, r_desc, r_price] if r is not None) + 1
        
        empty_rows_patience = 0 # Contador para tolerar filas vacías
        max_patience = 3 # Permitir hasta 3 filas vacías antes de cortar
        
        current_r = start_r
        while current_r < len(self.raw_data):
            desc_val = str(self.raw_data[current_r][c_desc]).strip() if c_desc is not None else ""
            
            # Chequeos de parada
            is_stop_word = any(x in desc_val.upper() for x in ["TOTAL", "OBSERVACIONES", "NOTES", "SUBTOTAL"])
            
            if not desc_val:
                # Si la celda de descripción está vacía, aumentamos paciencia
                empty_rows_patience += 1
                if empty_rows_patience > max_patience:
                    break # Se acabaron las filas de datos
            elif is_stop_word:
                break
            else:
                # Encontramos datos, reiniciamos paciencia
                empty_rows_patience = 0
                
                qty_val = self.raw_data[current_r][c_qty] if c_qty is not None else "0"
                price_val = self.raw_data[current_r][c_price] if c_price is not None else "0"
                total_val = self.raw_data[current_r][c_total] if c_total is not None else "0"
                
                # Limpieza básica de NaN en celdas numéricas
                qty_val = "" if qty_val == "nan" else qty_val
                price_val = "" if price_val == "nan" else price_val
                total_val = "" if total_val == "nan" else total_val

                products.append({
                    "CANTIDAD": qty_val,
                    "DESCRIPCION": desc_val,
                    "PRECIO UNITARIO": price_val,
                    "TOTAL LINEA": total_val
                })
            
            current_r += 1
            
        return products
    
    def extract_observations(self):
        """Busca observaciones al final del documento."""
        r, c = self._find_coordinates(["OBSERVACIONES", "OBSERVATIONS", "NOTES", "COMENTARIOS"])
        if r is not None:
            # Intentar leer abajo
            val = self._scan_neighborhood(r, c, direction='down', max_steps=2)
            if val != "N/A": return val
            # Intentar leer a la derecha
            val = self._scan_neighborhood(r, c, direction='right', max_steps=5)
            return val
        return ""

    def process(self):
        # 1. Extracción de Cabecera usando Neighborhood Scan
        r_cli, c_cli = self._find_coordinates(["CLIENTE", "CUSTOMER"])
        cliente = self._scan_neighborhood(r_cli, c_cli, direction='down')
        
        r_exp, c_exp = self._find_coordinates(["EXP", "EXP N°", "REF EXP"])
        exp = self._scan_neighborhood(r_exp, c_exp, direction='down')
        
        fecha_unificada = self.extract_date()
        
        # Condición de Venta
        r_cond, c_cond = self._find_coordinates(["CONDICION VENTA", "CONDICION DE VENTA", "TERMS OF SALE"])
        raw_cond = self._scan_neighborhood(r_cond, c_cond, direction='down')
        if raw_cond != "N/A":
            parts = re.split(r'\s*[-–]\s*', raw_cond)
            tipo_venta = parts[0].strip() if len(parts) > 0 else "N/A"
            incoterm = parts[1].strip() if len(parts) > 1 else "N/A"
        else:
            tipo_venta, incoterm = "N/A", "N/A"

        # Puertos
        r_pe, c_pe = self._find_coordinates(["PUERTO EMBARQUE", "PORT OF LOADING"])
        puerto_emb = self._scan_neighborhood(r_pe, c_pe, direction='down')
        
        r_pd, c_pd = self._find_coordinates(["PUERTO DESTINO", "PORT OF DESTINATION", "DISCHARGING PORT"])
        puerto_dest = self._scan_neighborhood(r_pd, c_pd, direction='down')
        
        moneda = self.extract_currency()
        observaciones = self.extract_observations()

        # 2. Extracción de Productos
        products = self.extract_products_table()
        
        # 3. Construcción Flat Table
        output_rows = []
        header_data = {
            "CLIENTE": cliente,
            "EXP": exp,
            "FECHA": fecha_unificada,
            "TIPO DE VENTA": tipo_venta,
            "INCOTERM": incoterm,
            "PUERTO EMBARQUE": puerto_emb,
            "PUERTO DESTINO": puerto_dest,
            "MONEDA": moneda,
            "OBSERVACIONES": observaciones
        }
        
        if not products:
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
# INTERFAZ DE USUARIO STREAMLIT (Igual que antes)
# ==========================================
def main():
    st.set_page_config(page_title="Extractor de Facturas", page_icon="📄", layout="wide")
    
    st.title("🤖 Extractor Inteligente de Facturas (V2.0)")
    st.markdown("""
    **Mejoras V2.0:** Detección de filas vacías en tablas, búsqueda de moneda en celdas lejanas y formateo automático de fechas.
    """)

    uploaded_files = st.file_uploader("Sube tus archivos Excel", type=['xlsx', 'xls'], accept_multiple_files=True)

    if uploaded_files:
        all_data = []
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                df_raw = pd.read_excel(file, header=None)
                parser = InvoiceParser(df_raw)
                df_result = parser.process()
                df_result.insert(0, "ARCHIVO_ORIGEN", file.name)
                all_data.append(df_result)
            except Exception as e:
                st.error(f"❌ Error en {file.name}: {str(e)}")
            progress_bar.progress((i + 1) / len(uploaded_files))

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            st.success("✅ Procesamiento completado")
            st.dataframe(final_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            st.download_button("📥 Descargar Reporte", output.getvalue(), "reporte_exportacion.xlsx")

if __name__ == "__main__":
    main()
