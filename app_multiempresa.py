import streamlit as st
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC
from sklearn.pipeline import Pipeline
import os
import io

# --- Configuración de la página ---
st.set_page_config(
    page_title="Asistente Contable Multi-Empresa",
    layout="wide"
)

# --- Funciones Principales ---

# Esta función se encarga de cargar y entrenar el modelo para UNA empresa específica.
# @st.cache_resource le dice a Streamlit que "recuerde" el resultado de esta función.
# Si volvemos a llamar la función con el mismo nombre_archivo, usará el modelo que ya
# tiene en memoria en lugar de entrenarlo de nuevo. ¡Esto la hace súper rápida!
@st.cache_resource
def cargar_y_entrenar_modelo(ruta_archivo_empresa):
    """Carga los datos de una empresa y entrena un modelo para ella."""
    try:
        df_train = pd.read_excel(ruta_archivo_empresa, header=None, names=['numero_cuenta', 'descripcion', 'nombre_cuenta'])
        df_train.dropna(subset=['descripcion', 'nombre_cuenta'], inplace=True)
        # Aseguramos que la descripción sea siempre texto
        df_train['descripcion'] = df_train['descripcion'].astype(str).str.lower()
        counts = df_train['nombre_cuenta'].value_counts()
        df_filtrado = df_train[df_train['nombre_cuenta'].isin(counts[counts >= 2].index)]
        
        if df_filtrado.empty:
            st.error(f"El archivo de entrenamiento para esta empresa no tiene suficientes datos válidos para entrenar un modelo.")
            return None

        X = df_filtrado['descripcion']
        y = df_filtrado['nombre_cuenta']
        model = Pipeline([
            ('vectorizer', TfidfVectorizer(ngram_range=(1, 2), sublinear_tf=True)),
            ('classifier', LinearSVC(random_state=42, class_weight='balanced'))
        ])
        model.fit(X, y)
        print(f"Modelo entrenado exitosamente para {ruta_archivo_empresa}")
        return model
    except FileNotFoundError:
        st.error(f"No se pudo encontrar el archivo de entrenamiento: {ruta_archivo_empresa}")
        return None
    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo de entrenamiento: {e}")
        return None

# --- Interfaz de Usuario ---

st.title('🤖 Asistente Contable Multi-Empresa')
st.write("Selecciona una empresa para cargar su modelo de clasificación y luego sube un archivo para procesar.")

# 1. Detectar automáticamente las empresas disponibles
#    Buscamos todos los archivos .xlsx en la carpeta 'datos_empresas'
try:
    lista_archivos = [f for f in os.listdir('datos_empresas') if f.endswith('.xlsx')]
    # Limpiamos los nombres para que se vean bien en el selector
    # "constructora_xyz.xlsx" -> "Constructora Xyz"
    nombres_empresas = {
        archivo: " ".join(word.capitalize() for word in archivo.replace('.xlsx', '').split('_'))
        for archivo in lista_archivos
    }
except FileNotFoundError:
    st.error("Error crítico: No se encontró la carpeta 'datos_empresas'. Por favor, asegúrate de que exista.")
    lista_archivos = []
    nombres_empresas = {}

if not lista_archivos:
    st.warning("No se encontraron archivos de entrenamiento de empresas en la carpeta 'datos_empresas'.")
else:
    # 2. Crear el menú desplegable para seleccionar la empresa
    archivo_seleccionado = st.selectbox(
        "Selecciona la Empresa",
        options=lista_archivos,
        format_func=lambda x: nombres_empresas[x] # Muestra el nombre bonito en el menú
    )

    if archivo_seleccionado:
        st.info(f"Has seleccionado: **{nombres_empresas[archivo_seleccionado]}**")
        ruta_completa_archivo = os.path.join('datos_empresas', archivo_seleccionado)
        
        # Cargar el modelo para la empresa seleccionada (usará el caché si ya está cargado)
        modelo_activo = cargar_y_entrenar_modelo(ruta_completa_archivo)

        # 3. Mostrar el resto de la interfaz SOLO si el modelo se cargó correctamente
        if modelo_activo:
            st.header("Cargar archivo para clasificar")
            uploaded_file = st.file_uploader("Elige un archivo Excel (.xlsx)", type="xlsx", key=archivo_seleccionado)

            if uploaded_file:
                df_a_clasificar = pd.read_excel(uploaded_file)
                st.write("Previsualización de los datos cargados:")
                st.dataframe(df_a_clasificar.head())
                
                nombre_columna = st.text_input("Escribe el nombre de la columna con las descripciones (ej: Glosa):")
                
                if st.button('Clasificar Archivo'):
                    if nombre_columna and nombre_columna in df_a_clasificar.columns:
                        with st.spinner('Clasificando...'):
                            descripciones = df_a_clasificar[nombre_columna].astype(str).str.lower()
                            predicciones = modelo_activo.predict(descripciones)
                            df_a_clasificar['cuenta_sugerida'] = predicciones
                            
                            st.success("¡Clasificación completada!")
                            st.dataframe(df_a_clasificar.head())
                            
                            # Ofrecer el archivo para descargar
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df_a_clasificar.to_excel(writer, index=False, sheet_name='Clasificaciones')
                            
                            st.download_button(
                                label="📥 Descargar Resultado",
                                data=output.getvalue(),
                                file_name=f"resultado_{archivo_seleccionado}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error(f"La columna '{nombre_columna}' no se encuentra en el archivo.")