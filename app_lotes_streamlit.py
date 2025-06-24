import streamlit as st
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC
from sklearn.pipeline import Pipeline
import io # Necesario para la descarga de archivos en memoria

# --- Cargar y entrenar el modelo (usamos el caché de Streamlit para eficiencia) ---
# Esta función solo se ejecutará una vez, la primera vez que se carga la app.
@st.cache_resource
def cargar_y_entrenar_modelo():
    print("Cargando datos y entrenando modelo por primera vez...")
    try:
        df_train = pd.read_excel('datos_contables.xlsx', header=None, names=['numero_cuenta', 'descripcion', 'nombre_cuenta'])
        df_train.dropna(subset=['descripcion', 'nombre_cuenta'], inplace=True)
        df_train['descripcion'] = df_train['descripcion'].str.lower()
        counts = df_train['nombre_cuenta'].value_counts()
        df_filtrado = df_train[df_train['nombre_cuenta'].isin(counts[counts >= 2].index)]
        
        if df_filtrado.empty:
            return None

        X = df_filtrado['descripcion']
        y = df_filtrado['nombre_cuenta']
        model = Pipeline([
            ('vectorizer', TfidfVectorizer(ngram_range=(1, 2), sublinear_tf=True)),
            ('classifier', LinearSVC(random_state=42, class_weight='balanced'))
        ])
        model.fit(X, y)
        print("Modelo entrenado.")
        return model
    except FileNotFoundError:
        return None

modelo = cargar_y_entrenar_modelo()

# --- Construcción de la Interfaz Gráfica ---
st.title('🤖 Automatizador de Clasificación Contable')
st.write("Sube un archivo Excel con los detalles de tus facturas y la herramienta agregará una columna con la cuenta contable sugerida.")

if modelo is None:
    st.error("Error: No se pudo encontrar el archivo 'datos_contables.xlsx' para entrenar el modelo. Asegúrate de que esté en la misma carpeta.")
else:
    # 1. Widget para subir el archivo
    uploaded_file = st.file_uploader("Elige un archivo Excel (.xlsx)", type="xlsx")

    if uploaded_file is not None:
        st.success("¡Archivo cargado exitosamente!")
        df_a_clasificar = pd.read_excel(uploaded_file)
        
        st.write("Previsualización de los datos cargados:")
        st.dataframe(df_a_clasificar.head())

        # 2. Pedir al usuario el nombre de la columna con las descripciones
        nombre_columna = st.text_input("Por favor, escribe el nombre exacto de la columna que contiene las descripciones (ej: Glosa, Descripción):")

        # 3. Botón para iniciar la clasificación
        if st.button('Clasificar Archivo'):
            if nombre_columna and nombre_columna in df_a_clasificar.columns:
                with st.spinner('Clasificando transacciones... esto puede tardar unos segundos.'):
                    # Aplicar el modelo a la columna especificada
                    descripciones = df_a_clasificar[nombre_columna].astype(str).str.lower()
                    predicciones = modelo.predict(descripciones)
                    
                    # Añadir las predicciones como una nueva columna
                    df_a_clasificar['cuenta_sugerida'] = predicciones

                    st.write("¡Clasificación completada! Previsualización del resultado:")
                    st.dataframe(df_a_clasificar.head())

                    # 4. Ofrecer el archivo para descargar
                    # Convertimos el dataframe a un objeto de Excel en memoria
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_a_clasificar.to_excel(writer, index=False, sheet_name='Clasificaciones')
                    
                    st.download_button(
                        label="📥 Descargar Excel con Clasificaciones",
                        data=output.getvalue(),
                        file_name="resultado_clasificado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error(f"Error: La columna '{nombre_columna}' no se encuentra en el archivo. Por favor, verifica el nombre.")