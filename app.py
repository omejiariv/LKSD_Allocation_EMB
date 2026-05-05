import streamlit as st
import pandas as pd
import numpy as np
import gc # Recolector de basura para liberar memoria RAM

st.set_page_config(page_title="LSKD Allocation Model", layout="wide")
st.title("📦 Sistema de Asignación Semanal: LSKD")

uploaded_file = st.file_uploader("Sube el archivo LSKD_Newness_EMB.xlsx", type=["xlsx"])

if uploaded_file:
    st.warning("⏱️ Iniciando diagnóstico de memoria...")
    
    try:
        # PUNTO DE CONTROL 1
        st.info("Paso 1: Leyendo hoja 'Newness'...")
        df_newness = pd.read_excel(uploaded_file, sheet_name='Newness', engine='openpyxl')
        
        # 1. Limpiar SOLO las FILAS vacías (quitamos el axis=1 para no borrar columnas vitales)
        df_newness = df_newness.dropna(how='all', axis=0)
        
        # 2. Eliminar espacios invisibles de los títulos
        df_newness.columns = df_newness.columns.astype(str).str.strip()
        
        st.success(f"✔️ 'Newness' cargada: {df_newness.shape[0]} filas.")
        
        # PUNTO DE CONTROL 2
        st.info("Paso 2: Leyendo hoja 'Store grading'...")
        df_stores = pd.read_excel(uploaded_file, sheet_name='Store grading', engine='openpyxl')
        df_stores = df_stores.dropna(how='all', axis=0)
        st.success(f"✔️ 'Store grading' cargada.")

        # PUNTO DE CONTROL 3
        st.info("Paso 3: Leyendo hoja 'Size Curve'...")
        df_curve = pd.read_excel(uploaded_file, sheet_name='Size Curve', engine='openpyxl')
        st.success(f"✔️ 'Size Curve' cargada.")
        
        # Liberar memoria de Pandas
        del uploaded_file
        gc.collect()

        st.info("Paso 4: Procesando motor matemático...")
        
        # LÓGICA SIMPLIFICADA PARA PRUEBA DE MEMORIA
        df_stores.columns = df_stores.columns.astype(str).str.strip() # Limpiamos también las tiendas por si acaso
        tiendas_destino = df_stores['Store'].dropna().tolist()
        
        # Extraer solo las columnas necesarias (ahora sin espacios fantasma)
        df_resultado = df_newness[['SKU', 'Product Name', 'Size', 'Gender', 'LSKD DC SOH', 'Grade']].copy()
        
        for tienda in tiendas_destino:
            df_resultado[tienda] = 1 
            
        st.success("✔️ Motor matemático procesado con éxito.")
        
        # PUNTO DE CONTROL FINAL
        st.subheader("Vista Previa")
        st.dataframe(df_resultado.head(10)) # Mostramos solo 10 para no saturar el navegador
        
    except Exception as e:
        st.error(f"❌ Ocurrió un error: {e}")
