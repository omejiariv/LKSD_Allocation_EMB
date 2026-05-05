import streamlit as st
import pandas as pd
import numpy as np
import gc

st.set_page_config(page_title="LSKD Allocation Model", layout="wide")
st.title("📦 Sistema de Asignación Semanal: LSKD")

# --- PANEL LATERAL (CONTROLES) ---
st.sidebar.header("⚙️ Parámetros de Asignación")
st.sidebar.markdown("Ajusta las reglas matemáticas para esta semana:")

max_send_pct = st.sidebar.slider("Límite a Enviar (% de Bodega)", min_value=10.0, max_value=40.0, value=30.0, step=1.0)
flex_margin = st.sidebar.slider("Margen de Flexibilidad (%)", min_value=0.0, max_value=5.0, value=3.0, step=0.5)

st.sidebar.subheader("Pesos por Categoría de Tienda")
peso_a = st.sidebar.number_input("Peso Tiendas A", value=40)
peso_b = st.sidebar.number_input("Peso Tiendas B", value=30)
peso_c = st.sidebar.number_input("Peso Tiendas C", value=20)
peso_d = st.sidebar.number_input("Peso Tiendas D", value=10)

# Normalización matemática de los pesos
total_peso = peso_a + peso_b + peso_c + peso_d
if total_peso == 0: total_peso = 1 
dict_pesos = {
    'A': peso_a / total_peso,
    'B': peso_b / total_peso,
    'C': peso_c / total_peso,
    'D': peso_d / total_peso
}

# --- APLICACIÓN PRINCIPAL ---
uploaded_file = st.file_uploader("Sube el archivo LSKD_Newness_EMB.xlsx", type=["xlsx"])

if uploaded_file:
    st.warning("⏱️ Procesando algoritmo de distribución...")
    
    try:
        # PASO 1: Carga y Limpieza
        df_newness = pd.read_excel(uploaded_file, sheet_name='Newness', engine='openpyxl')
        df_newness = df_newness.dropna(how='all', axis=0)
        df_newness.columns = df_newness.columns.astype(str).str.strip().str.replace(' ', '_')
        
        # PASO 2: Tiendas
        df_stores = pd.read_excel(uploaded_file, sheet_name='Store_Grading', engine='openpyxl')
        df_stores = df_stores.dropna(how='all', axis=0)
        df_stores.columns = df_stores.columns.astype(str).str.strip()

        # Liberamos memoria
        del uploaded_file
        gc.collect()

        # PASO 3: MOTOR MATEMÁTICO
        # 1. Mapeo de la calificación de cada tienda (Usamos Womens Allocation Grade por defecto)
        store_grades = pd.Series(df_stores['Womens Allocation Grade'].values, index=df_stores['Store']).to_dict()
        tiendas_destino = df_stores['Store'].dropna().tolist()

        # 2. Extraer columnas base
        df_resultado = df_newness[['SKU', 'Product_Name', 'Size', 'Gender', 'LSKD_DC_SOH', 'Grade']].copy()
        
        # Asegurarnos de que el SOH es un número válido
        df_resultado['LSKD_DC_SOH'] = pd.to_numeric(df_resultado['LSKD_DC_SOH'], errors='coerce').fillna(0)

        # 3. Regla del 30% (Límite Máximo)
        limite_absoluto = (max_send_pct + flex_margin) / 100.0
        df_resultado['Max_Allocable'] = np.floor(df_resultado['LSKD_DC_SOH'] * limite_absoluto)

        # 4. Distribución por Tienda
        for tienda in tiendas_destino:
            # Identificar qué tipo de tienda es (A, B, C, o D)
            grade_tienda = store_grades.get(tienda, 'C') # Asume C si está en blanco
            peso_tienda = dict_pesos.get(grade_tienda, 0)
            
            # Asignar unidades matemáticamente
            df_resultado[tienda] = np.floor(df_resultado['Max_Allocable'] * peso_tienda)
            
            # REGLA: "No todo va a todas las tiendas"
            # Si el producto está marcado como TOP TIER, eliminamos la asignación a tiendas C y D
            mask_toptier = (df_resultado['Grade'] == 'TOP TIER') & (grade_tienda in ['C', 'D'])
            df_resultado.loc[mask_toptier, tienda] = 0

        st.success("✔️ Distribución generada con éxito basándose en las reglas de asignación.")
        
        # PASO 4: VISTA PREVIA Y DESCARGA
        st.subheader("Matriz de Asignación")
        st.dataframe(df_resultado)

        # Botón para descargar el Excel final
        csv = df_resultado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar Resultado (CSV)",
            data=csv,
            file_name='Asignacion_LSKD.csv',
            mime='text/csv',
        )

    except Exception as e:
        st.error(f"❌ Ocurrió un error en el cálculo: {e}")
