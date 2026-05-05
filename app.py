import streamlit as st
import pandas as pd
import numpy as np
import gc
import plotly.express as px

st.set_page_config(page_title="LSKD Allocation Model", layout="wide")
st.title("📦 Sistema de Asignación Semanal: LSKD")

# --- PANEL LATERAL (CONTROLES) ---
st.sidebar.header("⚙️ Parámetros de Asignación")

# 1. Reglas de Bodega
st.sidebar.subheader("1. Límite de Bodega")
max_send_pct = st.sidebar.slider("Límite a Enviar (% de Bodega)", min_value=10.0, max_value=40.0, value=30.0, step=1.0)
flex_margin = st.sidebar.slider("Margen de Flexibilidad (%)", min_value=0.0, max_value=5.0, value=3.0, step=0.5)

# 2. Regla de Clima
st.sidebar.subheader("2. Filtro de Temporada")
st.sidebar.markdown("Define a qué clima se enviarán los productos marcados como 'CLIMATE SPECIFIC'.")
temporada_actual = st.sidebar.radio("Temporada Objetivo:", ["verano", "invierno", "Ambos (Ignorar regla)"])

# 3. Pesos de Tiendas
st.sidebar.subheader("3. Pesos por Categoría")
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
    st.warning("⏱️ Procesando algoritmo de distribución avanzado...")
    
    try:
        # --- PASO 1: CARGA Y LIMPIEZA ---
        # 1A. Newness
        df_newness = pd.read_excel(uploaded_file, sheet_name='Newness', engine='openpyxl')
        df_newness = df_newness.dropna(how='all', axis=0)
        df_newness.columns = df_newness.columns.astype(str).str.strip().str.replace(' ', '_')
        
        # 1B. Store Grading
        df_stores = pd.read_excel(uploaded_file, sheet_name='Store_Grading', engine='openpyxl')
        df_stores = df_stores.dropna(how='all', axis=0)
        df_stores.columns = df_stores.columns.astype(str).str.strip()
        
        # 1C. Size Curve
        df_curve = pd.read_excel(uploaded_file, sheet_name='Size_Curve', engine='openpyxl')
        df_curve.columns = df_curve.columns.astype(str).str.strip()

        del uploaded_file
        gc.collect()

        # --- PASO 2: MAPEO DE DICCIONARIOS ---
        # Diccionario de grados (A,B,C,D) y climas de cada tienda
        store_grades = pd.Series(df_stores['Womens Allocation Grade'].values, index=df_stores['Store']).to_dict()
        store_climates = pd.Series(df_stores['Climate'].astype(str).str.lower().str.strip().values, index=df_stores['Store']).to_dict()
        tiendas_destino = df_stores['Store'].dropna().tolist()

        # --- PASO 3: MOTOR MATEMÁTICO ---
        df_resultado = df_newness[['SKU', 'Product_Name', 'Size', 'Gender', 'Gender_&_Category', 'LSKD_DC_SOH', 'Grade']].copy()
        df_resultado['LSKD_DC_SOH'] = pd.to_numeric(df_resultado['LSKD_DC_SOH'], errors='coerce').fillna(0)

        # A) Cálculo de Límite Base
        limite_absoluto = (max_send_pct + flex_margin) / 100.0
        df_resultado['Max_Allocable'] = df_resultado['LSKD_DC_SOH'] * limite_absoluto

        # B) Extracción del Multiplicador de Curva de Tallas
        def obtener_multiplicador(row):
            try:
                talla = str(row['Size']).strip()
                categoria = str(row['Gender_&_Category']).strip()
                # Buscar la fila de la talla en la matriz
                fila_curva = df_curve[df_curve['SIZE'].astype(str).str.strip() == talla]
                if not fila_curva.empty and categoria in fila_curva.columns:
                    valor = fila_curva.iloc[0][categoria]
                    return float(valor) if pd.notna(valor) else 1.0
            except:
                pass
            return 1.0 # Multiplicador neutro si no encuentra cruce

        df_resultado['Curve_Multiplier'] = df_resultado.apply(obtener_multiplicador, axis=1)

        # C) Distribución e Inyección de Reglas por Tienda
        for tienda in tiendas_destino:
            grade_tienda = store_grades.get(tienda, 'C')
            peso_tienda = dict_pesos.get(grade_tienda, 0)
            clima_tienda = store_climates.get(tienda, '')

            # Aplicar peso de tienda + multiplicador de la curva de tallas
            unidades_calculadas = np.floor(df_resultado['Max_Allocable'] * peso_tienda * df_resultado['Curve_Multiplier'])
            df_resultado[tienda] = unidades_calculadas
            
            # REGLA 1: Exclusión "TOP TIER"
            mask_toptier = (df_resultado['Grade'] == 'TOP TIER') & (grade_tienda in ['C', 'D'])
            df_resultado.loc[mask_toptier, tienda] = 0

            # REGLA 2: Exclusión "CLIMATE SPECIFIC"
            if temporada_actual != "Ambos (Ignorar regla)":
                # Si el producto tiene restricción de clima y la tienda NO es del clima seleccionado, mandamos 0
                mask_climate = (df_resultado['Grade'] == 'CLIMATE SPECIFIC') & (clima_tienda != temporada_actual)
                df_resultado.loc[mask_climate, tienda] = 0

        st.success("✔️ Motor procesado: Curvas de Tallas aplicadas y Filtros de Clima ejecutados.")
        
        # --- PASO 4: MÉTRICAS VISUALES (PLOTLY) ---
        st.markdown("---")
        st.subheader("📊 Métricas Globales de la Semana")
        
        # Cálculos de resumen
        total_inventario = df_resultado['LSKD_DC_SOH'].sum()
        df_solo_tiendas = df_resultado[tiendas_destino]
        total_asignado = df_solo_tiendas.sum().sum()
        
        # 1. Tarjetas de métricas (KPIs)
        col1, col2, col3 = st.columns(3)
        col1.metric("📦 Inventario Total (DC SOH)", f"{int(total_inventario):,}")
        col2.metric("🚚 Unidades a Distribuir", f"{int(total_asignado):,}")
        pct_global = (total_asignado / total_inventario * 100) if total_inventario > 0 else 0
        col3.metric("🎯 % de Asignación Global", f"{pct_global:.1f}%")

        # 2. Gráfico interactivo
        st.markdown("##### Unidades Asignadas por Tienda")
        asignacion_por_tienda = df_solo_tiendas.sum().reset_index()
        asignacion_por_tienda.columns = ['Tienda', 'Unidades']
        asignacion_por_tienda = asignacion_por_tienda.sort_values(by='Unidades', ascending=False)
        
        fig = px.bar(
            asignacion_por_tienda, 
            x='Tienda', 
            y='Unidades',
            color='Unidades',
            color_continuous_scale='Blues',
            text_auto=True
        )
        fig.update_layout(xaxis_tickangle=-45, showlegend=False, margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

        # --- PASO 5: VISTA PREVIA Y DESCARGA ---
        st.markdown("---")
        st.subheader("📋 Matriz de Asignación Final")
        st.dataframe(df_resultado)

        csv = df_resultado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar Matriz (CSV)",
            data=csv,
            file_name='Asignacion_LSKD_Inteligente.csv',
            mime='text/csv',
        )

    except Exception as e:
        st.error(f"❌ Ocurrió un error en el cálculo: {e}")
