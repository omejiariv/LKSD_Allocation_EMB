import streamlit as st
import pandas as pd
import numpy as np
import gc
import plotly.express as px

st.set_page_config(page_title="LSKD Allocation Model", layout="wide")

# --- SISTEMA DE TRADUCCIONES (i18n) ---
# Selector de idioma en la parte superior del sidebar
idioma = st.sidebar.selectbox("🌐 Language / Idioma", ["Español", "English"])

# Diccionario general de textos
t = {
    "Español": {
        "title": "📦 Sistema de Asignación Semanal: LSKD",
        "sidebar_header": "⚙️ Parámetros de Asignación",
        "limit_bodega_title": "1. Límite de Bodega",
        "limit_bodega_slider": "Límite a Enviar (% de Bodega)",
        "flex_margin": "Margen de Flexibilidad (%)",
        "season_filter_title": "2. Filtro de Temporada",
        "season_filter_desc": "Define a qué clima se enviarán los productos marcados como 'CLIMATE SPECIFIC'.",
        "season_target": "Temporada Objetivo:",
        "season_summer": "Verano",
        "season_winter": "Invierno",
        "season_both": "Ambos (Ignorar regla)",
        "weights_title": "3. Pesos por Categoría",
        "weight_a": "Peso Tiendas A",
        "weight_b": "Peso Tiendas B",
        "weight_c": "Peso Tiendas C",
        "weight_d": "Peso Tiendas D",
        "upload_file": "Sube el archivo LSKD_Newness_EMB.xlsx",
        "processing": "⏱️ Procesando algoritmo de distribución avanzado...",
        "success_msg": "✔️ Motor procesado: Curvas de Tallas aplicadas y Filtros de Clima ejecutados.",
        "metrics_title": "📊 Métricas Globales de la Semana",
        "metric_inv": "📦 Inventario Total (DC SOH)",
        "metric_dist": "🚚 Unidades a Distribuir",
        "metric_pct": "🎯 % de Asignación Global",
        "chart_title": "##### Unidades Asignadas por Tienda",
        "matrix_title": "📋 Matriz de Asignación Final",
        "download_btn": "📥 Descargar Matriz (CSV)",
        "error_msg": "❌ Ocurrió un error en el cálculo: "
    },
    "English": {
        "title": "📦 LSKD Weekly Allocation System",
        "sidebar_header": "⚙️ Allocation Parameters",
        "limit_bodega_title": "1. DC SOH Limit",
        "limit_bodega_slider": "Send Limit (% of DC)",
        "flex_margin": "Flexibility Margin (%)",
        "season_filter_title": "2. Season Filter",
        "season_filter_desc": "Define the target climate for 'CLIMATE SPECIFIC' products.",
        "season_target": "Target Season:",
        "season_summer": "Summer",
        "season_winter": "Winter",
        "season_both": "Both (Ignore rule)",
        "weights_title": "3. Category Weights",
        "weight_a": "Tier A Weight",
        "weight_b": "Tier B Weight",
        "weight_c": "Tier C Weight",
        "weight_d": "Tier D Weight",
        "upload_file": "Upload LSKD_Newness_EMB.xlsx file",
        "processing": "⏱️ Processing advanced allocation algorithm...",
        "success_msg": "✔️ Engine processed: Size Curves applied and Climate Filters executed.",
        "metrics_title": "📊 Weekly Global Metrics",
        "metric_inv": "📦 Total Inventory (DC SOH)",
        "metric_dist": "🚚 Units to Allocate",
        "metric_pct": "🎯 Global Allocation %",
        "chart_title": "##### Units Allocated per Store",
        "matrix_title": "📋 Final Allocation Matrix",
        "download_btn": "📥 Download Matrix (CSV)",
        "error_msg": "❌ An error occurred during calculation: "
    }
}

# Variable rápida para acceder a los textos del idioma seleccionado
txt = t[idioma]

st.title(txt["title"])

# --- PANEL LATERAL (CONTROLES) ---
st.sidebar.header(txt["sidebar_header"])

# 1. Reglas de Bodega
st.sidebar.subheader(txt["limit_bodega_title"])
max_send_pct = st.sidebar.slider(txt["limit_bodega_slider"], min_value=10.0, max_value=40.0, value=30.0, step=1.0)
flex_margin = st.sidebar.slider(txt["flex_margin"], min_value=0.0, max_value=5.0, value=3.0, step=0.5)

# 2. Regla de Clima
st.sidebar.subheader(txt["season_filter_title"])
st.sidebar.markdown(txt["season_filter_desc"])

# Opciones de clima traducidas pero mapeadas al backend en español (para cruzar con el Excel)
opciones_clima_ui = [txt["season_summer"], txt["season_winter"], txt["season_both"]]
seleccion_clima_ui = st.sidebar.radio(txt["season_target"], opciones_clima_ui)

# Mapeo al valor real del backend
if seleccion_clima_ui == txt["season_summer"]:
    temporada_backend = "verano"
elif seleccion_clima_ui == txt["season_winter"]:
    temporada_backend = "invierno"
else:
    temporada_backend = "ambos"

# 3. Pesos de Tiendas
st.sidebar.subheader(txt["weights_title"])
peso_a = st.sidebar.number_input(txt["weight_a"], value=40)
peso_b = st.sidebar.number_input(txt["weight_b"], value=30)
peso_c = st.sidebar.number_input(txt["weight_c"], value=20)
peso_d = st.sidebar.number_input(txt["weight_d"], value=10)

total_peso = peso_a + peso_b + peso_c + peso_d
if total_peso == 0: total_peso = 1 
dict_pesos = {
    'A': peso_a / total_peso,
    'B': peso_b / total_peso,
    'C': peso_c / total_peso,
    'D': peso_d / total_peso
}

# --- APLICACIÓN PRINCIPAL ---
uploaded_file = st.file_uploader(txt["upload_file"], type=["xlsx"])

if uploaded_file:
    st.warning(txt["processing"])
    
    try:
        # --- PASO 1: CARGA Y LIMPIEZA ---
        df_newness = pd.read_excel(uploaded_file, sheet_name='Newness', engine='openpyxl')
        df_newness = df_newness.dropna(how='all', axis=0)
        df_newness.columns = df_newness.columns.astype(str).str.strip().str.replace(' ', '_')
        
        df_stores = pd.read_excel(uploaded_file, sheet_name='Store_Grading', engine='openpyxl')
        df_stores = df_stores.dropna(how='all', axis=0)
        df_stores.columns = df_stores.columns.astype(str).str.strip()
        
        df_curve = pd.read_excel(uploaded_file, sheet_name='Size_Curve', engine='openpyxl')
        df_curve.columns = df_curve.columns.astype(str).str.strip()

        del uploaded_file
        gc.collect()

        # --- PASO 2: MAPEO DE DICCIONARIOS ---
        store_grades = pd.Series(df_stores['Womens_Allocation_Grade'].values, index=df_stores['Store']).to_dict()
        store_climates = pd.Series(df_stores['Climate'].astype(str).str.lower().str.strip().values, index=df_stores['Store']).to_dict()
        tiendas_destino = df_stores['Store'].dropna().tolist()

        # --- PASO 3: MOTOR MATEMÁTICO (MÉTODO DEL RESTO MAYOR) ---
        df_resultado = df_newness[['SKU', 'Product_Name', 'Size', 'Gender', 'Gender_&_Category', 'LSKD_DC_SOH', 'Grade']].copy()
        df_resultado['LSKD_DC_SOH'] = pd.to_numeric(df_resultado['LSKD_DC_SOH'], errors='coerce').fillna(0)

        def obtener_multiplicador(row):
            try:
                talla = str(row['Size']).strip()
                categoria = str(row['Gender_&_Category']).strip()
                fila_curva = df_curve[df_curve['SIZE'].astype(str).str.strip() == talla]
                if not fila_curva.empty and categoria in fila_curva.columns:
                    valor = fila_curva.iloc[0][categoria]
                    return float(valor) if pd.notna(valor) else 1.0
            except:
                pass
            return 1.0

        df_resultado['Curve_Multiplier'] = df_resultado.apply(obtener_multiplicador, axis=1)
        df_resultado['Norm_Curve'] = df_resultado.groupby('Product_Name')['Curve_Multiplier'].transform(lambda x: x / x.mean() if x.mean() > 0 else 1)

        limite_absoluto = (max_send_pct + flex_margin) / 100.0
        
        # Calculamos la bolsa exacta de unidades que DEBEMOS repartir por fila
        df_resultado['Max_Allocable'] = np.clip(df_resultado['LSKD_DC_SOH'] * limite_absoluto * df_resultado['Norm_Curve'], 0, df_resultado['LSKD_DC_SOH'])

        df_pesos = pd.DataFrame(index=df_resultado.index, columns=tiendas_destino)

        # Matriz de validación (qué tienda califica)
        for tienda in tiendas_destino:
            grade_tienda = store_grades.get(tienda, 'C')
            peso_tienda = dict_pesos.get(grade_tienda, 0)
            clima_tienda = store_climates.get(tienda, '')

            peso_actual = pd.Series(peso_tienda, index=df_resultado.index)

            mask_toptier = (df_resultado['Grade'] == 'TOP TIER') & (grade_tienda in ['C', 'D'])
            peso_actual = np.where(mask_toptier, 0, peso_actual)

            if temporada_backend != "ambos":
                mask_climate = (df_resultado['Grade'] == 'CLIMATE SPECIFIC') & (clima_tienda != temporada_backend)
                peso_actual = np.where(mask_climate, 0, peso_actual)

            df_pesos[tienda] = peso_actual

        # DISTRIBUCIÓN AVANZADA: Método del Resto Mayor (Largest Remainder Method)
        for idx in df_resultado.index:
            # Unidades totales a repartir en esta fila (Talla específica)
            max_units = int(np.round(df_resultado.loc[idx, 'Max_Allocable']))
            pesos_row = df_pesos.loc[idx, tiendas_destino].values.astype(float)
            suma_pesos = np.sum(pesos_row)

            if suma_pesos == 0 or max_units <= 0:
                df_resultado.loc[idx, tiendas_destino] = 0
                continue

            # 1. Asignación exacta con decimales
            exact_alloc = max_units * (pesos_row / suma_pesos)
            
            # 2. Asignación base (Parte entera)
            alloc = np.floor(exact_alloc).astype(int)
            
            # 3. Unidades sobrantes por culpa de los decimales
            remainder = int(max_units - np.sum(alloc))
            fractions = exact_alloc - alloc

            # 4. Repartir el sobrante a las tiendas con el decimal más alto
            if remainder > 0:
                # Encontramos los índices de las tiendas con los decimales más cercanos a 1
                indices = np.argsort(fractions)[-remainder:]
                for i in indices:
                    alloc[i] += 1

            # Guardar en el DataFrame
            df_resultado.loc[idx, tiendas_destino] = alloc
        
        # --- PASO 4: MÉTRICAS VISUALES (PLOTLY) ---
        st.markdown("---")
        st.subheader(txt["metrics_title"])
        
        total_inventario = df_resultado['LSKD_DC_SOH'].sum()
        df_solo_tiendas = df_resultado[tiendas_destino]
        total_asignado = df_solo_tiendas.sum().sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric(txt["metric_inv"], f"{int(total_inventario):,}")
        col2.metric(txt["metric_dist"], f"{int(total_asignado):,}")
        pct_global = (total_asignado / total_inventario * 100) if total_inventario > 0 else 0
        col3.metric(txt["metric_pct"], f"{pct_global:.1f}%")

        st.markdown(txt["chart_title"])
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
        st.subheader(txt["matrix_title"])
        st.dataframe(df_resultado)

        csv = df_resultado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=txt["download_btn"],
            data=csv,
            file_name='Asignacion_LSKD_Inteligente.csv',
            mime='text/csv',
        )

    except Exception as e:
        st.error(f"{txt['error_msg']}{e}")
