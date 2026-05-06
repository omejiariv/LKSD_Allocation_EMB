import streamlit as st
import pandas as pd
import numpy as np
import gc
import plotly.express as px
import io

st.set_page_config(page_title="LSKD Allocation Model", layout="wide")

# --- SISTEMA DE TRADUCCIONES (i18n) ---
idioma = st.sidebar.selectbox("🌐 Language / Idioma", ["Español", "English"])

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
        "success_msg": "Motor procesado con éxito.",
        "metrics_title": "📊 Métricas Globales de la Semana",
        "metric_inv": "📦 Inventario Total (DC SOH)",
        "metric_dist": "🚚 Unidades a Distribuir",
        "metric_pct": "🎯 % de Asignación Global",
        "chart_title": "##### Unidades Asignadas por Tienda",
        "matrix_title": "📋 Matriz de Asignación Final",
        "download_btn": "📥 Descargar Matriz (CSV)",
        "error_msg": "❌ Ocurrió un error en el cálculo: ",
        # TOOLTIPS ESPAÑOL
        "tt_limit": "Porcentaje base del inventario total disponible que se permite distribuir a las tiendas.",
        "tt_flex": "Margen adicional temporal. Permite absorber redondeos y subidas de demanda en tallas populares (ej. M o L) sin romper la regla global.",
        "tt_season": "Aplica una restricción a los SKU marcados como 'CLIMATE SPECIFIC'. Las tiendas cuyo clima no coincida recibirán 0 unidades.",
        "tt_weight": "Define qué proporción matemática del envío total se asignará a este nivel de tienda.",
        "tt_upload": "Carga el archivo Excel estandarizado. Debe contener las hojas: 'Newness', 'Store_Grading' y 'Size_Curve'.",
        # DOCUMENTACIÓN ESPAÑOL
        "doc_title": "📚 Documentación, Metodología e Insumos",
        "doc_content": """
        ### 📌 Resumen de la App
        Esta aplicación automatiza el proceso semanal de *Allocation* (Asignación de Inventario) para LSKD. Cruza el inventario disponible en bodega (SOH) con la calificación de cada tienda, respetando reglas estrictas de capacidad y adaptándose orgánicamente a la curva de tallas histórica.
        
        ### ⚙️ Metodología y Conceptos Clave
        * **Método del Resto Mayor (Largest Remainder):** Algoritmo utilizado en el paso final del reparto. Evita la "pérdida por redondeo hacia abajo", asegurando que las fracciones decimales sobrantes se sumen y se asignen como unidades enteras a las tiendas más cercanas al siguiente decimal. Esto garantiza que el porcentaje global exacto se cumpla siempre.
        * **Multiplicador de Curvas Dinámico (`obtener_multiplicador`):** El código cruza la columna `Size` y `Gender_&_Category` de tu producto con la hoja de `Size_Curve`. Si una talla "S" o "M" tiene un peso de 2.0 en la matriz, automáticamente infla la asignación para esas tallas antes de hacer el redondeo hacia abajo, respetando las proporciones orgánicas de venta de cada categoría.
        * **Selector de Temporada (`mask_climate`):** Botones interactivos para "Verano", "Invierno" o "Ambos". Si el producto de la semana dice `CLIMATE SPECIFIC` en su columna de Grado, el motor revisará el clima asignado a cada tienda en `Store_Grading` y dejará en cero (0) la asignación a las ciudades que no encajen.
        * **Filtro TOP TIER:** Si un producto tiene el grado `TOP TIER`, se excluye automáticamente de las tiendas con calificación C y D.
        * **Dashboards Integrados:** Añadimos `plotly.express` para inyectar KPIs visuales. Automáticamente consolidamos las sumas de toda la tabla para darte el total de bodega y un diagrama de barras interactivo (ordenado de mayor a menor) donde puedes pasar el ratón para ver exactamente cuántas unidades viajarán a cada tienda esta semana.

        ### 📥 Insumos Requeridos y Estructura
        La herramienta requiere la carga semanal de un archivo Excel (`.xlsx`) con **tres hojas estrictamente nombradas**:
        1. **`Newness`**: Base de datos de productos. Debe contener columnas como `SKU`, `Product Name`, `Size`, `Gender_&_Category`, `LSKD DC SOH` (Inventario) y `Grade`.
        2. **`Store_Grading`**: Base de datos de tiendas. Columnas requeridas: `Store`, `Womens_Allocation_Grade` (A, B, C, D) y `Climate` (verano/invierno).
        3. **`Size_Curve`**: Matriz de multiplicadores de demanda cruzando Talla vs. Categoría.
        """
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
        "success_msg": "Engine processed successfully.",
        "metrics_title": "📊 Weekly Global Metrics",
        "metric_inv": "📦 Total Inventory (DC SOH)",
        "metric_dist": "🚚 Units to Allocate",
        "metric_pct": "🎯 Global Allocation %",
        "chart_title": "##### Units Allocated per Store",
        "matrix_title": "📋 Final Allocation Matrix",
        "download_btn": "📥 Download Matrix (CSV)",
        "error_msg": "❌ An error occurred during calculation: ",
        # TOOLTIPS INGLÉS
        "tt_limit": "Base percentage of total available inventory allowed to be distributed to stores.",
        "tt_flex": "Additional buffer margin. Allows absorbing roundings and high demand in popular sizes (e.g., M or L) without breaking the global rule.",
        "tt_season": "Applies a restriction to SKUs marked as 'CLIMATE SPECIFIC'. Stores whose climate does not match will receive 0 units.",
        "tt_weight": "Defines what mathematical proportion of the total shipment goes to this store tier.",
        "tt_upload": "Upload the standardized Excel file. Must contain tabs: 'Newness', 'Store_Grading' and 'Size_Curve'.",
        # DOCUMENTACIÓN INGLÉS
        "doc_title": "📚 Documentation, Methodology & Inputs",
        "doc_content": """
        ### 📌 App Summary
        This application automates the weekly *Allocation* process for LSKD. It crosses the available warehouse inventory (SOH) with each store's grading, respecting strict capacity rules and adapting organically to the historical size curve.
        
        ### ⚙️ Methodology & Key Concepts
        * **Largest Remainder Method:** Algorithm used in the final distribution step. It avoids "rounding down loss", ensuring that remaining decimal fractions are summed and assigned as whole units to the stores closest to the next decimal. This guarantees exact global targets.
        * **Dynamic Curve Multiplier (`obtener_multiplicador`):** The code crosses the product's `Size` and `Gender_&_Category` with the `Size_Curve` sheet. If a size "S" or "M" has a weight of 2.0 in the matrix, it automatically inflates the allocation for those sizes before rounding down, respecting the organic sales proportions.
        * **Season Selector (`mask_climate`):** Interactive buttons for "Summer", "Winter", or "Both". If the weekly product is graded `CLIMATE SPECIFIC`, the engine checks the climate assigned to each store in `Store_Grading` and zeroes out the allocation to unmatched cities.
        * **TOP TIER Filter:** If a product is graded `TOP TIER`, it is automatically excluded from C and D graded stores.
        * **Integrated Dashboards:** We added `plotly.express` to inject visual KPIs. We automatically consolidate the sums of the entire table to give you total SOH and an interactive bar chart (sorted high to low) where you can hover to see exactly how many units will travel to each store this week.

        ### 📥 Required Inputs & Structure
        The tool requires the weekly upload of an Excel (`.xlsx`) file with **three strictly named sheets**:
        1. **`Newness`**: Product database. Must contain columns like `SKU`, `Product Name`, `Size`, `Gender_&_Category`, `LSKD DC SOH` (Inventory) and `Grade`.
        2. **`Store_Grading`**: Store database. Required columns: `Store`, `Womens_Allocation_Grade` (A, B, C, D) and `Climate` (verano/invierno).
        3. **`Size_Curve`**: Matrix of demand multipliers crossing Size vs. Category.
        """
    }
}

txt = t[idioma]
st.title(txt["title"])

# --- PANEL LATERAL (CONTROLES) ---
st.sidebar.header(txt["sidebar_header"])

st.sidebar.subheader(txt["limit_bodega_title"])
max_send_pct = st.sidebar.slider(txt["limit_bodega_slider"], min_value=10.0, max_value=40.0, value=30.0, step=1.0, help=txt["tt_limit"])
flex_margin = st.sidebar.slider(txt["flex_margin"], min_value=0.0, max_value=5.0, value=3.0, step=0.5, help=txt["tt_flex"])

st.sidebar.subheader(txt["season_filter_title"])
st.sidebar.markdown(txt["season_filter_desc"])

opciones_clima_ui = [txt["season_summer"], txt["season_winter"], txt["season_both"]]
seleccion_clima_ui = st.sidebar.radio(txt["season_target"], opciones_clima_ui, help=txt["tt_season"])

if seleccion_clima_ui == txt["season_summer"]:
    temporada_backend = "verano"
elif seleccion_clima_ui == txt["season_winter"]:
    temporada_backend = "invierno"
else:
    temporada_backend = "ambos"

st.sidebar.subheader(txt["weights_title"])
peso_a = st.sidebar.number_input(txt["weight_a"], value=40, help=txt["tt_weight"])
peso_b = st.sidebar.number_input(txt["weight_b"], value=30, help=txt["tt_weight"])
peso_c = st.sidebar.number_input(txt["weight_c"], value=20, help=txt["tt_weight"])
peso_d = st.sidebar.number_input(txt["weight_d"], value=10, help=txt["tt_weight"])

total_peso = peso_a + peso_b + peso_c + peso_d
if total_peso == 0: total_peso = 1 
dict_pesos = {
    'A': peso_a / total_peso,
    'B': peso_b / total_peso,
    'C': peso_c / total_peso,
    'D': peso_d / total_peso
}

# --- DOCUMENTACIÓN DESPLEGABLE (SIEMPRE VISIBLE) ---
with st.expander(txt["doc_title"]):
    st.markdown(txt["doc_content"])

# --- APLICACIÓN PRINCIPAL ---
uploaded_file = st.file_uploader(txt["upload_file"], type=["xlsx"], key="cargador_excel", help=txt["tt_upload"])

if uploaded_file:
    try:
        # Usamos st.spinner para que el mensaje desaparezca al terminar
        with st.spinner(txt["processing"]):
            df_newness = pd.read_excel(uploaded_file, sheet_name='Newness', engine='openpyxl')
            df_newness = df_newness.dropna(how='all', axis=0)
            df_newness.columns = df_newness.columns.astype(str).str.strip().str.replace(' ', '_')
            
            df_stores = pd.read_excel(uploaded_file, sheet_name='Store_Grading', engine='openpyxl')
            df_stores = df_stores.dropna(how='all', axis=0)
            df_stores.columns = df_stores.columns.astype(str).str.strip()
            
            df_curve = pd.read_excel(uploaded_file, sheet_name='Size_Curve', engine='openpyxl')
            df_curve.columns = df_curve.columns.astype(str).str.strip()

            gc.collect()

            store_grades = pd.Series(df_stores['Womens_Allocation_Grade'].values, index=df_stores['Store']).to_dict()
            store_climates = pd.Series(df_stores['Climate'].astype(str).str.lower().str.strip().values, index=df_stores['Store']).to_dict()
            tiendas_destino = df_stores['Store'].dropna().tolist()

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
            df_resultado['Max_Allocable'] = np.clip(df_resultado['LSKD_DC_SOH'] * limite_absoluto * df_resultado['Norm_Curve'], 0, df_resultado['LSKD_DC_SOH'])

            df_pesos = pd.DataFrame(index=df_resultado.index, columns=tiendas_destino)

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

            for idx in df_resultado.index:
                max_units = int(np.round(df_resultado.loc[idx, 'Max_Allocable']))
                pesos_row = df_pesos.loc[idx, tiendas_destino].values.astype(float)
                suma_pesos = np.sum(pesos_row)

                if suma_pesos == 0 or max_units <= 0:
                    df_resultado.loc[idx, tiendas_destino] = 0
                    continue

                exact_alloc = max_units * (pesos_row / suma_pesos)
                alloc = np.floor(exact_alloc).astype(int)
                remainder = int(max_units - np.sum(alloc))
                fractions = exact_alloc - alloc

                if remainder > 0:
                    indices = np.argsort(fractions)[-remainder:]
                    for i in indices:
                        alloc[i] += 1

                df_resultado.loc[idx, tiendas_destino] = alloc

        # Lanzamos una notificación emergente (Toast) en lugar de un bloque estático
        st.toast(txt["success_msg"], icon="✅")
        
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

        # Crear un buffer en memoria para el archivo Excel
        buffer = io.BytesIO()
        
        # Escribir el DataFrame en el buffer usando openpyxl
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Asignacion_Semanal')
            
        # Obtener los datos del buffer
        excel_data = buffer.getvalue()

        # Botón de descarga para el archivo .xlsx
        st.download_button(
            label=txt["download_btn"],
            data=excel_data,
            file_name='Asignacion_LSKD_Inteligente.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    except Exception as e:
        st.error(f"{txt['error_msg']}{e}")

# --- FOOTER ---
st.divider()
st.caption("© 2026 Elomejia LSKD | Elo-cations v1.0 | Newness")
