import streamlit as st
import pandas as pd
import numpy as np
import gc
import plotly.express as px
import io
from datetime import datetime

st.set_page_config(page_title="LSKD Allocation Model", layout="wide")

# --- INICIALIZACIÓN DE MEMORIA (SESSION STATE) ---
# Esto garantiza que los datos no desaparezcan de la pantalla al hacer cambios
if 'datos_cargados' not in st.session_state:
    st.session_state.datos_cargados = False
    st.session_state.df_newness = None
    st.session_state.df_stores = None
    st.session_state.df_curve = None
    st.session_state.fecha_es = None
    st.session_state.fecha_en = None
    st.session_state.last_file_id = None

# --- FUNCIONES DE FECHA ---
def obtener_fecha_es():
    now = datetime.now()
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    return f"{now.day:02d} de {meses[now.month - 1]} de {now.year}"

def obtener_fecha_en():
    now = datetime.now()
    return now.strftime("%B %d, %Y")

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
        "admin_title": "🔒 Acceso Administrador",
        "admin_pwd": "Clave de acceso",
        "upload_file": "Sube el archivo LSKD_Newness_EMB.xlsx",
        "processing": "⏱️ Procesando datos en memoria...",
        "success_msg": "Datos cargados y procesados con éxito.",
        "metrics_title": "📊 Métricas Globales de la Semana",
        "metric_inv": "📦 Inventario Total (DC SOH)",
        "metric_dist": "🚚 Unidades a Distribuir",
        "metric_pct": "🎯 % de Asignación Global",
        "chart_title_store": "##### Unidades Asignadas por Tienda",
        "chart_title_size": "##### 📏 Unidades Asignadas por Talla",
        "size_filter": "Selecciona las tallas a visualizar:",
        "matrix_title": "📋 Matriz de Asignación Final",
        "download_btn": "📥 Descargar Matriz (Excel)",
        "error_msg": "❌ Ocurrió un error en el cálculo: ",
        "error_msg": "❌ Ocurrió un error en el cálculo: ", # <--- Asegúrate de que esta coma esté aquí
        
        "doc_title": "📚 Documentación, Metodología e Insumos",
        "doc_content": """
        ### 📌 Resumen de la App
        Esta aplicación automatiza el proceso semanal de *Allocation* (Asignación de Inventario) para LSKD. Cruza el inventario disponible en bodega (SOH) con la calificación de cada tienda, respetando reglas estrictas de capacidad y adaptándose orgánicamente a la curva de tallas histórica.
        
        ### ⚙️ Metodología y Conceptos Clave
        * **Método del Resto Mayor (Largest Remainder):** Algoritmo utilizado en el paso final del reparto. Evita la "pérdida por redondeo hacia abajo", asegurando que las fracciones decimales sobrantes se sumen y se asignen como unidades enteras a las tiendas más cercanas al siguiente decimal.
        * **Multiplicador de Curvas Dinámico:** Cruza la talla y categoría de tu producto con la matriz de `Size_Curve`. Si una talla popular tiene un peso alto, automáticamente infla su asignación antes de redondear.
        * **Selector de Temporada:** Si el producto dice `CLIMATE SPECIFIC`, el motor revisará el clima de cada tienda y dejará en cero (0) a las ciudades que no encajen con la temporada seleccionada.
        * **Filtro TOP TIER:** Si un producto tiene grado `TOP TIER`, se excluye de las tiendas C y D.

        ### 📥 Insumos Requeridos y Estructura
        Requiere un Excel (`.xlsx`) con tres hojas:
        1. **`Newness`**: Base de datos de productos.
        2. **`Store_Grading`**: Base de datos de tiendas con calificación y clima.
        3. **`Size_Curve`**: Matriz de multiplicadores de demanda.
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
        "admin_title": "🔒 Administrator Access",
        "admin_pwd": "Access password",
        "upload_file": "Upload LSKD_Newness_EMB.xlsx file",
        "processing": "⏱️ Processing data into memory...",
        "success_msg": "Data loaded and processed successfully.",
        "metrics_title": "📊 Weekly Global Metrics",
        "metric_inv": "📦 Total Inventory (DC SOH)",
        "metric_dist": "🚚 Units to Allocate",
        "metric_pct": "🎯 Global Allocation %",
        "chart_title_store": "##### Units Allocated per Store",
        "chart_title_size": "##### 📏 Units Allocated by Size",
        "size_filter": "Select sizes to display:",
        "matrix_title": "📋 Final Allocation Matrix",
        "download_btn": "📥 Download Matrix (Excel)",
        "error_msg": "❌ An error occurred during calculation: ",
        "error_msg": "❌ An error occurred during calculation: ", # <--- Asegúrate de que esta coma esté aquí
        
        # --- AÑADIR DESDE AQUÍ ---
        "doc_title": "📚 Documentation, Methodology & Inputs",
        "doc_content": """
        ### 📌 App Summary
        This application automates the weekly *Allocation* process for LSKD. It crosses the available warehouse inventory (SOH) with each store's grading, respecting strict capacity rules and adapting organically to the historical size curve.
        
        ### ⚙️ Methodology & Key Concepts
        * **Largest Remainder Method:** Algorithm used in the final distribution step. It avoids "rounding down loss", ensuring exact global targets.
        * **Dynamic Curve Multiplier:** Crosses the product's size and category with the `Size_Curve` sheet to naturally inflate allocation for popular sizes based on historical weights.
        * **Season Selector:** If the weekly product is graded `CLIMATE SPECIFIC`, the engine zeroes out the allocation to cities that do not match the selected season.
        * **TOP TIER Filter:** Excludes `TOP TIER` products from C and D graded stores automatically.

        ### 📥 Required Inputs & Structure
        Requires an Excel (`.xlsx`) file with three sheets:
        1. **`Newness`**: Product database.
        2. **`Store_Grading`**: Store database with grading and climate.
        3. **`Size_Curve`**: Matrix of demand multipliers.
        """
    }
}

txt = t[idioma]
st.title(txt["title"])
# --- DOCUMENTACIÓN DESPLEGABLE (SIEMPRE VISIBLE) ---
with st.expander(txt["doc_title"]):
    st.markdown(txt["doc_content"])

# --- PANEL LATERAL (CONTROLES) ---
st.sidebar.header(txt["sidebar_header"])

st.sidebar.subheader(txt["limit_bodega_title"])
max_send_pct = st.sidebar.slider(txt["limit_bodega_slider"], min_value=10.0, max_value=40.0, value=30.0, step=1.0)
flex_margin = st.sidebar.slider(txt["flex_margin"], min_value=0.0, max_value=5.0, value=3.0, step=0.5)

st.sidebar.subheader(txt["season_filter_title"])
st.sidebar.markdown(txt["season_filter_desc"])
opciones_clima_ui = [txt["season_summer"], txt["season_winter"], txt["season_both"]]
seleccion_clima_ui = st.sidebar.radio(txt["season_target"], opciones_clima_ui)

temporada_backend = "verano" if seleccion_clima_ui == txt["season_summer"] else ("invierno" if seleccion_clima_ui == txt["season_winter"] else "ambos")

st.sidebar.subheader(txt["weights_title"])
peso_a = st.sidebar.number_input(txt["weight_a"], value=40)
peso_b = st.sidebar.number_input(txt["weight_b"], value=30)
peso_c = st.sidebar.number_input(txt["weight_c"], value=20)
peso_d = st.sidebar.number_input(txt["weight_d"], value=10)

total_peso = max(peso_a + peso_b + peso_c + peso_d, 1)
dict_pesos = {'A': peso_a / total_peso, 'B': peso_b / total_peso, 'C': peso_c / total_peso, 'D': peso_d / total_peso}


# --- ACCESO ADMINISTRADOR Y CARGA DE ARCHIVO ---
with st.expander(txt["admin_title"], expanded=not st.session_state.datos_cargados):
    clave = st.text_input(txt["admin_pwd"], type="password")
    
    # CLAVE DE ACCESO CONFIGURABLE AQUÍ
    if clave == "LSKD2026":
        uploaded_file = st.file_uploader(txt["upload_file"], type=["xlsx"], key="cargador_excel")
        
        if uploaded_file:
            current_file_id = uploaded_file.name + str(uploaded_file.size)
            
            # Solo procesamos si es un archivo nuevo, sino usamos la memoria
            if st.session_state.last_file_id != current_file_id:
                with st.spinner(txt["processing"]):
                    df_newness = pd.read_excel(uploaded_file, sheet_name='Newness', engine='openpyxl')
                    df_newness = df_newness.dropna(how='all', axis=0)
                    df_newness.columns = df_newness.columns.astype(str).str.strip().str.replace(' ', '_')
                    
                    df_stores = pd.read_excel(uploaded_file, sheet_name='Store_Grading', engine='openpyxl')
                    df_stores = df_stores.dropna(how='all', axis=0)
                    df_stores.columns = df_stores.columns.astype(str).str.strip()
                    
                    df_curve = pd.read_excel(uploaded_file, sheet_name='Size_Curve', engine='openpyxl')
                    df_curve.columns = df_curve.columns.astype(str).str.strip()

                    # Guardar en memoria RAM de Streamlit
                    st.session_state.df_newness = df_newness
                    st.session_state.df_stores = df_stores
                    st.session_state.df_curve = df_curve
                    st.session_state.fecha_es = obtener_fecha_es()
                    st.session_state.fecha_en = obtener_fecha_en()
                    st.session_state.last_file_id = current_file_id
                    st.session_state.datos_cargados = True
                    
                    st.toast(txt["success_msg"], icon="✅")


# --- MOTOR MATEMÁTICO Y DASHBOARD (EJECUTADO DESDE LA MEMORIA) ---
if st.session_state.datos_cargados:
    try:
        # Recuperamos los datos limpios de la memoria
        df_newness = st.session_state.df_newness.copy()
        df_stores = st.session_state.df_stores.copy()
        df_curve = st.session_state.df_curve.copy()

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
            except: pass
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

        # --- SECCIÓN VISUAL ---
        st.markdown("---")
        
        # 1. Título con Fecha Dinámica
        fecha_mostrar = st.session_state.fecha_es if idioma == "Español" else st.session_state.fecha_en
        st.subheader(f"{txt['metrics_title']} ({fecha_mostrar})")

        # 2. Creamos los contenedores vacíos (ahora incluimos uno para alertas)
        cont_metricas = st.container()
        cont_alertas = st.container()  # <--- NUEVO
        cont_graficos = st.container()
        
        st.markdown("---")
        st.subheader(txt["matrix_title"])
        
        if idioma == "Español":
            st.info("💡 **Ajuste Fino:** Haz doble clic en cualquier celda debajo del nombre de una tienda para modificar la cantidad manualmente. Los gráficos y métricas se actualizarán al instante.")
        else:
            st.info("💡 **Fine-tuning:** Double-click any cell under a store name to manually modify the quantity. Charts and metrics will update instantly.")

        columnas_protegidas = ['SKU', 'Product_Name', 'Size', 'Gender', 'Gender_&_Category', 'LSKD_DC_SOH', 'Grade', 'Curve_Multiplier', 'Norm_Curve', 'Max_Allocable']
        
        # 3. Inicializamos una "llave" para el editor en la memoria
        if 'editor_key' not in st.session_state:
            st.session_state.editor_key = 0

        # 4. El editor ahora usa la llave dinámica
        df_editado = st.data_editor(
            df_resultado,
            disabled=columnas_protegidas,
            use_container_width=True,
            hide_index=True,
            key=f"editor_matriz_{st.session_state.editor_key}"
        )

        # 5. Calculamos e inyectamos las Métricas
        with cont_metricas:
            total_inventario = df_editado['LSKD_DC_SOH'].sum()
            df_solo_tiendas = df_editado[tiendas_destino]
            total_asignado = df_solo_tiendas.sum().sum()
            
            col1, col2, col3 = st.columns(3)
            col1.metric(txt["metric_inv"], f"{int(total_inventario):,}")
            col2.metric(txt["metric_dist"], f"{int(total_asignado):,}")
            pct_global = (total_asignado / total_inventario * 100) if total_inventario > 0 else 0
            col3.metric(txt["metric_pct"], f"{pct_global:.1f}%")

        # 6. LÓGICA DE ALERTA Y RECALCULO AUTOMÁTICO
        with cont_alertas:
            limite_maximo_permitido = max_send_pct + flex_margin
            
            # Si el % sobrepasa la regla configurada en el panel izquierdo...
            if pct_global > limite_maximo_permitido:
                if idioma == "Español":
                    st.warning(f"⚠️ **Límite Excedido:** Tus ajustes manuales han elevado la asignación global al **{pct_global:.1f}%**, superando la regla máxima permitida (**{limite_maximo_permitido}%**).")
                    btn_text = "🔄 Restaurar Regla Matemática (Borrar Cambios Manuales)"
                else:
                    st.warning(f"⚠️ **Limit Exceeded:** Your manual edits raised the global allocation to **{pct_global:.1f}%**, exceeding the maximum allowed rule (**{limite_maximo_permitido}%**).")
                    btn_text = "🔄 Reset to Mathematical Rule (Clear Manual Edits)"
                
                # Si el usuario hace clic en el botón, cambiamos la llave y recargamos la página
                if st.button(btn_text):
                    st.session_state.editor_key += 1
                    st.rerun()

        # 7. Inyectamos los Gráficos
        with cont_graficos:
            graf_col1, graf_col2 = st.columns(2)

            with graf_col1:
                st.markdown(txt["chart_title_store"])
                asignacion_por_tienda = df_solo_tiendas.sum().reset_index()
                asignacion_por_tienda.columns = ['Tienda', 'Unidades']
                asignacion_por_tienda = asignacion_por_tienda.sort_values(by='Unidades', ascending=False)
                
                fig_store = px.bar(
                    asignacion_por_tienda, x='Tienda', y='Unidades',
                    color='Unidades', color_continuous_scale='Blues', text_auto=True
                )
                fig_store.update_layout(xaxis_tickangle=-45, showlegend=False, margin=dict(t=10, b=10))
                st.plotly_chart(fig_store, use_container_width=True)

            with graf_col2:
                st.markdown(txt["chart_title_size"])
                df_editado_tallas = df_editado.copy()
                df_editado_tallas['Total_Asignado'] = df_solo_tiendas.sum(axis=1)
                df_tallas = df_editado_tallas.groupby('Size')['Total_Asignado'].sum().reset_index()
                
                tallas_disponibles = df_tallas['Size'].unique().tolist()
                tallas_seleccionadas = st.multiselect(txt["size_filter"], tallas_disponibles, default=tallas_disponibles)
                
                df_tallas_filt = df_tallas[df_tallas['Size'].isin(tallas_seleccionadas)]
                
                fig_size = px.bar(
                    df_tallas_filt, x='Size', y='Total_Asignado',
                    color='Total_Asignado', color_continuous_scale='Teal', text_auto=True
                )
                fig_size.update_layout(margin=dict(t=10, b=10))
                st.plotly_chart(fig_size, use_container_width=True)

        # 8. Descarga de Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_editado.to_excel(writer, index=False, sheet_name='Asignacion_Semanal')
        excel_data = buffer.getvalue()

        st.download_button(
            label=txt["download_btn"],
            data=excel_data,
            file_name='Asignacion_LSKD_Ajustada.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    except Exception as e:
        st.error(f"{txt['error_msg']}{e}")
        
# --- FOOTER ---
st.divider()
st.caption("© 2026 Elomejia LSKD | Elo-cations v1.0 | Newness")
