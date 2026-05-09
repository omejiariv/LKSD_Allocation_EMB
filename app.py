# app.py

import streamlit as st
import pandas as pd
import numpy as np
import gc
import plotly.express as px
import io
from datetime import datetime

st.set_page_config(page_title="LSKD Allocation Model", layout="wide")

# --- INICIALIZACIÓN DE MEMORIA (SESSION STATE) ---
if 'datos_cargados' not in st.session_state:
    st.session_state.datos_cargados = False
    st.session_state.df_newness = None
    st.session_state.df_stores = None
    st.session_state.df_curve = None
    st.session_state.df_metrics = None
    st.session_state.fecha_es = None
    st.session_state.fecha_en = None
    st.session_state.last_file_id = None
    st.session_state.editor_key = 0 # Llave maestra

# Función para limpiar la tabla cuando muevas un slider
def reset_editor():
    st.session_state.editor_key += 1

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
        "error_msg": "❌ Ocurrió un error en el cálculo: ",
        "target_woc_title": "4. Objetivo de Inventario",
        "target_woc_slider": "Semanas de Cobertura (Target WOC)",
        "tt_woc": "Define cuántas semanas de venta quieres cubrir en la tienda. Ejemplo: 4 semanas significa que la tienda siempre tendrá stock para vender un mes.",
        "risk_alert": "⚠️ ALERTA DE QUIEBRE",
        "doc_title": "📚 Documentación, Metodología e Insumos",
        "doc_content": """
        ### 📌 Resumen de la App
        Esta aplicación automatiza el proceso semanal de *Allocation* y Reposición para LSKD. Cruza el inventario de bodega (SOH) con el desempeño real de las tiendas, permitiendo un flujo de inventario inteligente y dinámico.
        
        ### ⚙️ Metodología y Conceptos Clave
        * **Híbrido Push/Pull:** Si el producto tiene historial de ventas, el sistema hace **Reposición (Pull)** basándose en el Target WOC. Si es nuevo, usa **Pesos (Push)** por grado de tienda (A, B, C, D).
        * **Target WOC (Weeks of Cover):** Calcula cuántas semanas de venta queremos cubrir. La app enviará unidades solo si el stock actual de la tienda no alcanza para cubrir el objetivo de semanas.
        * **Método del Resto Mayor:** Garantiza que el 100% de las unidades calculadas se repartan sin perderse por redondeos decimales.
        * **Filtros Inteligentes:** Aplica reglas de Clima (verano/invierno) y exclusión de tiendas C/D para productos TOP TIER.

        ### 📥 Insumos Requeridos y Estructura
        Requiere un Excel (`.xlsx`) con cuatro hojas:
        1. **`Newness`**: Base de productos a enviar.
        2. **`Store_Grading`**: Calificación y clima de tiendas.
        3. **`Size_Curve`**: Curva de tallas estática (para productos nuevos).
        4. **`Store_Metrics`**: Ventas L4W e inventario actual por tienda (para Reposición).
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
        "error_msg": "❌ An error occurred during calculation: ",
        "target_woc_title": "4. Inventory Target",
        "target_woc_slider": "Target Weeks of Cover (WOC)",
        "tt_woc": "Defines how many weeks of sales you want to cover in-store. Example: 4 weeks means the store will always have stock for one month of sales.",
        "risk_alert": "⚠️ STOCKOUT RISK",
        
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
max_send_pct = st.sidebar.slider(txt["limit_bodega_slider"], min_value=10.0, max_value=40.0, value=30.0, step=1.0, on_change=reset_editor)
flex_margin = st.sidebar.slider(txt["flex_margin"], min_value=0.0, max_value=5.0, value=3.0, step=0.5, on_change=reset_editor)

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

# --- NUEVO CONTROL: REPOSICIÓN (PULL) ---
st.sidebar.subheader(txt["target_woc_title"])
target_woc = st.sidebar.slider(txt["target_woc_slider"], 1, 8, 4, help=txt["tt_woc"], on_change=reset_editor)


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
                    # 1. Leer hojas base
                    df_newness = pd.read_excel(uploaded_file, sheet_name='Newness', engine='openpyxl')
                    df_newness.columns = df_newness.columns.astype(str).str.strip().str.replace(' ', '_')
                    
                    df_stores = pd.read_excel(uploaded_file, sheet_name='Store_Grading', engine='openpyxl')
                    df_stores.columns = df_stores.columns.astype(str).str.strip()
                    
                    df_curve = pd.read_excel(uploaded_file, sheet_name='Size_Curve', engine='openpyxl')
                    df_curve.columns = df_curve.columns.astype(str).str.strip()

                    # 2. Leer Hoja SOH (Actualiza el inventario en tiempo real)
                    try:
                        df_soh_sheet = pd.read_excel(uploaded_file, sheet_name='SOH', engine='openpyxl')
                        df_soh_sheet.columns = df_soh_sheet.columns.astype(str).str.strip()
                        soh_map = pd.Series(df_soh_sheet['SOH_Quantity'].values, index=df_soh_sheet['SKU'].astype(str).str.upper().str.strip()).to_dict()
                        df_newness['LSKD_DC_SOH'] = df_newness['SKU'].astype(str).str.upper().str.strip().map(soh_map).fillna(df_newness['LSKD_DC_SOH'])
                    except:
                        pass # Si no existe la hoja SOH, no hace nada y usa el dato de Newness

                    # 3. Leer Hoja Store_Metrics
                    try:
                        df_metrics = pd.read_excel(uploaded_file, sheet_name='Store_Metrics', engine='openpyxl')
                        df_metrics.columns = df_metrics.columns.astype(str).str.strip()
                        st.session_state.df_metrics = df_metrics
                    except:
                        st.session_state.df_metrics = None

                    # Guardar en memoria
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

        # Recuperamos las métricas con limpieza extrema
        if st.session_state.df_metrics is not None:
            df_metrics = st.session_state.df_metrics.copy()
            # Forzamos Mayúsculas y quitamos espacios para asegurar el cruce
            df_metrics['Key'] = df_metrics['SKU'].astype(str).str.upper().str.strip() + "_" + df_metrics['Store'].astype(str).str.upper().str.strip()
            metrics_dict = df_metrics.set_index('Key').to_dict('index')
        else:
            metrics_dict = {}

        # Mapeos base
        store_grades = pd.Series(df_stores['Womens_Allocation_Grade'].values, index=df_stores['Store']).to_dict()
        store_climates = pd.Series(df_stores['Climate'].astype(str).str.lower().str.strip().values, index=df_stores['Store']).to_dict()
        tiendas_destino = df_stores['Store'].dropna().tolist()

        df_resultado = df_newness[['SKU', 'Product_Name', 'Size', 'Gender', 'Gender_&_Category', 'LSKD_DC_SOH', 'Grade']].copy()
        df_resultado['LSKD_DC_SOH'] = pd.to_numeric(df_resultado['LSKD_DC_SOH'], errors='coerce').fillna(0)

        # 1. Curva Estática (Para Newness)
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

        # 2. Matriz de Pesos Inteligente y Asignación Máxima Dinámica
        df_pesos = pd.DataFrame(index=df_resultado.index, columns=tiendas_destino)
        allocable_real = []
        limite_absoluto = (max_send_pct + flex_margin) / 100.0

        for idx, row in df_resultado.iterrows():
            sku_limpio = str(row['SKU']).upper().strip()
            soh_bodega = row['LSKD_DC_SOH']
            
            suma_necesidades = 0
            pesos_fila = []
            es_reposicion = False

            for tienda in tiendas_destino:
                t_limpia = str(tienda).upper().strip()
                key = f"{sku_limpio}_{t_limpia}"
                metricas = metrics_dict.get(key, {})
                
                ventas = metricas.get('Sales_L4W', 0)
                soh_tienda = metricas.get('Store_SOH', 0)
                
                if ventas > 0:
                    es_reposicion = True
                    # Venta Semanal * Target WOC - Inventario en tienda
                    target_stock = (ventas / 4) * target_woc
                    necesidad = target_stock - soh_tienda
                    peso_final = max(necesidad, 0.001)
                    if necesidad > 0: suma_necesidades += necesidad
                else:
                    peso_final = dict_pesos.get(store_grades.get(tienda, 'C'), 0)

                pesos_fila.append(peso_final)
            
            df_pesos.loc[idx] = pesos_fila

            if es_reposicion:
                # Si es reposición, la "sed" de la tienda manda, pero no más de lo que hay en bodega
                max_unidades = min(suma_necesidades, soh_bodega)
            else:
                # Si es nuevo, manda la regla del 33%
                max_unidades = soh_bodega * ((max_send_pct + flex_margin) / 100) * row['Norm_Curve']
            
            allocable_real.append(max(0, min(max_unidades, soh_bodega)))
    
        df_resultado['Max_Allocable'] = allocable_real

        # 4. Reparto con Método del Resto Mayor
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
        
        fecha_mostrar = st.session_state.fecha_es if idioma == "Español" else st.session_state.fecha_en
        st.subheader(f"{txt['metrics_title']} ({fecha_mostrar})")

        cont_metricas = st.container()
        cont_alertas = st.container()
        cont_graficos = st.container()
        
        st.markdown("---")
        st.subheader(txt["matrix_title"])
        
        if idioma == "Español":
            st.info("💡 **Ajuste Fino:** Haz doble clic en cualquier celda debajo del nombre de una tienda para modificar la cantidad manualmente. Los gráficos y métricas se actualizarán al instante.")
        else:
            st.info("💡 **Fine-tuning:** Double-click any cell under a store name to manually modify the quantity. Charts and metrics will update instantly.")

        columnas_protegidas = ['SKU', 'Product_Name', 'Size', 'Gender', 'Gender_&_Category', 'LSKD_DC_SOH', 'Grade', 'Curve_Multiplier', 'Norm_Curve', 'Max_Allocable']
        
        if 'editor_key' not in st.session_state:
            st.session_state.editor_key = 0

        df_editado = st.data_editor(
            df_resultado,
            disabled=columnas_protegidas,
            use_container_width=True,
            hide_index=True,
            key=f"editor_matriz_{st.session_state.editor_key}"
        )

        with cont_metricas:
            total_inventario = df_editado['LSKD_DC_SOH'].sum()
            df_solo_tiendas = df_editado[tiendas_destino]
            total_asignado = df_solo_tiendas.sum().sum()
            
            col1, col2, col3 = st.columns(3)
            col1.metric(txt["metric_inv"], f"{int(total_inventario):,}")
            col2.metric(txt["metric_dist"], f"{int(total_asignado):,}")
            pct_global = (total_asignado / total_inventario * 100) if total_inventario > 0 else 0
            col3.metric(txt["metric_pct"], f"{pct_global:.1f}%")

        with cont_alertas:
            # CORRECCIÓN DE LA ALERTA: Aseguramos que ambos son porcentajes flotantes reales
            limite_porcentual = float(max_send_pct + flex_margin)
            asignacion_porcentual = float(pct_global)
            
            if round(asignacion_porcentual, 1) > round(limite_porcentual, 1):
                if idioma == "Español":
                    st.warning(f"⚠️ **Límite Excedido:** Tus ajustes manuales han elevado la asignación global al **{asignacion_porcentual:.1f}%**, superando la regla máxima permitida (**{limite_porcentual:.1f}%**).")
                    btn_text = "🔄 Restaurar Regla Matemática"
                else:
                    st.warning(f"⚠️ **Limit Exceeded:** Current allocation (**{asignacion_porcentual:.1f}%**) exceeds the limit (**{limite_porcentual:.1f}%**).")
                    btn_text = "🔄 Reset to Math Rule"
                
                if st.button(btn_text):
                    st.session_state.editor_key += 1
                    st.rerun()

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
                
                # NUEVO SELECTOR DE ORDEN
                opciones_orden = ["Mayor a Menor", "Menor a Mayor", "Orden Natural (Tallas)"]
                orden = st.radio("Ordenar barras por:", opciones_orden, horizontal=True, key="orden_tallas_radio")
                
                if orden == "Mayor a Menor":
                    df_tallas_filt = df_tallas_filt.sort_values(by='Total_Asignado', ascending=False)
                elif orden == "Menor a Mayor":
                    df_tallas_filt = df_tallas_filt.sort_values(by='Total_Asignado', ascending=True)
                else:
                    orden_tallas_retail = ['2XS', 'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', 'ONE SIZE']
                    df_tallas_filt['Size'] = pd.Categorical(df_tallas_filt['Size'], categories=orden_tallas_retail, ordered=True)
                    df_tallas_filt = df_tallas_filt.sort_values('Size')
                
                fig_size = px.bar(
                    df_tallas_filt, x='Size', y='Total_Asignado',
                    color='Total_Asignado', color_continuous_scale='Teal', text_auto=True
                )
                # FORZAR A PLOTLY A OBEDECER TU ORDEN
                fig_size.update_xaxes(categoryorder='array', categoryarray=df_tallas_filt['Size'].astype(str))
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

        # --- REGISTRO HISTÓRICO EN EXCEL ---
        st.markdown("---")
        st.subheader("💾 Registro Histórico Acumulado")
        
        col_h1, col_h2 = st.columns(2)
        nombre_csv = "historico_asignaciones_lskd.csv"

        with col_h1:
            if st.button("➕ Integrar esta semana al Historial"):
                df_hist = df_editado.copy()
                df_hist['Fecha_Registro'] = st.session_state.fecha_es
                try:
                    df_ex = pd.read_csv(nombre_csv)
                    pd.concat([df_ex, df_hist], ignore_index=True).to_csv(nombre_csv, index=False)
                except:
                    df_hist.to_csv(nombre_csv, index=False)
                st.success("✅ Datos guardados con éxito.")

        with col_h2:
            import os
            if os.path.exists(nombre_csv):
                # Leemos el CSV oculto y lo convertimos a Excel en memoria para la descarga
                df_full = pd.read_csv(nombre_csv)
                buffer_xlsx = io.BytesIO()
                with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                    df_full.to_excel(writer, index=False, sheet_name='Historial')
                
                st.download_button(
                    label="📥 Descargar Historial Completo (Excel)",
                    data=buffer_xlsx.getvalue(),
                    file_name="Master_Historico_LSKD.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_desc_hist_xlsx"
                )

    except Exception as e:
        st.error(f"{txt['error_msg']}{e}")

# --- FOOTER ---
st.divider()
st.caption("© 2026 EloMejiaB LSKD | Elo-cations v1.0 | Newness & Replenishment")
