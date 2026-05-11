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

if 'escenario_a' not in st.session_state:
    st.session_state.escenario_a = None
if 'escenario_b' not in st.session_state:
    st.session_state.escenario_b = None

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
        "target_woc_title": "4. Objetivo de Inventario",
        "target_woc_slider": "Semanas de Cobertura (Target WOC)",
        "tt_woc": "Define cuántas semanas de venta quieres cubrir en la tienda. Ejemplo: 4 semanas significa que la tienda siempre tendrá stock para vender un mes.",
        "risk_alert": "⚠️ ALERTA DE QUIEBRE",
        "doc_title": "📚 Documentación, Metodología e Insumos",
        "doc_content": """
        ### 📌 Resumen de la App
        Esta aplicación automatiza el proceso semanal de *Allocation* y Reposición para LSKD. Es un sistema modular que se adapta a la información disponible: puede hacer un reparto básico (Push) o cruzar el inventario en tiempo real (SOH) con el desempeño real de las tiendas (Pull), permitiendo un flujo de inventario inteligente y dinámico.
        
        ### ⚙️ Metodología y Conceptos Clave
        * **Híbrido Push/Pull:** Si el producto tiene historial de ventas, el sistema activa la **Reposición (Pull)** basándose en el Target WOC. Para lanzamientos nuevos, utiliza **Pesos (Push)** por grado de tienda (A, B, C, D).
        * **Target WOC (Weeks of Cover):** Define el objetivo de cobertura en semanas. La app calcula la "sed" de inventario de cada tienda restando su stock actual del stock ideal necesario para cubrir el periodo seleccionado.
        * **Método del Resto Mayor:** Algoritmo de precisión que elimina el error de redondeo decimal, asegurando que el 100% de las unidades disponibles se asignen equitativamente.
        * **Filtros Inteligentes:** Gestión automática de exclusiones para productos **TOP TIER** (solo tiendas A y B) y segmentación por **Clima** (Verano/Invierno/Ambos).
        * **Protección de Sobre-stock:** El motor bloquea envíos a tiendas cuyo inventario actual ya supera el objetivo de cobertura, priorizando el flujo hacia puntos de venta con riesgo de quiebre.

        ### 🧠 Inteligencia de Curvas y Entrenamiento
        El sistema cuenta con un motor de **Machine Learning** incipiente que lee el archivo de registro histórico acumulado. 
        * **Semanas de Entrenamiento:** Indica cuántos periodos de datos históricos tiene la app en su memoria (`historico_asignaciones_lskd.csv`). 
        * **Selección Dinámica:** Permite elegir qué semanas específicas usar para entrenar el modelo, permitiendo ignorar periodos atípicos (como promociones agresivas o cierres temporales) para obtener una curva de demanda más limpia.

        ### 📥 Estructura de Insumos (Las 5 Hojas del Master Excel)
        El sistema escanea el archivo cargado y activa niveles de inteligencia según las pestañas presentes:
        1. **`Newness`**: Base maestra con la lista de productos, categorías, tallas y el grado del producto.
        2. **`Store_Grading`**: Directorio de tiendas con su calificación (A-D) y zona climática asignada.
        3. **`Size_Curve`**: Matriz de multiplicadores estáticos por categoría (usada cuando no hay datos de venta).
        4. **`Store_Metrics`**: El pulso de las tiendas. Contiene las ventas de las últimas 4 semanas (`Sales_L4W`) y el stock físico en tienda (`Store_SOH`).
        5. **`SOH`**: Reporte vivo del Centro de Distribución. Utiliza las columnas `SKU` y `LSKD_DC` para actualizar las existencias en bodega justo antes de iniciar el cálculo, garantizando que no se asigne mercancía inexistente.
        """,
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
        "target_woc_title": "4. Inventory Target",
        "target_woc_slider": "Target Weeks of Cover (WOC)",
        "tt_woc": "Defines how many weeks of sales you want to cover in-store. Example: 4 weeks means the store will always have stock for one month of sales.",
        "risk_alert": "⚠️ STOCKOUT RISK",
        "doc_title": "📚 Documentation, Methodology & Inputs",
        "doc_content": """
        ### 📌 App Summary
        This application automates the weekly *Allocation* and Replenishment process for LSKD. It is a modular system that adapts to the available information: it can execute a basic distribution (Push) or cross real-time inventory (SOH) with actual store performance (Pull), ensuring a smart and dynamic inventory flow.
        
        ### ⚙️ Methodology & Key Concepts
        * **Hybrid Push/Pull:** If the product has a sales history, the system activates **Replenishment (Pull)** based on the Target WOC. For new releases, it uses **Weights (Push)** by store grade (A, B, C, D).
        * **Target WOC (Weeks of Cover):** Defines the coverage goal in weeks. The app calculates each store's inventory "thirst" by subtracting its current stock from the ideal stock needed to cover the selected period.
        * **Largest Remainder Method:** A precision algorithm that eliminates decimal rounding errors, ensuring 100% of the available units are allocated accurately.
        * **Smart Filters:** Automatic exclusion management for **TOP TIER** products (only A and B stores) and **Climate** segmentation (Summer/Winter/Both).
        * **Over-stock Protection:** The engine blocks shipments to stores whose current inventory already exceeds the coverage target, prioritizing flow to locations at risk of out-of-stocks.

        ### 🧠 Curve Intelligence & Training
        The system features an emerging **Machine Learning** engine that reads the accumulated historical log. 
        * **Training Weeks:** Indicates how many periods of historical data the app has in its memory (`historico_asignaciones_lskd.csv`). 
        * **Dynamic Selection:** Allows choosing which specific weeks to use for training the model. This makes it possible to ignore atypical periods (like aggressive promotions or temporary closures) to obtain a cleaner demand curve.

        ### 📥 Inputs Structure (The 5 Sheets of the Master Excel)
        The system scans the uploaded file and activates intelligence levels based on the present sheets:
        1. **`Newness`**: Master base with the product list, categories, sizes, and product grade.
        2. **`Store_Grading`**: Store directory with their grading (A-D) and assigned climate zone.
        3. **`Size_Curve`**: Matrix of static multipliers by category (used when there is no sales data).
        4. **`Store_Metrics`**: The pulse of the stores. Contains the sales of the last 4 weeks (`Sales_L4W`) and physical stock in store (`Store_SOH`).
        5. **`SOH`**: Live report from the Distribution Center. Uses the `SKU` and `LSKD_DC` columns to update warehouse stock right before calculating, ensuring non-existent merchandise is never allocated.
        """
    }
}

txt = t[idioma]

# --- ENCABEZADO CON LOGO ---
# Creamos dos columnas: una grande para el título (85% del espacio) y una pequeña para el logo (15%)
col_titulo, col_logo = st.columns([0.85, 0.15])

with col_titulo:
    st.title(txt["title"])

with col_logo:
    # Usamos un ancho fijo (ej. 80 o 100 píxeles) para que iguale la altura del texto del título
    try:
        st.image("Logo LSKD.png", width=90) 
    except FileNotFoundError:
        pass # Si por alguna razón el archivo no está, la app no se caerá
        
# --- DOCUMENTACIÓN DESPLEGABLE (SIEMPRE VISIBLE) ---
with st.expander(txt["doc_title"]):
    st.markdown(txt["doc_content"])

# --- PANEL LATERAL (CONTROLES) ---
st.sidebar.header(txt["sidebar_header"])

# 1. Expandible de Límite de Bodega (Abierto por defecto)
with st.sidebar.expander(txt["limit_bodega_title"], expanded=True):
    max_send_pct = st.slider(txt["limit_bodega_slider"], min_value=10.0, max_value=40.0, value=30.0, step=1.0, on_change=reset_editor)
    flex_margin = st.slider(txt["flex_margin"], min_value=0.0, max_value=5.0, value=3.0, step=0.5, on_change=reset_editor)

# 2. Expandible de Filtro de Temporada (Cerrado por defecto para ahorrar espacio)
with st.sidebar.expander(txt["season_filter_title"], expanded=False):
    st.write(txt["season_filter_desc"])
    temporada_backend = st.radio(txt["season_target"], [txt["season_summer"], txt["season_winter"], txt["season_both"]], index=2, on_change=reset_editor).lower()
    
    # Normalización del clima para el motor
    if temporada_backend == txt["season_summer"].lower():
        temporada_backend = "verano"
    elif temporada_backend == txt["season_winter"].lower():
        temporada_backend = "invierno"
    else:
        temporada_backend = "ambos"

# 3. Expandible de Pesos por Categoría (Cerrado por defecto)
with st.sidebar.expander(txt["weights_title"], expanded=False):
    w_a = st.number_input(txt["weight_a"], min_value=0.0, value=0.40, step=0.05, on_change=reset_editor)
    w_b = st.number_input(txt["weight_b"], min_value=0.0, value=0.30, step=0.05, on_change=reset_editor)
    w_c = st.number_input(txt["weight_c"], min_value=0.0, value=0.20, step=0.05, on_change=reset_editor)
    w_d = st.number_input(txt["weight_d"], min_value=0.0, value=0.10, step=0.05, on_change=reset_editor)
    dict_pesos = {'A': w_a, 'B': w_b, 'C': w_c, 'D': w_d}

# 4. Target WOC (Lo dejamos afuera porque es el control principal de Reposición)
st.sidebar.markdown("---")
st.sidebar.subheader(txt["target_woc_title"])
target_woc = st.sidebar.slider(txt["target_woc_slider"], min_value=1, max_value=8, value=4, help=txt["tt_woc"], on_change=reset_editor)

# --- NUEVO CONTROL: ENTRENAMIENTO DE CURVAS ---
st.sidebar.markdown("---")
st.sidebar.subheader("🧠 Entrenamiento de Curvas (Machine Learning)")
try:
    import os
    if os.path.exists("historico_asignaciones_lskd.csv"):
        df_historial = pd.read_csv("historico_asignaciones_lskd.csv")
        semanas_disponibles = df_historial['Fecha_Registro'].dropna().unique().tolist()
        num_semanas = len(semanas_disponibles)
        
        st.sidebar.success(f"✅ ¡Tenemos **{num_semanas} semanas** de datos de entrenamiento disponibles!")
        
        semanas_seleccionadas = st.sidebar.multiselect(
            "Selecciona qué semanas usar para el cálculo de curvas dinámicas:",
            options=semanas_disponibles,
            default=semanas_disponibles
        )
    else:
        st.sidebar.info("⏳ Aún no hay semanas de entrenamiento. (Se usarán las curvas estáticas de Size_Curve). Guarda tu primera asignación para empezar a entrenar el modelo.")
        semanas_seleccionadas = []
except Exception as e:
    st.sidebar.info("⏳ Aún no hay semanas de entrenamiento. Guarda tu primera asignación para empezar a entrenar.")
    semanas_seleccionadas = []

# --- ACCESO ADMINISTRADOR Y CARGA DE ARCHIVO ---
with st.expander(txt["admin_title"], expanded=not st.session_state.datos_cargados):
    clave = st.text_input(txt["admin_pwd"], type="password")
    
    # CLAVE DE ACCESO CONFIGURABLE AQUÍ
    if clave == "LSKD2026":
        uploaded_file = st.file_uploader(txt["upload_file"], type=["xlsx"], key="cargador_excel")
        
        if uploaded_file:
            current_file_id = uploaded_file.name + str(uploaded_file.size)
            
            if st.session_state.last_file_id != current_file_id:
                with st.spinner(txt["processing"]):
                    # Escanear las hojas que trae el archivo
                    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
                    hojas_disponibles = xls.sheet_names
                    
                    # Buscador inteligente de hojas (ignora espacios, guiones y mayúsculas)
                    def buscar_hoja(nombre_ideal):
                        for h in hojas_disponibles:
                            if h.replace(' ', '').replace('_', '').lower() == nombre_ideal.replace(' ', '').replace('_', '').lower():
                                return h
                        return nombre_ideal # Si no la encuentra, devuelve el original para forzar el error manejado

                    try:
                        # 1. Leer hojas base (OBLIGATORIAS)
                        df_newness = pd.read_excel(xls, sheet_name=buscar_hoja('Newness'))
                        df_newness.columns = df_newness.columns.astype(str).str.strip().str.replace(' ', '_')
                        
                        df_stores = pd.read_excel(xls, sheet_name=buscar_hoja('Store_Grading'))
                        df_stores.columns = df_stores.columns.astype(str).str.strip()
                        
                        df_curve = pd.read_excel(xls, sheet_name=buscar_hoja('Size_Curve'))
                        df_curve.columns = df_curve.columns.astype(str).str.strip()
                    except ValueError:
                        # Si realmente falta una hoja, mostramos un error amigable y detenemos el proceso
                        st.error("❌ **Error Crítico:** Tu archivo Excel no tiene las 3 hojas base obligatorias (`Newness`, `Store_Grading` y `Size_Curve`). Por favor revisa el archivo y vuelve a subirlo.")
                        st.stop() # Esto detiene la app limpiamente sin mostrar la pantalla roja

                    modo_app = "Básico (Push)"

                    # 2. Leer Hoja SOH (OPCIONAL - NIVEL AVANZADO)
                    hoja_soh_encontrada = buscar_hoja('SOH')
                    if hoja_soh_encontrada in hojas_disponibles:
                        try:
                            df_soh_sheet = pd.read_excel(xls, sheet_name=hoja_soh_encontrada)
                            df_soh_sheet.columns = df_soh_sheet.columns.astype(str).str.strip()
                            nombre_columna_sku = 'SKU' 
                            nombre_columna_cantidad = 'LSKD_DC'
                            soh_map = pd.Series(df_soh_sheet[nombre_columna_cantidad].values, index=df_soh_sheet[nombre_columna_sku].astype(str).str.upper().str.strip()).to_dict()
                            df_newness['LSKD_DC_SOH'] = df_newness['SKU'].astype(str).str.upper().str.strip().map(soh_map).fillna(df_newness['LSKD_DC_SOH'])
                            modo_app = "Avanzado (Con SOH)"
                        except Exception as e:
                            st.warning(f"⚠️ La hoja 'SOH' existe pero hubo un error al leerla: {e}")

                    # 3. Leer Hoja Store_Metrics (OPCIONAL - NIVEL HÍBRIDO)
                    hoja_metrics_encontrada = buscar_hoja('Store_Metrics')
                    if hoja_metrics_encontrada in hojas_disponibles:
                        try:
                            df_metrics = pd.read_excel(xls, sheet_name=hoja_metrics_encontrada)
                            df_metrics.columns = df_metrics.columns.astype(str).str.strip()
                            st.session_state.df_metrics = df_metrics
                            if modo_app == "Avanzado (Con SOH)":
                                modo_app = "Enterprise (5 Hojas)"
                            else:
                                modo_app = "Híbrido (Push/Pull)"
                        except Exception as e:
                            st.warning(f"⚠️ Error en hoja 'Store_Metrics': {e}")
                            st.session_state.df_metrics = None
                    else:
                        st.session_state.df_metrics = None

                    # Guardar en memoria
                    st.session_state.df_newness = df_newness
                    st.session_state.df_stores = df_stores
                    st.session_state.df_curve = df_curve
                    st.session_state.fecha_es = obtener_fecha_es()
                    st.session_state.fecha_en = obtener_fecha_en()
                    st.session_state.last_file_id = current_file_id
                    st.session_state.datos_cargados = True
                    
                    # Mensaje de éxito informando el modo
                    st.success(f"✅ Archivo cargado exitosamente. Nivel de inteligencia activado: **{modo_app}**")

# --- MOTOR MATEMÁTICO Y DASHBOARD ---
if st.session_state.datos_cargados:
    try:
        df_newness = st.session_state.df_newness.copy()
        df_stores = st.session_state.df_stores.copy()
        df_curve = st.session_state.df_curve.copy()

        if st.session_state.df_metrics is not None:
            df_metrics = st.session_state.df_metrics.copy()
            df_metrics['Key'] = df_metrics['SKU'].astype(str).str.upper().str.strip() + "_" + df_metrics['Store'].astype(str).str.upper().str.strip()
            metrics_dict = df_metrics.set_index('Key').to_dict('index')
        else:
            metrics_dict = {}

        # MAPEOS BLINDADOS (Elimina espacios invisibles y problemas de mayúsculas en Excel)
        col_grade = [c for c in df_stores.columns if 'WOMEN' in c.upper()][0] # Detecta 'Womens Allocation Grade' automáticamente
        store_grades = pd.Series(df_stores[col_grade].astype(str).str.upper().str.strip().values, index=df_stores['Store'].astype(str).str.upper().str.strip()).to_dict()
        store_climates = pd.Series(df_stores['Climate'].astype(str).str.lower().str.strip().values, index=df_stores['Store'].astype(str).str.upper().str.strip()).to_dict()
        tiendas_destino = df_stores['Store'].dropna().tolist()

        df_resultado = df_newness[['SKU', 'Product_Name', 'Size', 'Gender', 'Gender_&_Category', 'LSKD_DC_SOH', 'Grade']].copy()
        df_resultado['LSKD_DC_SOH'] = pd.to_numeric(df_resultado['LSKD_DC_SOH'], errors='coerce').fillna(0)

        # 1. Curva Estática
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

        for idx, row in df_resultado.iterrows():
            sku_limpio = str(row['SKU']).upper().strip()
            soh_bodega = float(row['LSKD_DC_SOH'])
            norm_curve = float(row['Norm_Curve'])
            
            suma_necesidades = 0.0
            pesos_fila = []
            es_reposicion = False

            for tienda in tiendas_destino:
                t_limpia = str(tienda).upper().strip()
                
                # Extracción súper segura del grado
                grade_tienda = store_grades.get(t_limpia, 'C') 
                clima_tienda = store_climates.get(t_limpia, '')

                key = f"{sku_limpio}_{t_limpia}"
                metricas = metrics_dict.get(key, {})
                
                try: ventas = float(metricas.get('Sales_L4W', 0))
                except: ventas = 0.0
                    
                try: soh_tienda = float(metricas.get('Store_SOH', 0))
                except: soh_tienda = 0.0
                
                if ventas > 0:
                    es_reposicion = True
                    target_stock = (ventas / 4.0) * float(target_woc)
                    necesidad = target_stock - soh_tienda
                    peso_final = max(necesidad, 0.001)
                    if necesidad > 0: 
                        suma_necesidades += necesidad
                else:
                    # Garantiza que el dict_pesos lea la letra exacta (A, B, C, D)
                    peso_final = float(dict_pesos.get(grade_tienda, 0))

                # Filtros Top Tier y Clima
                if (str(row['Grade']).strip().upper() == 'TOP TIER' and grade_tienda in ['C', 'D']) or \
                   (temporada_backend != "ambos" and str(row['Grade']).strip().upper() == 'CLIMATE SPECIFIC' and clima_tienda != temporada_backend):
                    peso_final = 0.0

                pesos_fila.append(peso_final)
            
            df_pesos.loc[idx] = pesos_fila

            # MAGIA HÍBRIDA
            if es_reposicion:
                max_unidades = min(suma_necesidades, soh_bodega)
            else:
                limite_decimal = (float(max_send_pct) + float(flex_margin)) / 100.0
                max_unidades = soh_bodega * limite_decimal * norm_curve
            
            allocable_real.append(max(0.0, min(max_unidades, soh_bodega)))
    
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
            
        # --- SECCIÓN VISUAL CON TABS ---
        st.markdown("---")
        fecha_mostrar = st.session_state.fecha_es if idioma == "Español" else st.session_state.fecha_en
        
        # CREACIÓN DE PESTAÑAS
        tab_principal, tab_comparador = st.tabs(["📊 Dashboard Principal", "⚖️ Comparador What-If"])

        with tab_principal:
            st.subheader(f"{txt['metrics_title']} ({fecha_mostrar})")
            
            st.markdown("---")
            st.subheader(txt["matrix_title"])
            
            if idioma == "Español":
                st.info("💡 **Ajuste Fino:** Haz doble clic en cualquier celda debajo del nombre de una tienda para modificar la cantidad.")
            else:
                st.info("💡 **Fine-tuning:** Double-click any cell under a store name to manually modify the quantity.")

            # 1. CREAMOS EL EDITOR PRIMERO (Esto define 'df_editado')
            columnas_protegidas = ['SKU', 'Product_Name', 'Size', 'Gender', 'Gender_&_Category', 'LSKD_DC_SOH', 'Grade', 'Curve_Multiplier', 'Norm_Curve', 'Max_Allocable']
            df_editado = st.data_editor(
                df_resultado, disabled=columnas_protegidas, use_container_width=True, hide_index=True, key=f"editor_matriz_{st.session_state.editor_key}"
            )

            # 2. AHORA SÍ CALCULAMOS LAS MÉTRICAS GLOBALES BASADAS EN LO EDITADO
            df_solo_tiendas = df_editado[tiendas_destino]
            total_inventario = df_editado['LSKD_DC_SOH'].sum()
            total_asignado = df_solo_tiendas.sum().sum()
            pct_global = (total_asignado / total_inventario * 100) if total_inventario > 0 else 0

            cont_metricas = st.container()
            cont_alertas = st.container()
            cont_graficos = st.container()

            with cont_metricas:
                col1, col2, col3 = st.columns(3)
                col1.metric(txt["metric_inv"], f"{int(total_inventario):,}")
                col2.metric(txt["metric_dist"], f"{int(total_asignado):,}")
                col3.metric(txt["metric_pct"], f"{pct_global:.1f}%")

            with cont_alertas:
                # Ajuste de precisión a 1 decimal para evitar falsas alertas por coma flotante
                limite_porcentual = round(float(max_send_pct + flex_margin), 1)
                asignacion_porcentual = round(float(pct_global), 1)
                
                if asignacion_porcentual > limite_porcentual:
                    st.warning(f"⚠️ Límite Excedido: Tu asignación ({asignacion_porcentual:.1f}%) supera la regla máxima permitida ({limite_porcentual:.1f}%).")
                    if st.button("🔄 Restaurar Regla Matemática"):
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
                    
                    orden_tallas_retail = ['2XS', 'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', 'ONE SIZE']
                    df_tallas['Size'] = pd.Categorical(df_tallas['Size'], categories=orden_tallas_retail, ordered=True)
                    df_tallas = df_tallas.sort_values('Size')
                    
                    fig_size = px.bar(
                        df_tallas, x='Size', y='Total_Asignado',
                        color='Total_Asignado', color_continuous_scale='Teal', text_auto=True
                    )
                    fig_size.update_xaxes(categoryorder='array', categoryarray=df_tallas['Size'].astype(str))
                    fig_size.update_layout(margin=dict(t=10, b=10))
                    st.plotly_chart(fig_size, use_container_width=True)

        with tab_comparador:
            st.markdown("### ⚖️ Comparador de Escenarios")
            st.info("💡 **Instrucciones:** Ajusta los parámetros (ej. WOC o Límite) en el panel izquierdo. Guarda el Escenario A. Luego, cambia los parámetros y guarda el Escenario B para compararlos lado a lado.")
            
            col_a, col_b = st.columns(2)
            
            with col_a:
                if st.button("📸 Guardar como Escenario A", use_container_width=True):
                    st.session_state.escenario_a = {
                        "nombre": f"WOC: {target_woc} | Límite: {max_send_pct}%",
                        "unidades": int(total_asignado),
                        "pct": pct_global,
                        "datos": df_solo_tiendas.sum().to_dict()
                    }
                if st.session_state.escenario_a:
                    st.success("✅ Escenario A cargado")
                    st.metric("Parámetros (A)", st.session_state.escenario_a["nombre"], f"{st.session_state.escenario_a['unidades']} unds ({st.session_state.escenario_a['pct']:.1f}%)")

            with col_b:
                if st.button("📸 Guardar como Escenario B", use_container_width=True):
                    st.session_state.escenario_b = {
                        "nombre": f"WOC: {target_woc} | Límite: {max_send_pct}%",
                        "unidades": int(total_asignado),
                        "pct": pct_global,
                        "datos": df_solo_tiendas.sum().to_dict()
                    }
                if st.session_state.escenario_b:
                    st.success("✅ Escenario B cargado")
                    st.metric("Parámetros (B)", st.session_state.escenario_b["nombre"], f"{st.session_state.escenario_b['unidades']} unds ({st.session_state.escenario_b['pct']:.1f}%)")

            # Mostrar gráfico comparativo si ambos existen
            if st.session_state.escenario_a and st.session_state.escenario_b:
                st.markdown("---")
                st.subheader("📈 Diferencia de Asignación por Tienda")
                
                # Preparar datos para el gráfico agrupado
                tiendas_a = list(st.session_state.escenario_a["datos"].keys())
                unidades_a = list(st.session_state.escenario_a["datos"].values())
                tiendas_b = list(st.session_state.escenario_b["datos"].keys())
                unidades_b = list(st.session_state.escenario_b["datos"].values())
                
                df_comp = pd.DataFrame({
                    "Tienda": tiendas_a + tiendas_b,
                    "Unidades": unidades_a + unidades_b,
                    "Escenario": ["A: " + st.session_state.escenario_a["nombre"]] * len(tiendas_a) + ["B: " + st.session_state.escenario_b["nombre"]] * len(tiendas_b)
                })
                
                fig_comp = px.bar(
                    df_comp, x="Tienda", y="Unidades", color="Escenario", 
                    barmode="group", color_discrete_sequence=["#1f77b4", "#ff7f0e"], text_auto=True
                )
                fig_comp.update_layout(xaxis_tickangle=-45, margin=dict(t=10, b=10))
                st.plotly_chart(fig_comp, use_container_width=True)
                
                if st.button("🗑️ Limpiar Escenarios"):
                    st.session_state.escenario_a = None
                    st.session_state.escenario_b = None
                    st.rerun()

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
