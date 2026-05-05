import streamlit as st
import pandas as pd
import numpy as np

# Configuración de la página
st.set_page_config(page_title="LKSD Allocation Model", layout="wide")

st.title("📦 Sistema de Asignación Semanal: LSKD")
st.markdown("Carga el archivo estándar y ajusta los parámetros para generar la distribución por tiendas.")

# 1. Carga del archivo
uploaded_file = st.file_uploader("Sube el archivo LSKD_Newness_EMB.xlsx", type=["xlsx"])

if uploaded_file:
    with st.spinner('Cargando y procesando datos...'):
        try:
            # Cargar el archivo en la memoria del servidor una sola vez
            xls = pd.ExcelFile(uploaded_file)
            
            # Extraer solo las hojas que necesitamos
            df_newness = pd.read_excel(xls, sheet_name='Newness')
            df_stores = pd.read_excel(xls, sheet_name='Store grading')
            df_curve = pd.read_excel(xls, sheet_name='Size Curve')
            
            st.success("Archivo cargado correctamente.")
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo: {e}")
            st.stop() # Detiene la ejecución si hay error para no dejar la pantalla en blanco

    st.success("Archivo cargado correctamente.")

    # 2. Panel de Control (Sidebar)
    st.sidebar.header("⚙️ Parámetros de Asignación")
    
    st.sidebar.subheader("Reglas de Bodega")
    # Límite máximo a enviar (ideal 30%, flexibilidad hasta 34%)
    max_send_pct = st.sidebar.slider("Límite Ideal a Enviar (%)", min_value=10.0, max_value=40.0, value=30.0, step=1.0)
    flex_margin = st.sidebar.slider("Margen de Flexibilidad (%)", min_value=0.0, max_value=5.0, value=3.0, step=0.5)
    
    st.sidebar.subheader("Pesos por Categoría de Tienda")
    st.sidebar.markdown("Define la proporción del inventario a distribuir según el 'Grade' de la tienda.")
    
    # Sliders para las tiendas
    peso_a = st.sidebar.number_input("Peso Tiendas A", value=40)
    peso_b = st.sidebar.number_input("Peso Tiendas B", value=30)
    peso_c = st.sidebar.number_input("Peso Tiendas C", value=20)
    peso_d = st.sidebar.number_input("Peso Tiendas D", value=10)
    
    # Normalización de pesos (por si el usuario ingresa valores que no sumen 100)
    total_peso = peso_a + peso_b + peso_c + peso_d
    if total_peso == 0: total_peso = 1 # Evitar división por cero
    
    dict_pesos = {
        'A': peso_a / total_peso,
        'B': peso_b / total_peso,
        'C': peso_c / total_peso,
        'D': peso_d / total_peso
    }

    # 3. Procesamiento y Lógica del Modelo
    st.header("📊 Resultados del Modelo de Asignación")
    
    # --- PASO A: Calcular tope máximo de envío por SKU ---
    # Asumimos que 'LSKD DC SOH' es el inventario total en bodega
    if 'LSKD DC SOH' in df_newness.columns:
        # Límite duro (30% + flexibilidad)
        limite_absoluto = (max_send_pct + flex_margin) / 100.0
        
        # Crear columna de unidades máximas a distribuir
        df_newness['Max_Allocable'] = np.floor(df_newness['LSKD DC SOH'] * limite_absoluto)
        
        # --- PASO B: Mapeo de Tiendas ---
        # Limpiar espacios en nombres de columnas si los hay
        df_stores.columns = df_stores.columns.str.strip()
        
        # Crear un diccionario rápido para saber el Grade de cada tienda (usamos Womens por defecto para el ejemplo)
        # Se podría cruzar con 'Mens Allocations Grade' dependiendo de la columna 'Gender' en df_newness
        store_grades = pd.Series(df_stores['Womens Allocation Grade'].values, index=df_stores['Store']).to_dict()
        
        # --- PASO C: Distribución Teórica (Ejemplo simplificado) ---
        # Aquí crearíamos la matriz de distribución iterando sobre las tiendas
        tiendas_destino = df_stores['Store'].tolist()
        
        # Dataframe para el resultado final
        df_resultado = df_newness[['SKU', 'Product Name', 'Size', 'Gender', 'LSKD DC SOH', 'Max_Allocable', 'Grade']].copy()
        
        # Lógica de distribución por tienda
        for tienda in tiendas_destino:
            grade_tienda = store_grades.get(tienda, 'C') # Default a C si no se encuentra
            peso_tienda = dict_pesos.get(grade_tienda, 0)
            
            # Cálculo de unidades para esta tienda basado en su peso. 
            # NOTA: En un modelo real de producción, aquí cruzaríamos con la curva de tallas (df_curve)
            # multiplicando este valor base por el coeficiente de la talla S, M o L.
            df_resultado[tienda] = np.floor(df_resultado['Max_Allocable'] * peso_tienda)
            
            # Filtro de "No todo va a todas las tiendas"
            # Si el producto dice "TOP TIER", borramos las unidades de las tiendas C y D
            mask_toptier = (df_resultado['Grade'] == 'TOP TIER') & (grade_tienda in ['C', 'D'])
            df_resultado.loc[mask_toptier, tienda] = 0

        # Mostrar métricas clave de la corrida
        total_bodega = df_newness['LSKD DC SOH'].sum()
        total_asignado = df_resultado[tiendas_destino].sum().sum()
        pct_asignado_real = (total_asignado / total_bodega) * 100 if total_bodega > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Inventario Total en Bodega", f"{int(total_bodega):,}")
        col2.metric("Unidades Asignadas", f"{int(total_asignado):,}")
        
        # Cambiar color si se pasa del 30%
        delta_color = "normal" if pct_asignado_real <= max_send_pct else "inverse"
        col3.metric("% Real Asignado", f"{pct_asignado_real:.1f}%", delta=f"{pct_asignado_real - max_send_pct:.1f}% vs Ideal", delta_color=delta_color)

        # 4. Vista previa y Descarga
        st.subheader("Vista Previa de Distribución por Tienda")
        st.dataframe(df_resultado.head(50))
        
        # Generar CSV para descarga
        csv = df_resultado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar Matriz de Asignación (CSV)",
            data=csv,
            file_name='Asignacion_LSKD_Generada.csv',
            mime='text/csv',
        )
    else:
        st.error("No se encontró la columna 'LSKD DC SOH' en la hoja de Newness. Verifica el archivo.")
