import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURACIÓN DE ETIQUETAS ---
PROB_MAP = {
    0.0: "Nula",
    0.25: "Remota",
    0.50: "Podría Ser",
    0.75: "Altamente probable",
    1.0: "Cierta"
}

# Columnas definitivas que debe tener el archivo de salida
COLUMNAS_FINALES = [
    'Número de caso', 'Número de siniestro', 'Nickname', 'División', 
    'Compañía de seguros', 'Corredora', 'Ajustador senior', 'Asegurado', 
    'Creado en', 'Divisa', 'Perdida bruta (en moneda del caso)', 
    'Deducible (en moneda del caso)', 'Monto asegurado (en moneda del caso)', 
    'Honorarios (UF)', 'Facturado', 'Último movimiento', 
    'Contenido último movimiento', 'Probabilidad cierre 2026', 
    'Indicación Probabilidad', 'Hon Probables 2026', 'Observaciones', 
    'Fecha probable de facturación'
]

st.set_page_config(page_title="JPV Pipeline y Seguimiento", layout="wide")
st.title("🚀 JPV: Pipeline de Facturación Probable")

st.sidebar.header("Carga de Documentos")
archivo_nuevo = st.sidebar.file_uploader("1. Nuevo Reporte de Acciones (Excel)", type=["xlsx"])
archivo_historial = st.sidebar.file_uploader("2. Pipeline Anterior (Excel Maestro)", type=["xlsx"])

if archivo_nuevo and archivo_historial:
    # --- CAMBIO QUIRÚRGICO: SELECCIÓN DE HOJA Y BÚSQUEDA INTELIGENTE DE ENCABEZADOS ---
    
    # 1. Cargar Reporte Nuevo (Siempre sabemos que los títulos están en la fila 6 -> skiprows=5)
    df_nuevo = pd.read_excel(archivo_nuevo, skiprows=5)
    df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]
    df_nuevo = df_nuevo.dropna(how='all', axis=0)

    # 2. Cargar Historial y dejar que el usuario elija la hoja
    xl_historial = pd.ExcelFile(archivo_historial)
    hojas_disponibles = xl_historial.sheet_names
    
    st.sidebar.subheader("Configuración del Cruce")
    hoja_seleccionada = st.sidebar.selectbox("Selecciona la hoja de la semana pasada:", hojas_disponibles)
    
    # Leer la hoja seleccionada sin saltar filas inicialmente para buscar el encabezado
    df_hist_raw = pd.read_excel(xl_historial, sheet_name=hoja_seleccionada, header=None)
    
    posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
    fila_header = 0
    
    # Escanear fila por fila hasta encontrar los títulos
    for i, row in df_hist_raw.iterrows():
        if any(str(val).strip() in posibles_nombres for val in row.values):
            fila_header = i
            break
            
    # Leer la hoja histórica aplicando el salto de filas exacto detectado
    df_hist = pd.read_excel(xl_historial, sheet_name=hoja_seleccionada, skiprows=fila_header)
    df_hist.columns = [str(c).strip() for c in df_hist.columns]
    df_hist = df_hist.dropna(how='all', axis=0)

    col_llave = next((c for c in df_nuevo.columns if c in posibles_nombres), None)
    
    if not col_llave:
        st.error("No se encontró la columna clave en el reporte de acciones.")
    elif col_llave not in df_hist.columns:
        st.warning(f"La hoja '{hoja_seleccionada}' no contiene la columna '{col_llave}'. Por favor selecciona otra hoja en la barra lateral.")
    else:
        st.success(f"Cruce exitoso utilizando la hoja histórica: **{hoja_seleccionada}**")

        # Estandarizar llave como texto para evitar el ValueError
        df_nuevo[col_llave] = df_nuevo[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_hist[col_llave] = df_hist[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        # Preparar columnas a recuperar del pasado
        cols_a_recuperar = [col_llave, 'Probabilidad cierre 2026', 'Observaciones', 'Fecha probable de facturación']
        for c in cols_a_recuperar:
            if c not in df_hist.columns: df_hist[c] = ""

        # CRUCE DE DATOS
        df_final = pd.merge(df_nuevo, df_hist[cols_a_recuperar], on=col_llave, how='left')

        # Limpieza post-merge de compatibilidad para Streamlit
        df_final['Probabilidad cierre 2026'] = pd.to_numeric(df_final['Probabilidad cierre 2026'], errors='coerce').fillna(0.0)
        df_final['Observaciones'] = df_final['Observaciones'].astype(str).replace(['nan', 'None', '<NA>'], '')
        df_final['Fecha probable de facturación'] = pd.to_datetime(df_final['Fecha probable de facturación'], errors='coerce').dt.date

        # Asegurar todas las columnas finales requeridas
        for col in COLUMNAS_FINALES:
            if col not in df_final.columns:
                df_final[col] = ""

        # Cálculos Automáticos de Honorarios Probables
        df_final['Indicación Probabilidad'] = df_final['Probabilidad cierre 2026'].map(PROB_MAP)
        if 'Honorarios (UF)' in df_final.columns:
            df_final['Honorarios (UF)'] = pd.to_numeric(df_final['Honorarios (UF)'], errors='coerce').fillna(0)
            df_final['Hon Probables 2026'] = df_final['Honorarios (UF)'] * df_final['Probabilidad cierre 2026']

        # Filtrar y ordenar
        df_final = df_final[COLUMNAS_FINALES]

        # --- PANEL DE EDICIÓN ---
        st.subheader("Panel de Gestión Semanal")
        df_editado = st.data_editor(
            df_final,
            column_config={
                "Probabilidad cierre 2026": st.column_config.SelectboxColumn("Probabilidad (%)", options=[0.0, 0.25, 0.50, 0.75, 1.0]),
                "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
            },
            hide_index=True, use_container_width=True
        )

        # Recalcular tras cualquier edición manual
        df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
        df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * df_editado['Probabilidad cierre 2026']
        
        st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

        # --- DESCARGA DEL ARCHIVO ---
        fecha_hoy = datetime.now().strftime("%d-%m-%y")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_editado.to_excel(writer, sheet_name=f"Casos {fecha_hoy}", index=False)
        
        st.sidebar.divider()
        st.sidebar.download_button(
            label="📥 Descargar Pipeline Actualizado",
            data=buffer.getvalue(),
            file_name=f"JPV_Pipeline_{fecha_hoy}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Sube los archivos en la barra lateral, selecciona la hoja del historial y el sistema procesará los datos al instante.")
