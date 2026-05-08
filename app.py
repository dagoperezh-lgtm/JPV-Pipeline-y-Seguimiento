import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURACIÓN Y MAPEOS ---
PROB_MAP = {
    0.0: "Nula",
    0.25: "Remota",
    0.50: "Podría Ser",
    0.75: "Altamente probable",
    1.0: "Cierta"
}

st.set_page_config(page_title="Pipeline de Facturación", layout="wide")

st.title("🚀 Pipeline de Facturación Probable")
st.markdown("""
Esta herramienta automatiza el cruce entre el **Reporte de Acciones** del sistema y el **Pipeline Histórico**, 
manteniendo observaciones, probabilidades y fechas de facturación de la semana anterior.
""")

# --- 1. CARGA DE ARCHIVOS ---
st.sidebar.header("Carga de Datos")
archivo_nuevo = st.sidebar.file_uploader("1. Nuevo Reporte de Acciones (CSV)", type=["csv"])
archivo_historial = st.sidebar.file_uploader("2. Pipeline de la Semana Pasada (Excel)", type=["xlsx"])

if archivo_nuevo and archivo_historial:
    # Leer el reporte nuevo
    df_nuevo = pd.read_csv(archivo_nuevo)
    
    # Leer el historial (buscamos la hoja más reciente o la primera)
    xl = pd.ExcelFile(archivo_historial)
    hoja_anterior = xl.sheet_names[0] # Asumimos que la primera hoja es la última actualizada
    df_hist = pd.read_excel(archivo_historial, sheet_name=hoja_anterior)
    
    # --- 2. TRATAMIENTO DE COLUMNAS ---
    # Columnas que queremos traer del pasado (Persistencia)
    cols_manuales = [
        'Número de caso', 
        'Probabilidad cierre 2026', 
        'Observaciones', 
        'Fecha probable de facturación'
    ]
    
    # Aseguramos que existan en el historial para evitar errores
    for c in cols_manuales:
        if c not in df_hist.columns:
            df_hist[c] = None if c != 'Probabilidad cierre 2026' else 0.0

    # Cruce de datos (Merge)
    # Traemos la información del reporte nuevo y le pegamos lo que ya teníamos anotado
    df_merge = pd.merge(
        df_nuevo, 
        df_hist[cols_manuales], 
        on='Número de caso', 
        how='left', 
        suffixes=('', '_old')
    )

    # Limpieza de duplicados tras el merge si existen
    df_merge['Probabilidad cierre 2026'] = df_merge['Probabilidad cierre 2026'].fillna(0.0)
    
    # --- 3. EDITOR DE DATOS (UI INTERACTIVA) ---
    st.subheader(f"Edición de Datos - Actualización Semanal")
    st.info("Modifica la Probabilidad, Fecha u Observaciones. Los cálculos se actualizarán automáticamente.")

    # Definir configuración de columnas para el editor
    config_columnas = {
        "Probabilidad cierre 2026": st.column_config.SelectboxColumn(
            "Probabilidad (%)",
            options=[0.0, 0.25, 0.50, 0.75, 1.0],
            required=True
        ),
        "Fecha probable de facturación": st.column_config.DateColumn("Fecha Prob. Facturación"),
        "Observaciones": st.column_config.TextColumn("Observaciones", width="large"),
        "Indicación Probabilidad": st.column_config.TextColumn("Estado", disabled=True),
        "Hon Probables 2026": st.column_config.NumberColumn("Hon Probables (UF)", format="%.2f", disabled=True),
        "Honorarios (UF)": st.column_config.NumberColumn("Hon (UF)", disabled=True)
    }

    # El editor
    df_editado = st.data_editor(
        df_merge,
        column_config=config_columnas,
        hide_index=True,
        use_container_width=True
    )

    # --- 4. CÁLCULOS AUTOMÁTICOS ---
    # Aplicamos lógica de etiquetas y cálculo de honorarios probables
    df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
    
    # Asegurar que Honorarios (UF) sea numérico para el cálculo
    df_editado['Honorarios (UF)'] = pd.to_numeric(df_editado['Honorarios (UF)'], errors='coerce').fillna(0)
    df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * df_editado['Probabilidad cierre 2026']

    # --- 5. INDICADORES (KPIs) ---
    total_facturacion = df_editado['Hon Probables 2026'].sum()
    
    st.divider()
    col_kpi1, col_kpi2 = st.columns(2)
    with col_kpi1:
        st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{total_facturacion:,.2f}")
    with col_kpi2:
        st.metric("Casos en Pipeline", len(df_editado))

    # --- 6. GENERACIÓN DEL EXCEL CON NUEVA HOJA ---
    fecha_hoy = datetime.now().strftime("%d-%m-%y")
    nombre_nueva_hoja = f"Casos {fecha_hoy}"

    # Botón para descargar
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Guardamos la nueva hoja al principio
        df_editado.to_excel(writer, sheet_name=nombre_nueva_hoja, index=False)
        
        # Opcional: Re-escribir las hojas antiguas para mantener el histórico en un solo archivo
        for sheet in xl.sheet_names:
            if sheet != nombre_nueva_hoja: # Evitar duplicar la de hoy si ya existiera
                temp_df = pd.read_excel(archivo_historial, sheet_name=sheet)
                temp_df.to_excel(writer, sheet_name=sheet, index=False)

    st.sidebar.divider()
    st.sidebar.download_button(
        label="📥 Descargar Pipeline Actualizado",
        data=output.getvalue(),
        file_name=f"Pipeline_Facturacion_{fecha_hoy}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("👈 Por favor, carga los archivos en la barra lateral para comenzar.")
    # Mostrar una imagen o guía de ayuda si es necesario
