import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURACIÓN ---
PROB_MAP = {
    0.0: "Nula",
    0.25: "Remota",
    0.50: "Podría Ser",
    0.75: "Altamente probable",
    1.0: "Cierta"
}

st.set_page_config(page_title="JPV Pipeline y Seguimiento", layout="wide")
st.title("🚀 JPV: Pipeline de Facturación Probable")

# --- FUNCIÓN DE CARGA SEGURA ---
def cargar_excel_limpio(archivo):
    if archivo is None: return None
    # Leer excel
    df = pd.read_excel(archivo)
    
    # Si la primera columna parece basura o vacía, reintentar buscando el encabezado
    if "Unnamed" in str(df.columns[0]) or df.columns[0] == "" or len(df.columns) < 2:
        df = pd.read_excel(archivo, skiprows=1) # Salta la primera fila de título si existe
        
    # Limpiar nombres de columnas (quitar espacios y saltos de línea)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- SIDEBAR ---
st.sidebar.header("Carga de Documentos")
archivo_nuevo = st.sidebar.file_uploader("1. Nuevo Reporte de Acciones (Excel)", type=["xlsx"])
archivo_historial = st.sidebar.file_uploader("2. Pipeline Anterior (Archivo Maestro)", type=["xlsx"])

if archivo_nuevo and archivo_historial:
    df_nuevo = cargar_excel_limpio(archivo_nuevo)
    
    xl = pd.ExcelFile(archivo_historial)
    hoja_reciente = xl.sheet_names[0]
    df_hist = cargar_excel_limpio(archivo_historial) # Carga la primera hoja por defecto

    # --- VALIDACIÓN DE LLAVE (KeyError Fix) ---
    # Buscamos cómo se llama la columna de 'Número de caso' realmente
    posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
    col_llave = next((c for c in df_nuevo.columns if c in posibles_nombres), None)

    if not col_llave:
        st.error(f"No encontré la columna 'Número de caso'. Columnas detectadas: {list(df_nuevo.columns[:5])}...")
    else:
        # Aseguramos que las columnas de persistencia existan
        cols_manuales = [col_llave, 'Probabilidad cierre 2026', 'Observaciones', 'Fecha probable de facturación']
        for c in cols_manuales:
            if c not in df_hist.columns:
                df_hist[c] = 0.0 if 'Probabilidad' in c else ""

        # Cruce Seguro
        df_final = pd.merge(
            df_nuevo, 
            df_hist[cols_manuales], 
            on=col_llave, 
            how='left', 
            suffixes=('', '_old')
        )

        # --- EDITOR Y CÁLCULOS ---
        df_final['Probabilidad cierre 2026'] = df_final['Probabilidad cierre 2026'].fillna(0.0)
        
        st.subheader("Panel de Edición")
        df_editado = st.data_editor(
            df_final,
            column_config={
                "Probabilidad cierre 2026": st.column_config.SelectboxColumn("Probabilidad (%)", options=[0.0, 0.25, 0.50, 0.75, 1.0]),
                "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
            },
            hide_index=True,
            use_container_width=True
        )

        # Mapeo de etiquetas y KPI
        df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
        
        # Buscar columna de Honorarios para el total
        col_hon = next((c for c in df_editado.columns if 'Honorarios' in c), None)
        if col_hon:
            df_editado[col_hon] = pd.to_numeric(df_editado[col_hon], errors='coerce').fillna(0)
            df_editado['Hon Probables 2026'] = df_editado[col_hon] * df_editado['Probabilidad cierre 2026']
            total_uf = df_editado['Hon Probables 2026'].sum()
            st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{total_uf:,.2f}")

        # --- DESCARGA ---
        fecha_hoy = datetime.now().strftime("%d-%m-%y")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_editado.to_excel(writer, sheet_name=f"Casos {fecha_hoy}", index=False)
        
        st.sidebar.download_button(
            label="📥 Descargar Pipeline",
            data=buffer.getvalue(),
            file_name=f"Pipeline_{fecha_hoy}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Sube ambos archivos Excel para procesar el cruce.")
