import streamlit as st
import pandas as pd
from datetime import datetime
import io
import re

# --- CONFIGURACIÓN DE ETIQUETAS ---
PROB_MAP = {
    "0%": "Nula",
    "25%": "Remota",
    "50%": "Podría Ser",
    "75%": "Altamente probable",
    "100%": "Cierta"
}

# Columnas definitivas para el reporte de salida
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
    # 1. Cargar Reporte Nuevo (Títulos en fila 6)
    df_nuevo = pd.read_excel(archivo_nuevo, skiprows=5)
    df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]
    df_nuevo = df_nuevo.dropna(how='all', axis=0)

    # 2. Identificar Automáticamente la Última Hoja por Fecha o Posición
    xl_historial = pd.ExcelFile(archivo_historial)
    hojas = xl_historial.sheet_names
    hoja_maestra = None
    fecha_reciente = datetime.min
    
    for h in hojas:
        match = re.search(r'(\d{2}-\d{2}-\d{2})', h)
        if match:
            try:
                fecha_hoja = datetime.strptime(match.group(1), "%d-%m-%y")
                if fecha_hoja > fecha_reciente:
                    fecha_reciente = fecha_hoja
                    hoja_maestra = h
            except:
                continue
    
    if not hoja_maestra:
        posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
        for h in reversed(hojas):
            df_check = pd.read_excel(xl_historial, sheet_name=h, nrows=10, header=None)
            if any(str(val).strip() in posibles_nombres for row in df_check.values for val in row):
                hoja_maestra = h
                break

    if not hoja_maestra:
        st.error("No se pudo identificar la hoja de datos en el historial.")
    else:
        st.info(f"Última actualización detectada: **{hoja_maestra}**")
        
        # Leer historial detectando fila de encabezado
        df_hist_raw = pd.read_excel(xl_historial, sheet_name=hoja_maestra, header=None)
        fila_h = 0
        for i, row in df_hist_raw.iterrows():
            if any(str(val).strip() in ['Número de caso', 'Numero de caso', 'N° caso', 'Caso'] for val in row.values):
                fila_h = i
                break
        
        df_hist = pd.read_excel(xl_historial, sheet_name=hoja_maestra, skiprows=fila_h)
        df_hist.columns = [str(c).strip() for c in df_hist.columns]
        
        col_llave = next((c for c in df_nuevo.columns if c in ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']), None)
        
        if col_llave:
            # Estandarización de llaves (Texto)
            df_nuevo[col_llave] = df_nuevo[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_hist[col_llave] = df_hist[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

            # Preparar columnas de persistencia
            cols_persistencia = [col_llave, 'Probabilidad cierre 2026', 'Observaciones', 'Fecha probable de facturación']
            for c in cols_persistencia:
                if c not in df_hist.columns: df_hist[c] = ""

            # --- CRUCE DE DATOS ---
            df_final = pd.merge(df_nuevo, df_hist[cols_persistencia], on=col_llave, how='left')

            # --- FORMATEO DE TIPOS PARA EL EDITOR ---
            def to_pct_str(val):
                try:
                    num = float(val)
                    if num <= 1.0: return f"{int(num * 100)}%"
                    return f"{int(num)}%"
                except: return "0%"

            df_final['Probabilidad cierre 2026'] = df_final['Probabilidad cierre 2026'].apply(to_pct_str)
            df_final['Observaciones'] = df_final['Observaciones'].astype(str).replace(['nan', 'None', '<NA>'], '')
            df_final['Fecha probable de facturación'] = pd.to_datetime(df_final['Fecha probable de facturación'], errors='coerce').dt.date

            for col in COLUMNAS_FINALES:
                if col not in df_final.columns: df_final[col] = ""
            
            # Cálculo inicial de Honorarios Probables
            if 'Honorarios (UF)' in df_final.columns:
                df_final['Honorarios (UF)'] = pd.to_numeric(df_final['Honorarios (UF)'], errors='coerce').fillna(0)
                prob_num = df_final['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100
                df_final['Hon Probables 2026'] = df_final['Honorarios (UF)'] * prob_num

            df_final = df_final[COLUMNAS_FINALES]

            # --- CAMBIO QUIRÚRGICO: REPORTE DE CASOS NUEVOS Y SALIENTES ---
            st.subheader("Panel de Gestión Semanal")
            
            casos_viejos = set(df_hist[col_llave].unique())
            casos_actuales = set(df_nuevo[col_llave].unique())
            
            nuevos_detectados = [c for c in casos_actuales if c not in casos_viejos]
            salientes_detectados = df_hist[~df_hist[col_llave].isin(casos_actuales)]
            
            col_res1, col_res2 = st.columns(2)
            with col_res1:
                st.success(f"🆕 **Ingresos:** Se incorporaron **{len(nuevos_detectados)} casos nuevos**.")
            with col_res2:
                st.warning(f"🔴 **Salidas:** **{len(salientes_detectados)} casos** del pipeline anterior ya no están en el reporte.")
            
            # Acordeón para revisar los casos que salieron del pipeline
            if not salientes_detectados.empty:
                with st.expander("🔍 Ver listado de casos salientes"):
                    columnas_salientes = [col_llave, 'Nickname', 'Probabilidad cierre 2026', 'Observaciones']
                    cols_mostrar = [c for c in columnas_salientes if c in salientes_detectados.columns]
                    st.dataframe(salientes_detectados[cols_mostrar].fillna(''), hide_index=True)

            # --- FUNCIÓN DE ESTILO PARA STREAMLIT ---
            def color_semaforo(val):
                if val in ["75%", "100%"]:
                    return 'background-color: #c6efce; color: #006100;'
                elif val == "50%":
                    return 'background-color: #ffeb9c; color: #9c5700;'
                elif val in ["0%", "25%"]:
                    return 'background-color: #ffc7ce; color: #9c0006;'
                return ''

            df_styled = df_final.style.map(color_semaforo, subset=['Probabilidad cierre 2026'])

            # --- EDITOR DE DATOS ---
            df_editado = st.data_editor(
                df_styled,
                column_config={
                    "Probabilidad cierre 2026": st.column_config.SelectboxColumn(
                        "Probabilidad (%)", 
                        options=["0%", "25%", "50%", "75%", "100%"]
                    ),
                    "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                    "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
                },
                hide_index=True, 
                use_container_width=True
            )

            # --- RECALCULAR TRAS EDICIÓN ---
            prob_num_final = df_editado['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100
            df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * prob_num_final
            df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
            
            st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

            # --- CAMBIO QUIRÚRGICO: FORMATEO CONDICIONAL NATIVO PARA EXCEL ---
            fecha_desc = datetime.now().strftime("%d-%m-%y")
            buffer = io.BytesIO()
            df_excel = df_editado.copy()
            
            # Devolver a decimal para que Excel lo calcule matemáticamente
            df_excel['Probabilidad cierre 2026'] = df_excel['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100

            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                nombre_hoja_descarga = f"Casos {fecha_desc}"
                df_excel.to_excel(writer, sheet_name=nombre_hoja_descarga, index=False)
                
                # Obtener los objetos del libro y la hoja para inyectar formatos
                workbook = writer.book
                worksheet = writer.sheets[nombre_hoja_descarga]
                
                # Crear los formatos de Excel
                formato_pct = workbook.add_format({'num_format': '0%'})
                formato_verde = workbook.add_format({'bg_color': '#c6efce', 'font_color': '#006100'})
                formato_amarillo = workbook.add_format({'bg_color': '#ffeb9c', 'font_color': '#9c5700'})
                formato_rojo = workbook.add_format({'bg_color': '#ffc7ce', 'font_color': '#9c0006'})

                # Encontrar el índice numérico de la columna de Probabilidad (base 0)
                idx_prob = COLUMNAS_FINALES.index('Probabilidad cierre 2026')
                
                # Aplicar formato de porcentaje a toda la columna
                worksheet.set_column(idx_prob, idx_prob, 15, formato_pct)
                
                # Aplicar el Semáforo Condicional directo al Excel
                filas_totales = len(df_excel)
                worksheet.conditional_format(1, idx_prob, filas_totales, idx_prob, 
                                             {'type': 'cell', 'criteria': '>=', 'value': 0.75, 'format': formato_verde})
                worksheet.conditional_format(1, idx_prob, filas_totales, idx_prob, 
                                             {'type': 'cell', 'criteria': '==', 'value': 0.50, 'format': formato_amarillo})
                worksheet.conditional_format(1, idx_prob, filas_totales, idx_prob, 
                                             {'type': 'cell', 'criteria': '<=', 'value': 0.25, 'format': formato_rojo})
            
            st.sidebar.divider()
            st.sidebar.download_button(
                label="📥 Descargar Pipeline Formateado",
                data=buffer.getvalue(),
                file_name=f"JPV_Pipeline_{fecha_desc}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Sube los archivos para procesar el Pipeline. El sistema reportará ingresos, salidas y aplicará el formato al Excel.")
