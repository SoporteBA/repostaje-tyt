import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Procesador de Repostajes", page_icon="⛽")

st.title("⛽ Gestor de Transacciones CEPSA")
st.write("Sube los archivos necesarios para generar el informe importable.")

# --- FUNCIÓN DE PROCESAMIENTO ---
def procesar_archivos(plantilla, datos):
    # Leer datos
    try:
        df_origen = pd.read_excel(datos, header=2, dtype={'Tarjeta': str})
    except:
        st.error("Error leyendo el archivo de datos. Verifica que el encabezado esté en la fila 3.")
        return None

    # Leer plantilla
    try:
        wb_destino = load_workbook(plantilla)
        ws_destino = wb_destino.active
    except:
        st.error("Error leyendo la plantilla.")
        return None

    # Limpiar plantilla
    max_row = ws_destino.max_row
    max_col = ws_destino.max_column
    if max_row > 1:
        for row in ws_destino.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
            for cell in row:
                cell.value = None

    # Configuración
    nueva_fila_inicio = 2
    matriculas_excluidas = ["TJT-001", "TJT-002", "TJT-003", "TJT-004", "TJT-005", "TJT-006", "TJT-007"]
    fila_destino = 2 

    # Barra de progreso
    barra = st.progress(0)
    total_filas = len(df_origen)

    for index, row in df_origen.iterrows():
        # Actualizar barra
        barra.progress((index + 1) / total_filas)

        # Filtro Matrícula
        matricula_actual = str(row['Matricula']).strip()
        if matricula_actual in matriculas_excluidas:
            continue

        # Procesar Fechas
        fecha_hora_raw = row['Fecha y hora']
        val_fecha = None
        val_hora = None
        if pd.notnull(fecha_hora_raw):
            try:
                if isinstance(fecha_hora_raw, datetime):
                    val_fecha = fecha_hora_raw
                    val_hora = fecha_hora_raw
                else:
                    dt_obj = datetime.strptime(str(fecha_hora_raw), "%d/%m/%Y %H:%M:%S")
                    val_fecha = dt_obj
                    val_hora = dt_obj
            except:
                pass 

        # Tarjeta
        tarjeta_valor = str(row['Tarjeta']) if pd.notnull(row['Tarjeta']) else ""
        if tarjeta_valor.endswith('.0'): tarjeta_valor = tarjeta_valor[:-2]

        # Traducción Productos
        concepto_original = df_origen.iloc[index, 10]
        producto_final = concepto_original
        if pd.notnull(concepto_original):
            nombre_concepto = str(concepto_original).strip()
            if nombre_concepto == "DIESEL STAR": producto_final = "Gasoleo"
            elif nombre_concepto == "ECOBLUE": producto_final = "AdBlue"
            elif nombre_concepto == "SIN PLOMO": producto_final = "Gasoil B"
            elif nombre_concepto == "AUTOPISTAS DE PEAJE": producto_final = "Peaje"
            elif nombre_concepto == "GEST. SERV. AUTOP. ESPAÑA": producto_final = "Otros"

        # Escribir
        ws_destino.cell(row=fila_destino, column=1).value = row['Matricula']
        ws_destino.cell(row=fila_destino, column=2).value = producto_final
        ws_destino.cell(row=fila_destino, column=3).value = df_origen.iloc[index, 4]
        
        c_tarjeta = ws_destino.cell(row=fila_destino, column=4)
        c_tarjeta.value = tarjeta_valor
        c_tarjeta.number_format = '@'
        
        ws_destino.cell(row=fila_destino, column=5).value = df_origen.iloc[index, 11]
        ws_destino.cell(row=fila_destino, column=6).value = df_origen.iloc[index, 8]
        ws_destino.cell(row=fila_destino, column=7).value = df_origen.iloc[index, 12]
        ws_destino.cell(row=fila_destino, column=8).value = df_origen.iloc[index, 13]

        if val_fecha:
            c_f = ws_destino.cell(row=fila_destino, column=11)
            c_f.value = val_fecha
            c_f.number_format = 'dd/mm/yyyy'
        
        if val_hora:
            c_h = ws_destino.cell(row=fila_destino, column=12)
            c_h.value = val_hora
            c_h.number_format = 'hh:mm:ss'

        fila_destino += 1

    # Guardar en memoria virtual
    output = io.BytesIO()
    wb_destino.save(output)
    output.seek(0)
    return output

# --- INTERFAZ DE USUARIO ---
col1, col2 = st.columns(2)
with col1:
    uploaded_plantilla = st.file_uploader("1. Sube la Plantilla (ImportadorGenerico)", type="xlsx")
with col2:
    uploaded_datos = st.file_uploader("2. Sube los Datos (transacciones-cepsa)", type="xlsx")

if uploaded_plantilla and uploaded_datos:
    if st.button("🚀 Procesar Archivos"):
        with st.spinner('Procesando datos...'):
            archivo_final = procesar_archivos(uploaded_plantilla, uploaded_datos)
            
            if archivo_final:
                st.success("¡Hecho! Descarga tu archivo aquí abajo:")
                st.download_button(
                    label="⬇️ Descargar Excel Relleno",
                    data=archivo_final,
                    file_name="ImportadorGenerico_RELLENO.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
