import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import datetime

# --- 1. CONFIGURACIÓN Y CONSTANTES ---
st.set_page_config(
    page_title="T&T | MOEVE > NOVATRANS",
    page_icon="🚛",
    layout="centered"
)

# Definimos las reglas de negocio al principio para fácil mantenimiento
MATRICULAS_EXCLUIDAS = {
    "TJT-001", "TJT-002", "TJT-003", "TJT-004", "TJT-005", "TJT-006", "TJT-007"
}

# Diccionario para mapear conceptos (Clave: Origen -> Valor: Destino)
MAPPING_PRODUCTOS = {
    "DIESEL STAR": "Gasoleo",
    "ECOBLUE": "AdBlue",
    "SIN PLOMO": "Gasoil B",
    "AUTOPISTAS DE PEAJE": "Peaje",
    "GEST. SERV. AUTOP. ESPAÑA": "Otros"
}

# --- 2. ESTILOS CSS ---
st.markdown("""
    <style>
    [data-testid="stFileUploader"] small { display: none; }
    [data-testid="stFileUploaderDropzone"] div div::before { content: "Arrastra y suelta tu archivo aquí"; }
    button[kind="secondary"] { background-color: #fbfbfb; border: 1px solid #d6d6d6; }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #888; font-size: 12px; }
    </style>
""", unsafe_allow_html=True)

# --- 3. LOGO Y TÍTULOS ---
col_izq, col_cen, col_der = st.columns([1, 2, 1])
with col_cen:
    try:
        st.image("tyt_logo_trans.png", use_container_width=True)
    except Exception:
        st.warning("⚠️ Logo no encontrado ('tyt_logo_trans.png').")

st.markdown("<h1 style='text-align: center;'>Creador de plantilla NOVATRANS</h1>", unsafe_allow_html=True)
st.markdown("### <div style='text-align: center;'>Importador de datos MOEVE/CEPSA</div>", unsafe_allow_html=True)
st.info("Sube el Excel de Moeve y la Plantilla vacía de Novatrans para cruzar los datos.")

# --- 4. FUNCIÓN DE PROCESAMIENTO ---
def procesar_archivos(plantilla, datos):
    # 1. Leer Datos Origen
    try:
        # header=2 implica que la fila 3 (índice 2) es la cabecera
        df_origen = pd.read_excel(datos, header=2, dtype={'Tarjeta': str})
    except Exception as e:
        st.error(f"❌ Error leyendo el archivo de Moeve: {e}")
        return None

    # Validación de seguridad: Verificar que tenemos suficientes columnas
    # Necesitamos acceder hasta el índice 13 (columna 14)
    if df_origen.shape[1] < 14:
        st.error(f"❌ El archivo de Moeve parece incompleto. Tiene {df_origen.shape[1]} columnas, pero se requieren al menos 14.")
        return None

    # 2. Leer Plantilla Destino
    try:
        wb_destino = load_workbook(plantilla)
        ws_destino = wb_destino.active
    except Exception as e:
        st.error(f"❌ Error leyendo la plantilla de Novatrans: {e}")
        return None

    # 3. Limpiar plantilla (Borrar datos anteriores si los hubiera)
    max_row = ws_destino.max_row
    max_col = ws_destino.max_column
    if max_row > 1:
        for row in ws_destino.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
            for cell in row:
                cell.value = None

    # 4. Procesamiento
    fila_destino = 2
    
    # UI: Barra de progreso
    texto_estado = st.empty()
    barra = st.progress(0)
    total_filas = len(df_origen)

    for index, row in df_origen.iterrows():
        # Actualizar UI cada 10 filas o al final para no ralentizar el bucle
        if index % 10 == 0 or index == total_filas - 1:
            progreso = (index + 1) / total_filas
            barra.progress(progreso)
            texto_estado.text(f"Procesando registro {index + 1} de {total_filas}...")

        # A. Filtro Matrícula
        matricula_actual = str(row.get('Matricula', '')).strip()
        if matricula_actual in MATRICULAS_EXCLUIDAS:
            continue

        # B. Manejo de Fechas (Pandas to_datetime es más robusto)
        fecha_hora_raw = row.get('Fecha y hora')
        val_fecha = None
        val_hora = None
        
        if pd.notnull(fecha_hora_raw):
            try:
                dt_obj = pd.to_datetime(fecha_hora_raw, dayfirst=True)
                val_fecha = dt_obj
                val_hora = dt_obj # OpenPyXL maneja el formato luego
            except:
                pass # Si falla el parseo, se queda en None

        # C. Limpieza Tarjeta
        tarjeta_valor = str(row.get('Tarjeta', ''))
        if tarjeta_valor == 'nan': tarjeta_valor = ""
        if tarjeta_valor.endswith('.0'): tarjeta_valor = tarjeta_valor[:-2]

        # D. Mapeo de Productos
        # Usamos iloc para columnas fijas como indicaste que la estructura es estable
        # Columna 10 (índice 10) es el concepto original
        concepto_original = df_origen.iloc[index, 10]
        nombre_concepto = str(concepto_original).strip() if pd.notnull(concepto_original) else ""
        
        # .get(clave, valor_por_defecto) busca en el diccionario, si no está, devuelve el original
        producto_final = MAPPING_PRODUCTOS.get(nombre_concepto, concepto_original)

        # 5. Escritura en Excel (Mapeo de Columnas)
        # Col 2 (B): Matrícula
        ws_destino.cell(row=fila_destino, column=2).value = row.get('Matricula')
        
        # Col 3 (C): Producto
        ws_destino.cell(row=fila_destino, column=3).value = producto_final
        
        # Col 4 (D): Kilómetros/Dato columna 4 origen
        ws_destino.cell(row=fila_destino, column=4).value = df_origen.iloc[index, 4]
        
        # Col 5 (E): Tarjeta (Texto)
        c_tarjeta = ws_destino.cell(row=fila_destino, column=5)
        c_tarjeta.value = tarjeta_valor
        c_tarjeta.number_format = '@' # Forzar formato texto
        
        # Resto de columnas mapeadas por posición
        ws_destino.cell(row=fila_destino, column=6).value = df_origen.iloc[index, 11]
        ws_destino.cell(row=fila_destino, column=7).value = df_origen.iloc[index, 8]
        ws_destino.cell(row=fila_destino, column=8).value = df_origen.iloc[index, 12]
        ws_destino.cell(row=fila_destino, column=9).value = df_origen.iloc[index, 13]

        # Fechas y Horas
        if val_fecha:
            # Col 12 (L): Fecha
            c_f = ws_destino.cell(row=fila_destino, column=12)
            c_f.value = val_fecha
            c_f.number_format = 'dd/mm/yyyy'
            
            # Col 13 (M): Hora
            c_h = ws_destino.cell(row=fila_destino, column=13)
            c_h.value = val_hora
            c_h.number_format = 'hh:mm:ss'

        fila_destino += 1

    texto_estado.success("✅ ¡Procesamiento completado con éxito!")
    barra.empty() # Limpiar barra al finalizar
    
    # Guardar en memoria
    output = io.BytesIO()
    wb_destino.save(output)
    output.seek(0)
    return output

# --- 5. INTERFAZ DE CARGA ---
col1, col2 = st.columns(2)
with col1:
    uploaded_plantilla = st.file_uploader("📂 1. Plantilla Novatrans", type=["xlsx"])
with col2:
    uploaded_datos = st.file_uploader("📄 2. Excel Moeve/Cepsa", type=["xlsx"])

st.divider()

if uploaded_plantilla and uploaded_datos:
    # Botón principal
    if st.button("🚀 Generar Archivo Importación", type="primary", use_container_width=True):
        with st.spinner('⏳ Procesando datos...'):
            archivo_final = procesar_archivos(uploaded_plantilla, uploaded_datos)
            
            if archivo_final:
                st.balloons()
                st.download_button(
                    label="⬇️ Descargar Archivo Resultado",
                    data=archivo_final,
                    file_name=f"Importacion_Novatrans_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
else:
    st.warning("⚠️ Sube ambos archivos para habilitar el generador.")

st.markdown(
    """<div class='footer'>
        Herramienta interna para Tránsitos y Transportes Logísticos | v1.1 Refined
    </div>""", 
    unsafe_allow_html=True
)
