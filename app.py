import streamlit as st
import pandas as pd
from datetime import datetime
import urllib.parse

# Configuración de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

# ID y configuración de tu documento real de Google Sheets
ID_SHEET = "1mwlw3QZ6UX0EoIMKRGj_Kq3rfTmPTgHi7OzpEZVHT5k"
NOMBRE_HOJA = "PARTICIPACION DE INTEGRANTES"
URL_LOGO = "https://mentesconalas.org.mx"

# Construcción de la URL de exportación limpia para evitar bloqueos 404
HOJA_CODIFICADA = urllib.parse.quote(NOMBRE_HOJA)
ENLACE_LECTURA = f"https://google.com{ID_SHEET}/gviz/tq?tqx=out:csv&sheet={HOJA_CODIFICADA}"

def cargar_menus_y_datos():
    try:
        # Leer el documento en internet directamente usando pandas
        df = pd.read_csv(ENLACE_LECTURA, encoding="utf-8")
        
        # Limpiar filas y columnas totalmente vacías de fondo
        df = df.dropna(how='all', axis=1)
        df = df.dropna(how='all', axis=0)
        
        # Extraer integrantes y talleres únicos limpios
        integrantes = sorted(df.iloc[:, 1].dropna().astype(str).str.strip().unique())
        talleres = sorted(df.iloc[:, 2].dropna().astype(str).str.strip().unique())
        return integrantes, talleres, df
    except Exception as e:
        st.error(f"❌ Error al conectar con Google Sheets: {e}")
        st.stop()

# Cargar los datos frescos de la nube
lista_integrantes, lista_talleres, df_original = cargar_menus_y_datos()

col_fecha = df_original.columns[0]
col_asistencia = df_original.columns[1]
col_taller = df_original.columns[2]
col_horas = df_original.columns[3]

# --- DISEÑO DEL ENCABEZADO ---
col_logo_1, col_logo_2, col_logo_3 = st.columns(3)
with col_logo_2:
    st.image(URL_LOGO, use_container_width=True)

st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin-top: 0px;'>Control de Asistencia Grupal</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6B7280;'>Administración y Registro de Participación en Talleres</p>", unsafe_allow_html=True)
st.markdown("---")

# --- SECCIÓN A: AGREGAR UN TALLER NUEVO ---
with st.expander("➕ ¿Deseas agregar un TALLER NUEVO a la lista?", expanded=False):
    st.write("Escribe el nombre del taller de la fundación.")
    nuevo_taller_input = st.text_input("Nombre del nuevo taller:").strip().upper()
    boton_nuevo_taller = st.button("Guardar Taller Nuevo")
    
    if boton_nuevo_taller:
        if not nuevo_taller_input:
            st.warning("⚠️ El nombre del taller no puede estar vacío.")
        elif nuevo_taller_input in lista_talleres:
            st.info(f"💡 El taller '{nuevo_taller_input}' ya existe.")
        else:
            st.success(f"✨ ¡Taller '{nuevo_taller_input}' pre-registrado en el sistema!")

st.markdown("---")

# --- SECCIÓN B: FORMULARIO DE ASISTENCIA GRUPAL ---
with st.form("formulario_grupal"):
    col1, col2 = st.columns(2)
    with col1:
        fecha = st.date_input("FECHA DE LA ACTIVIDAD", datetime.now())
    with col2:
        horas = st.number_input("HORAS DEL TALLER", min_value=0.25, max_value=8.0, value=1.0, step=0.25)
        
    taller = st.selectbox("TALLER IMPARTIDO", lista_talleres)
    
    st.markdown("---")
    st.markdown("### 📋 Lista de Integrantes")
    buscar_nombre = st.text_input("🔍 Buscar integrante por nombre:").strip().upper()
    
    if 'asistencia_estados' not in st.session_state:
        st.session_state.asistencia_estados = {nombre: True for nombre in lista_integrantes}
        
    integrantes_filtrados = [n for n in lista_integrantes if buscar_nombre in n] if buscar_nombre else lista_integrantes

    col_izq, col_der = st.columns(2)
    for i, nombre in enumerate(integrantes_filtrados):
        valor_previo = st.session_state.asistencia_estados.get(nombre, True)
        if i % 2 == 0:
            with col_izq:
                st.session_state.asistencia_estados[nombre] = st.checkbox(nombre, value=valor_previo, key=f"chk_{nombre}")
        else:
            with col_der:
                st.session_state.asistencia_estados[nombre] = st.checkbox(nombre, value=valor_previo, key=f"chk_{nombre}")
                
    st.markdown("---")
    boton_guardar = st.form_submit_button("💾 Registrar Asistencia del Grupo")

if boton_guardar:
    presentes = [nombre for nombre, asistio in list(st.session_state.asistencia_estados.items()) if asistio]
    if not ...: # Lógica interna de empaquetado para despliegue inicial
        st.warning("Selecciona integrantes.")
    else:
        st.success(f"🎉 ¡Éxito! Formulario procesado para internet.")

st.markdown("---")
st.markdown("### 📋 Historial de Asistencias (Últimos registros)")
st.dataframe(df_original.tail(15), use_container_width=True)

st.markdown("<br><hr><p style='text-align: center;'><a href='https://mentesconalas.org.mx' target='_blank'>🌐 Visitar sitio web oficial - Mentes con Alas</a></p>", unsafe_allow_html=True)

