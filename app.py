import streamlit as st
import pandas as pd
from datetime import datetime
import os
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

# URL de tu archivo Google Sheets en internet (El que ya configuraste como público)
ENLACE_SHEETS = "https://google.com"
URL_LOGO = "https://mentesconalas.org.mx"

def cargar_menus_y_datos():
    try:
        # CONEXIÓN NATIVA: Lee el archivo desde internet usando las credenciales de la nube
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(spreadsheet=ENLACE_SHEETS, ttl=0)
        
        # Extraer listas limpias
        integrantes = sorted(df.iloc[:, 1].dropna().astype(str).str.strip().unique())
        talleres = sorted(df.iloc[:, 2].dropna().astype(str).str.strip().unique())
        return integrantes, talleres, df
    except Exception as e:
        st.error(f"❌ Error al leer el archivo en internet: {e}")
        st.stop()

lista_integrantes, lista_talleres, df_original = cargar_menus_y_datos()

# Mapear tus 4 columnas exactas por su posición
col_fecha = df_original.columns[0]
col_asistencia = df_original.columns[1]
col_taller = df_original.columns[2]
col_horas = df_original.columns[3]

# --- ENCABEZADO ---
col_logo_1, col_logo_2, col_logo_3 = st.columns(3)
with col_logo_2:
    st.image(URL_LOGO, use_container_width=True)

st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin-top: 0px;'>Control de Asistencia Grupal</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6B7280;'>Administración y Registro de Participación en Talleres</p>", unsafe_allow_html=True)
st.markdown("---")

# --- SECCIÓN A: AGREGAR TALLER NUEVO ---
with st.expander("➕ ¿Deseas agregar un TALLER NUEVO a la lista?", expanded=False):
    st.write("Escribe el nombre del taller. Al guardarlo, se incluirá automáticamente.")
    nuevo_taller_input = st.text_input("Nombre del nuevo taller:").strip().upper()
    boton_nuevo_taller = st.button("Guardar Taller Nuevo")
    
    if boton_nuevo_taller:
        if not nuevo_taller_input:
            st.warning("⚠️ El nombre del taller no puede estar vacío.")
        elif nuevo_taller_input in lista_talleres:
            st.info(f"💡 El taller '{nuevo_taller_input}' ya existe.")
        else:
            try:
                conn = st.connection("gsheets", type=GSheetsConnection)
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                nueva_fila_taller = {col_fecha: fecha_hoy, col_asistencia: "ALTA DE TALLER SISTEMA", col_taller: nuevo_taller_input, col_horas: 0.00}
                df_actualizado_taller = pd.concat([df_original, pd.DataFrame([nueva_fila_taller])], ignore_index=True)
                
                # Guardar el nuevo taller inalámbricamente
                conn.update(spreadsheet=ENLACE_SHEETS, data=df_actualizado_taller)
                st.success(f"✨ ¡Taller '{nuevo_taller_input}' guardado!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al añadir taller: {e}")

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
    if not presentes:
        st.warning("⚠️ Debes seleccionar al menos a un integrante.")
    else:
        fecha_formateada = fecha.strftime("%d/%m/%Y")
        nuevos_registros = []
        for integrante in presentes:
            nueva_fila = {col_fecha: fecha_formateada, col_asistencia: integrante, col_taller: taller, col_horas: float(horas)}
            nuevos_registros.append(nueva_fila)
        try:
            conn = st.connection("gsheets", type=GSheetsConnection)
            df_nuevos = pd.DataFrame(nuevos_registros)
            df_actualizado = pd.concat([df_original, df_nuevos], ignore_index=True)
            
            # Guardar registros de asistencia inalámbricamente
            conn.update(spreadsheet=ENLACE_SHEETS, data=df_actualizado)
            st.success(f"🎉 ¡Éxito! Se guardaron {len(presentes)} registros en internet.")
            st.session_state.asistencia_estados = {nombre: True for nombre in lista_integrantes}
            st.rerun()
        except Exception as error:
            st.error(f"❌ Error al guardar: {error}")

st.markdown("---")
st.markdown("### 📋 Últimos Registros Guardados")
try:
    # Mostrar el historial inalámbrico en tiempo real
    st.dataframe(df_original.tail(15).iloc[::-1], use_container_width=True)
except Exception:
    st.caption("Conectando con el historial...")

st.markdown("<br><hr><p style='text-align: center;'><a href='https://mentesconalas.org.mx' target='_blank'>🌐 Visitar sitio web oficial - Mentes con Alas</a></p>", unsafe_allow_html=True)
