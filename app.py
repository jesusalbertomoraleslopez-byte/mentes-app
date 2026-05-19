import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# Configuración de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

URL_LOGO = "https://mentesconalas.org.mx"

def cargar_menus_y_datos():
    try:
        # CONEXIÓN DIRECTA A LA NUBE DE GOOGLE SHEETS
        conn = st.connection("gsheets", type=GSheetsConnection)
        
        # Saltamos la fila 1 vacía del Excel para tomar los encabezados reales de la Fila 2
        df = conn.read(ttl=0, header=1) 
        
        # Limpieza de celdas y columnas fantasmas del documento
        df = df.dropna(how='all', axis=1)
        df = df.dropna(subset=[df.columns[0], df.columns[1]])
        
        # Extraer los talleres reales únicos de la columna B (TALLER)
        talleres = sorted(df.iloc[:, 1].dropna().astype(str).str.strip().unique())
        
        # LISTA DE INTEGRANTES MANUAL (Para evitar errores de lectura al no existir la columna en esta pestaña)
        integrantes = [
            "JULIA MARISOL GARCÍA ALCARAZ",
            "ISAAC IGNACIO GONZÁLEZ CRUZ",
            "ANTONIO DE JESÚS RAMÍREZ",
            "MARÍA FERNANDA SÁNCHEZ",
            "CARLOS ALBERTO LÓPEZ"
        ]
        return integrantes, talleres, df
    except Exception as e:
        st.error(f"❌ Error al mapear las columnas del Google Sheet: {e}")
        st.stop()

# Cargar datos en tiempo real
lista_integrantes, lista_talleres, df_original = cargar_menus_y_datos()

# Mapear los nombres de columna reales detectados en tu Sheets
col_fecha = df_original.columns[0]
col_taller = df_original.columns[1]

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
            try:
                conn = st.connection("gsheets", type=GSheetsConnection)
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                nueva_fila_taller = {
                    col_fecha: fecha_hoy,
                    col_taller: nuevo_taller_input
                }
                df_actualizado_taller = pd.concat([df_original, pd.DataFrame([nueva_fila_taller])], ignore_index=True)
                conn.update(data=df_actualizado_taller)
                st.success(f"✨ ¡Taller '{nuevo_taller_input}' guardado en la nube!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al guardar taller: {e}")

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
        st.warning("⚠️ Selecciona integrantes.")
    else:
        try:
            conn = st.connection("gsheets", type=GSheetsConnection)
            fecha_formateada = fecha.strftime("%d/%m/%Y")
            nuevos_registros = []
            
            for integrante in presentes:
                nueva_fila = {
                    col_fecha: fecha_formateada,
                    col_taller: taller
                }
                nuevos_registros.append(nueva_fila)
                
            df_nuevos = pd.DataFrame(nuevos_registros)
            df_actualizado = pd.concat([df_original, df_nuevos], ignore_index=True)
            
            conn.update(data=df_actualizado)
            st.success(f"🎉 ¡Éxito! Se guardaron {len(presentes)} registros en internet.")
            st.session_state.asistencia_estados = {nombre: True for nombre in lista_integrantes}
            st.rerun()
        except Exception as error:
            st.error(f"❌ Error al guardar asistencia: {error}")

st.markdown("---")
st.markdown("### 📋 Historial de Asistencias (Últimos registros)")
st.dataframe(df_original.tail(15), use_container_width=True)

st.markdown("<br><hr><p style='text-align: center;'><a href='https://mentesconalas.org.mx' target='_blank'>🌐 Visitar sitio web oficial - Mentes con Alas</a></p>", unsafe_allow_html=True)
