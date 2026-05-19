import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# Configuración de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

# Logotipo oficial de la fundación alojado en la web para que cargue en cualquier celular
URL_LOGO = "https://mentesconalas.org.mx"

def cargar_menus_y_datos():
    try:
        # CONEXIÓN NATIVA DE LA NUBE: Toma el enlace automáticamente desde los Secrets
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        
        # Limpiar y obtener valores únicos para los menús desplegables
        # Columna 1: Integrantes / Columna 2: Talleres
        integrantes = sorted(df.iloc[:, 1].dropna().astype(str).str.strip().unique())
        talleres = sorted(df.iloc[:, 2].dropna().astype(str).str.strip().unique())
        return integrantes, talleres, df
    except Exception as e:
        st.error(f"❌ Error al leer el archivo en internet: {e}")
        st.stop()

# Cargar la base de datos actual desde internet
lista_integrantes, lista_talleres, df_original = cargar_menus_y_datos()

# Mapear tus 4 columnas exactas por su orden de posición en el archivo
col_fecha = df_original.columns[0]
col_asistencia = df_original.columns[1]
col_taller = df_original.columns[2]
col_horas = df_original.columns[3]

# --- ENCABEZADO VISUAL INSTITUCIONAL ---
col_logo_1, col_logo_2, col_logo_3 = st.columns(3)
with col_logo_2:
    st.image(URL_LOGO, use_container_width=True)

st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin-top: 0px;'>Control de Asistencia Grupal</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6B7280;'>Administración y Registro de Participación en Talleres</p>", unsafe_allow_html=True)
st.markdown("---")

# --- SECCIÓN A: AGREGAR UN TALLER NUEVO ---
with st.expander("➕ ¿Deseas agregar un TALLER NUEVO a la lista?", expanded=False):
    st.write("Escribe el nombre del taller. Al guardarlo, se incluirá en el menú automáticamente.")
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
                    col_asistencia: "ALTA DE TALLER SISTEMA", 
                    col_taller: nuevo_taller_input, 
                    col_horas: 0.00
                }
                df_actualizado_taller = pd.concat([df_original, pd.DataFrame([nueva_fila_taller])], ignore_index=True)
                
                # Sincronizar nuevo taller con la nube
                conn.update(data=df_actualizado_taller)
                st.success(f"✨ ¡Taller '{nuevo_taller_input}' guardado en internet con éxito!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al añadir el taller en la nube: {e}")

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
    
    # Inicializar el estado de los checkboxes si no existe
    if 'asistencia_estados' not in st.session_state:
        st.session_state.asistencia_estados = {nombre: True for nombre in lista_integrantes}
        
    integrantes_filtrados = [n for n in lista_integrantes if buscar_nombre in n] if buscar_nombre else lista_integrantes

    # Mostrar la lista distribuida de forma estética en dos columnas
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

# Procesar los datos al enviar el formulario
if boton_guardar:
    # Filtrar solo a los integrantes que mantuvieron activa su casilla de verificación
    presentes = [nombre for nombre, asistio in list(st.session_state.asistencia_estados.items()) if asistio]
    
    if not presentes:
        st.warning("⚠️ Debes seleccionar al menos a un integrante para procesar la lista.")
    else:
        fecha_formateada = fecha.strftime("%d/%m/%Y")
        nuevos_registros = []
        
        for integrante in presentes:
            nueva_fila = {
                col_fecha: fecha_formateada, 
                col_asistencia: integrante, 
                col_taller: taller, 
                col_horas: float(horas)
            }
            nuevos_registros.append(nueva_fila)
            
        try:
            conn = st.connection("gsheets", type=GSheetsConnection)
            df_nuevos = pd.DataFrame(nuevos_registros)
            df_actualizado = pd.concat([df_original, df_nuevos], ignore_index=True)
            
            # Subir de forma inalámbrica los registros de asistencia en bloque
            conn.update(data=df_actualizado)
            st.success(f"🎉 ¡Éxito! Se guardaron {len(presentes)} registros en internet de forma automática.")
            
            # Resetear el formulario limpio para la siguiente actividad
            st.session_state.asistencia_estados = {nombre: True for nombre in lista_integrantes}
            st.rerun()
        except Exception as error:
            st.error(f"❌ Error al guardar registros en Google Sheets: {error}")

# --- SECCIÓN C: HISTORIAL VISUAL EN TIEMPO REAL ---
st.markdown("---")
st.markdown("### 📋 Últimos Registros Guardados (Historial)")
try:
    if not df_original.empty:
        st.dataframe(df_original.tail(15).iloc[::-1], use_container_width=True)
    else:
        st.info("💡 Aún no hay registros en la base de datos.")
except Exception:
    st.caption("Conectando con el historial de la nube...")

# Pie de página de la institución
st.markdown("<br><hr><p style='text-align: center;'><a href='https://mentesconalas.org.mx' target='_blank'>🌐 Visitar sitio web oficial - Mentes con Alas</a></p>", unsafe_allow_html=True)
