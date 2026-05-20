import streamlit as st
import pandas as pd
from datetime import datetime
from github import Github
import io

# Configuración de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

EXCEL_FILE = "asistencia.xlsx"
URL_LOGO = "https://mentesconalas.org.mx"

# Conexión segura con GitHub usando Secrets
def conectar_github():
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(st.secrets["GITHUB_REPO"])
        return repo
    except Exception as e:
        st.error(f"❌ Error de autenticación con GitHub: {e}")
        st.stop()

def cargar_menus_y_datos():
    repo = conectar_github()
    try:
        # Intentar leer el archivo desde GitHub
        file_content = repo.get_contents(EXCEL_FILE)
        df = pd.read_excel(io.BytesIO(file_content.decoded_content))
        sha = file_content.sha
    except Exception:
        # Si el archivo no existe, lo inicializa con las 4 columnas exactas
        df = pd.DataFrame(columns=["FECHA", "INTEGRANTE / TALLER", "ACTIVIDAD", "HORAS"])
        sha = None

    # Limpiar y obtener valores únicos para los menús desplegables
    integrantes = sorted(df.iloc[:, 1].dropna().astype(str).str.strip().unique())
    talleres = sorted(df.iloc[:, 2].dropna().astype(str).str.strip().unique())
    
    if "ALTA DE TALLER SISTEMA" in integrantes:
        integrantes.remove("ALTA DE TALLER SISTEMA")
        
    return integrantes, talleres, df, sha

def guardar_en_github(df_nuevo, sha_actual, mensaje_commit):
    repo = conectar_github()
    try:
        # Convertir el DataFrame de Pandas a bytes de Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_nuevo.to_excel(writer, index=False)
        excel_bytes = output.getvalue()

        if sha_actual:
            # Actualizar archivo existente
            repo.update_file(EXCEL_FILE, mensaje_commit, excel_bytes, sha_actual)
        else:
            # Crear archivo si no existía
            repo.create_file(EXCEL_FILE, mensaje_commit, excel_bytes)
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar en GitHub: {e}")
        return False

# Cargar datos iniciales
lista_integrantes, lista_talleres, df_original, archivo_sha = cargar_menus_y_datos()

# Mapear columnas de forma dinámica
col_fecha = df_original.columns[0] if len(df_original.columns) > 0 else "FECHA"
col_asistencia = df_original.columns[1] if len(df_original.columns) > 1 else "INTEGRANTE / TALLER"
col_taller = df_original.columns[2] if len(df_original.columns) > 2 else "ACTIVIDAD"
col_horas = df_original.columns[3] if len(df_original.columns) > 3 else "HORAS"

# --- ENCABEZADO VISUAL INSTITUCIONAL ---
col_logo_1, col_logo_2, col_logo_3 = st.columns(3)
with col_logo_2:
    st.image(URL_LOGO, use_container_width=True)

st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin-top: 0px;'>Control de Asistencia Grupal</h2>", unsafe_allow_html=True)
st.markdown("---")

# --- SECCIÓN A: AGREGAR UN TALLER NUEVO ---
with st.expander("➕ ¿Deseas agregar un TALLER NUEVO a la lista?", expanded=False):
    nuevo_taller_input = st.text_input("Nombre del nuevo taller:").strip().upper()
    boton_nuevo_taller = st.button("Guardar Taller Nuevo")
    
    if boton_nuevo_taller:
        if not nuevo_taller_input:
            st.warning("⚠️ El nombre del taller no puede estar vacío.")
        elif nuevo_taller_input in lista_talleres:
            st.info(f"💡 El taller '{nuevo_taller_input}' ya existe.")
        else:
            fecha_hoy = datetime.now().strftime("%d/%m/%Y")
            nueva_fila_taller = {
                col_fecha: fecha_hoy, 
                col_asistencia: "ALTA DE TALLER SISTEMA", 
                col_taller: nuevo_taller_input, 
                col_horas: 0.00
            }
            df_actualizado_taller = pd.concat([df_original, pd.DataFrame([nueva_fila_taller])], ignore_index=True)
            
            if guardar_en_github(df_actualizado_taller, archivo_sha, f"Sistema: Alta de taller {nuevo_taller_input}"):
                st.success(f"✨ ¡Taller '{nuevo_taller_input}' guardado en GitHub!")
                st.rerun()

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
            
        df_nuevos = pd.DataFrame(nuevos_registros)
        df_actualizado = pd.concat([df_original, df_nuevos], ignore_index=True)
        
        if guardar_en_github(df_actualizado, archivo_sha, f"App: Registro asistencia - {taller}"):
            st.success(f"🎉 ¡Éxito! Se guardaron los cambios directamente en tu repositorio de GitHub.")
            st.session_state.asistencia_estados = {nombre: True for nombre in lista_integrantes}
            st.rerun()

# --- SECCIÓN C: HISTORIAL VISUAL EN TIEMPO REAL ---
st.markdown("---")
st.markdown("### 📋 Últimos Registros Guardados (Historial)")
if not df_original.empty:
    st.dataframe(df_original.tail(15).iloc[::-1], use_container_width=True)
else:
    st.info("💡 Aún no hay registros en la base de datos. El archivo Excel se creará automáticamente en tu GitHub al guardar tu primer taller o asistencia.")

st.markdown("<br><hr><p style='text-align: center;'><a href='https://mentesconalas.org.mx' target='_blank'>🌐 Visitar sitio web oficial</a></p>", unsafe_allow_html=True)
