import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Configuración de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

# --- CONEXIÓN DE ESCRITORIO LOCAL ---
RUTA_EXCEL = r"G:\Mi unidad\MENTES CON ALAS\ENTREGABLES KPI-LUPITA.xlsx"
NOMBRE_HOJA = "PARTICIPACION DE INTEGRANTES"
RUTA_LOGO = r"G:\Mi unidad\MENTES CON ALAS\logo-mentes.png"

def cargar_menus_y_datos():
    if not os.path.exists(RUTA_EXCEL):
        st.error(f"⚠️ No se encontró el archivo en la ruta: {RUTA_EXCEL}. Verifica que Google Drive de escritorio esté abierto.")
        st.stop()
    try:
        df = pd.read_excel(RUTA_EXCEL, sheet_name=NOMBRE_HOJA)
        integrantes = sorted(df.iloc[:, 1].dropna().astype(str).str.strip().unique())
        talleres = sorted(df.iloc[:, 2].dropna().astype(str).str.strip().unique())
        return integrantes, talleres, df
    except Exception as e:
        st.error(f"❌ Error al leer el archivo Excel: {e}")
        st.stop()

# Cargar la base de datos local
lista_integrantes, lista_talleres, df_original = cargar_menus_y_datos()

# Mapear los nombres de columnas
col_fecha = df_original.columns
col_asistencia = df_original.columns
col_taller = df_original.columns
col_horas = df_original.columns

# --- DISEÑO DEL ENCABEZADO ---
col_logo_1, col_logo_2, col_logo_3 = st.columns(3)
with col_logo_2:
    if os.path.exists(RUTA_LOGO):
        st.image(RUTA_LOGO, use_container_width=True)
    else:
        st.title("🦅 Mentes Con Alas")

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
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                nueva_fila_taller = {col_fecha: fecha_hoy, col_asistencia: "ALTA DE TALLER SISTEMA", col_taller: nuevo_taller_input, col_horas: 0.00}
                df_actualizado_taller = pd.concat([df_original, pd.DataFrame([nueva_fila_taller])], ignore_index=True)
                with pd.ExcelWriter(RUTA_EXCEL, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df_actualizado_taller.to_excel(writer, sheet_name=NOMBRE_HOJA, index=False)
                st.success(f"✨ ¡Taller '{nuevo_taller_input}' guardado con éxito!")
                st.rerun()
            except PermissionError:
                st.error("❌ Error de acceso: Cierra tu archivo de Excel en la computadora antes de añadir el taller.")

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
            df_nuevos = pd.DataFrame(nuevos_registros)
            df_actualizado = pd.concat([df_original, df_nuevos], ignore_index=True)
            with pd.ExcelWriter(RUTA_EXCEL, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                df_actualizado.to_excel(writer, sheet_name=NOMBRE_HOJA, index=False)
            st.success(f"🎉 ¡Éxito! Se guardaron {len(presentes)} registros en tu Excel local.")
            st.session_state.asistencia_estados = {nombre: True for nombre in lista_integrantes}
            st.rerun()
        except PermissionError:
            st.error("❌ Cierra tu archivo Excel 'ENTREGABLES KPI-LUPITA.xlsx' antes de guardar.")
        except Exception as error:
            st.error(f"❌ Error: {error}")

st.markdown("---")
st.markdown("### 📋 Últimos Registros Guardados")
try:
    df_vista = pd.read_excel(RUTA_EXCEL, sheet_name=NOMBRE_HOJA)
    if not df_vista.empty:
        st.dataframe(df_vista.tail(15).iloc[::-1], use_container_width=True)
except Exception:
    st.caption("Conectando con el historial...")

st.markdown("<br><hr><p style='text-align: center;'><a href='https://mentesconalas.org.mx' target='_blank'>🌐 Visitar sitio web oficial - Mentes con Alas</a></p>", unsafe_allow_html=True)
