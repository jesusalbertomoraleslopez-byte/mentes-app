import streamlit as st
import pandas as pd
from datetime import datetime
from github import Github
import io

# Configuración estética de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

EXCEL_FILE = "asistencia.xlsx"
CALENDARIO_FILE = "calendario.xlsx"  # Nuevo archivo independiente para la planeación
URL_LOGO = "logo-mentes.png"

# --- INYECTAR DISEÑO VISUAL INSPIRADO EN LA WEB OFICIAL ---
st.markdown("""
    <style>
        html, body, [data-testid="stAppViewContainer"] {
            background-color: #FAFAFA;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        .titulo-web {
            text-align: center; 
            color: #0A2540; 
            font-size: 32px; 
            font-weight: 700;
            margin-bottom: 5px;
        }
        .subtitulo-web {
            text-align: center; 
            color: #627D98; 
            font-size: 16px; 
            margin-bottom: 30px;
        }
        h3 {
            color: #0A2540 !important;
            font-weight: 600 !important;
        }
        div[data-testid="stExpander"] {
            border: 1px solid #E4E7EB !important;
            border-radius: 8px !important;
            background-color: #FFFFFF !important;
            box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.02) !important;
        }
        div[data-testid="stCheckbox"] label p, div[data-testid="stExpander"] div[data-testid="stCheckbox"] label p {
            color: #0A2540 !important;
            font-weight: 700 !important;
            font-size: 15px !important;
        }
        div[data-testid="stExpander"] p, div[data-testid="stExpander"] span {
            color: #0A2540 !important;
            font-weight: 500 !important;
        }
        div[data-testid="stNotification"] {
            background-color: #FFEEEE !important;
            border: 2px solid #E53E3E !important;
            border-radius: 8px !important;
        }
        div[data-testid="stNotification"] p, div[data-testid="stNotification"] span, div[data-testid="stNotification"] li {
            color: #9B1C1C !important;
            font-weight: 700 !important;
            font-size: 16px !important;
        }
        .contenedor-asistencia {
            background-color: #FFFFFF !important;
            padding: 20px !important;
            border-radius: 8px !important;
            border: 1px solid #E4E7EB !important;
            margin-top: 15px !important;
            margin-bottom: 15px !important;
        }
        .stButton > button, div[data-testid="stForm"] button {
            background-color: #0A2540 !important;
            color: white !important;
            border-radius: 6px !important;
            border: none !important;
            padding: 10px 24px !important;
            font-weight: 600 !important;
            transition: all 0.3s ease !important;
        }
        .stButton > button:hover, div[data-testid="stForm"] button:hover {
            background-color: #1F3B5C !important;
            box-shadow: 0px 4px 10px rgba(10, 37, 64, 0.15) !important;
            transform: translateY(-1px);
        }
        input, select, div[data-baseweb="select"] {
            border-radius: 6px !important;
        }
        .footer-web {
            text-align: center;
            font-size: 13px;
            color: #829AB1;
            margin-top: 50px;
            line-height: 1.6;
        }
        .footer-web a {
            color: #0A2540 !important;
            text-decoration: none;
            font-weight: bold;
        }
        .boton-borrado-html {
            background-color: #FFFFFF !important;
            color: #D32F2F !important;
            border: 3px solid #D32F2F !important;
            padding: 14px 20px !important;
            text-align: center !important;
            font-size: 16px !important;
            font-weight: 800 !important;
            margin: 4px 2px !important;
            cursor: pointer !important;
            border-radius: 8px !important;
            width: 100% !important;
            text-transform: uppercase !important;
            box-shadow: 0px 4px 6px rgba(0,0,0,0.1) !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- LISTA OFICIAL DE INTEGRANTES ---
INTEGRANTES_FIJOS = [
    "ANA DE LOS ÁNGELES TORRES REVELES", "CARLOS ALBERTO DE LA TORRE SANTELLANO",
    "CRISTABEL DE LA CRUZ MALDONADO", "EFRAÍN MAYNEZ PORRAS", "FERNANDO ÁVILA BERLANGA",
    "FLORINDA ESTEVANÉ PIZARRO", "GUADALUPE LÓPEZ TOVAR", "ISAAC IGNACIO GONZÁLEZ CRUZ",
    "JESÚS ALEJANDRO SIFUENTES ESPINO", "JESÚS SALCIDO ZAMORA", "JOSE REYNALDO ALCORTA BENAVIDES",
    "JUAN JOSÉ OCHOA GONZÁLEZ", "JUAN RAFAEL YAÑEZ SERNA", "JUAN SILVERIO LÓPEZ DE LA ROSA",
    "KARLA LISSETTE PEDROZA GONZÁLEZ", "LUIS FERNANDO ÁVILA BERLANGA", "LUIS JAVIER GARCÍA MARTÍNEZ",
    "MA. ELIZABETH GONZÁLEZ HIDROGO", "MARÍA DE LOS ÁNGELES ORDUÑA RODRÍGUEZ",
    "MARÍA GUADALUPE DE LA CONCEPCIÓN TORRES LIRA", "PEDRO ANTONIO DE LA CERDA TREVIÑO",
    "TERESITA DEL NIÑO JESÚS RODRÍGUEZ SALAZAR", "TOMASITA MARÍA ENRIQUETA RIVERA DEL BOSQUE"
]

# --- LISTA OFICIAL DE TALLERES ---
TALLERES_FIJOS = [
    "AJEDREZ", "ARTE", "DESARROLLO EDUCATIVO", "FELDENKRAIS", "GRUPO DE CRECIMIENTO",
    "MOVIMIENTO VITAL EXPRESIVO", "TALLER \"SÍ PUEDO\"", "TALLER DE COMUNICACIÓN", "TEATRO", "VIDA DIARIA"
]

# Columnas exactas del archivo original (Se quedan idénticas)
COL_FECHA = "FECHA"
COL_ASISTENCIA = "INTEGRANTE / TALLER"
COL_TALLER = "ACTIVIDAD"
COL_HORAS = "HORAS"
CLAVE_BORRADO = "LupitaMentes1978"

def conectar_github():
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(st.secrets["GITHUB_REPO"])
        return repo
    except Exception as e:
        st.error(f"❌ Error de autenticación con GitHub: {e}")
        st.stop()

def cargar_datos_sistema():
    repo = conectar_github()
    
    # 1. Cargar archivo de Asistencia Histórica
    try:
        file_asistencia = repo.get_contents(EXCEL_FILE)
        df_asistencia = pd.read_excel(io.BytesIO(file_asistencia.decoded_content))
        sha_asistencia = file_asistencia.sha
    except Exception:
        df_asistencia = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS])
        sha_asistencia = None

    # 2. Cargar archivo de Calendario Independiente
    try:
        file_calendario = repo.get_contents(CALENDARIO_FILE)
        df_calendario = pd.read_excel(io.BytesIO(file_calendario.decoded_content))
        sha_calendario = file_calendario.sha
    except Exception:
        df_calendario = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS])
        sha_calendario = None

    return df_asistencia, sha_asistencia, df_calendario, sha_calendario

# Cargar ambos dataframes de manera aislada
df_original, archivo_sha, df_calendario, sha_calendario = cargar_datos_sistema()

def guardar_archivo_github(nombre_archivo, df_nuevo, sha_actual, mensaje_commit):
    repo = conectar_github()
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_nuevo.to_excel(writer, index=False)
        excel_bytes = output.getvalue()
        if sha_actual:
            repo.update_file(nombre_archivo, mensaje_commit, excel_bytes, sha_actual)
        else:
            repo.create_file(nombre_archivo, mensaje_commit, excel_bytes)
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar {nombre_archivo} en GitHub: {e}")
        return False

# --- ENCABEZADO VISUAL INSTITUCIONAL ---
col_logo_1, col_logo_2, col_logo_3 = st.columns([1, 1.2, 1])
with col_logo_2:
    st.image(URL_LOGO, use_container_width=True)

st.markdown('<div class="titulo-web">Control de Asistencia Grupal</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo-web">Comunidad de Vida para Adultos con Parálisis Cerebral</div>', unsafe_allow_html=True)

# --- SECCIÓN A: CALENDARIO DE ACTIVIDADES POR DÍA (NUEVO ARCHIVO SEPARADO) ---
with st.expander("📅 1. PROGRAMAR ACTIVIDADES EN EL CALENDARIO (NUEVO ARCHIVO)", expanded=False):
    st.write("Planifica qué alumnos asistirán a qué talleres en fechas específicas. Esto guardará datos en `calendario.xlsx`.")
    with st.form("form_calendario_independiente"):
        prog_fecha = st.date_input("FECHA PROGRAMADA", datetime.now())
        prog_taller = st.selectbox("TALLER A IMPARTIR", TALLERES_FIJOS)
        prog_horas = st.number_input("HORAS ESTIMADAS", min_value=0.25, max_value=8.0, value=1.0, step=0.25)
        
        st.markdown("### 👥 Selecciona los Integrantes Programados:")
        st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
        col_c_izq, col_c_der = st.columns(2)
        agenda_estados = {}
        for i, nombre in enumerate(INTEGRANTES_FIJOS):
            if i % 2 == 0:
                with col_c_izq:
                    agenda_estados[nombre] = st.checkbox(nombre, value=False, key=f"cal_{nombre}")
            else:
                with col_c_der:
                    agenda_estados[nombre] = st.checkbox(nombre, value=False, key=f"cal_{nombre}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        boton_guardar_calendario = st.form_submit_button("🗓️ Guardar Programación de Actividad")

    if boton_guardar_calendario:
        seleccionados = [nombre for nombre, check in agenda_estados.items() if check]
        if not seleccionados:
            st.warning("⚠️ Selecciona al menos a un participante para la planeación.")
        else:
            fecha_prog_str = prog_fecha.strftime("%d/%m/%Y")
            nuevas_filas_cal = []
            for integrante in seleccionados:
                nuevas_filas_cal.append({
                    COL_FECHA: fecha_prog_str,
                    COL_ASISTENCIA: integrante,
                    COL_TALLER: prog_taller,
                    COL_HORAS: float(prog_horas)
                })
            df_nuevo_cal = pd.concat([df_calendario, pd.DataFrame(nuevas_filas_cal)], ignore_index=True)
            if guardar_archivo_github(CALENDARIO_FILE, df_nuevo_cal, sha_calendario, f"Calendario: Se agendó {prog_taller} para {fecha_prog_str}"):
                st.success(f"✨ ¡Agenda guardada con éxito en {CALENDARIO_FILE} para {len(seleccionados)} integrantes!")
                st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

# --- SECCIÓN B: FORMULARIO DE ASISTENCIA REAL (BASADO EN EL CALENDARIO) ---
st.markdown("### 📝 2. PASE DE LISTA BASADO EN LO PROGRAMADO")
fecha_lista_buscar = st.date_input("Selecciona la fecha a evaluar:", datetime.now()).strftime("%d/%m/%Y")

# Buscar si existen alumnos citados en el archivo calendario para esta fecha
df_citados_hoy = df_calendario[df_calendario[COL_FECHA] == fecha_lista_buscar]

if df_citados_hoy.empty:
    st.info(f"💡 No hay ninguna actividad agendada en el calendario para el día {fecha_lista_buscar}. Por favor programa una en la sección superior.")
else:
    # Obtener dinámicamente el taller y las horas que se habían planeado
    taller_programado = df_citados_hoy[COL_TALLER].iloc[0]
    horas_programadas = df_citados_hoy[COL_HORAS].iloc[0]
    st.success(f"📚 Taller Citado en Calendario: **{taller_programado}** ({horas_programadas} horas)")
    
    with st.form("formulario_asistencia_real"):
        st.write("📋 **Lista de asistencia inteligente:** Las casillas ya están marcadas con los que citaste. **Desmarca a las personas que hayan faltado**.")
        
        st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
        col_r_izq, col_r_der = st.columns(2)
        
        pase_lista_checks = {}
        # Mostrar EN EXCLUSIVA a los que se registraron en el calendario para este día
        for i, (idx, fila) in enumerate(df_citados_hoy.iterrows()):
            nombre_alumno = fila[COL_ASISTENCIA]
            if i % 2 == 0:
                with col_r_izq:
                    # Aparecen checkeados por defecto. Si el usuario desmarca, se asume Falta (no se guarda).
                    pase_lista_checks[nombre_alumno] = st.checkbox(nombre_alumno, value=True, key=f"real_{idx}")
            else:
                with col_r_der:
                    pase_lista_checks[nombre_alumno] = st.checkbox(nombre_alumno, value=True, key=f"real_{idx}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        boton_guardar_asistencia_final = st.form_submit_button("💾 Guardar y Registrar Asistencia en el Historial")

    if boton_guardar_asistencia_final:
        # Filtrar solo los nombres que se quedaron con la casilla verificada (Asistieron)
        asistentes_reales = [nombre for nombre, asistio in pase_lista_checks.items() if asistio]
        
        if not asistentes_reales:
            st.warning("⚠️ Al desmarcar a todos, estás indicando que nadie asistió. Debes procesar al menos una asistencia.")
        else:
            nuevos_registros_asistencia = []
            for integrante in asistentes_reales:
                nuevos_registros_asistencia.append({
                    COL_FECHA: fecha_lista_buscar,
                    COL_ASISTENCIA: integrante,
                    COL_TALLER: taller_programado,
                    COL_HORAS: float(horas_programadas)
                })
            
            # Guardar ÚNICAMENTE los que asistieron en el archivo de asistencia tradicional
            df_final_asistencia = pd.concat([df_original, pd.DataFrame(nuevos_registros_asistencia)], ignore_index=True)
            if guardar_archivo_github(EXCEL_FILE, df_final_asistencia, archivo_sha, f"Asistencia: Registro real de {taller_programado} - {fecha_lista_buscar}"):
                st.success(f"🎉 ¡Éxito! Se han guardado de forma permanente {len(asistentes_reales)} registros de asistencia en {EXCEL_FILE}.")
                st.rerun()

# --- SECCIÓN C: HISTORIAL VISUAL EN TIEMPO REAL CON FILTRADO ---
st.markdown("---")
st.markdown("### 📋 Historial y Buscador de Asistencias")

if not df_original.empty:
    df_visual = df_original.copy()
    try:
        df_visual['FECHA_DT'] = pd.to_datetime(df_visual[COL_FECHA], format="%d/%m/%Y", errors='coerce')
        df_visual = df_visual.sort_values(by='FECHA_DT', ascending=False).drop(columns=['FECHA_DT'])
    except Exception:
        pass

    filtro_busqueda = st.text_input("🔍 Escribe una FECHA, NOMBRE o ACTIVIDAD para filtrar y descargar:").strip().upper()
    
    if filtro_busqueda:
        mask = (
            df_visual[COL_FECHA].astype(str).str.upper().str.contains(filtro_busqueda) |
            df_visual[COL_ASISTENCIA].astype(str).str.upper().str.contains(filtro_busqueda) |
            df_visual[COL_TALLER].astype(str).str.upper().str.contains(filtro_busqueda)
        )
        df_filtrado_tabla = df_visual[mask]
    else:
        df_filtrado_tabla = df_visual

    st.dataframe(df_filtrado_tabla, use_container_width=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    output_descarga = io.BytesIO()
    with pd.ExcelWriter(output_descarga, engine='openpyxl') as writer:
        df_visual.sort_values(by=COL_FECHA, ascending=True).to_excel(writer, index=False)
    excel_completo_bytes = output_descarga.getvalue()
    
    st.download_button(
        label="📥 Descargar Toda la Base de Datos Histórica (Completa)",
        data=excel_completo_bytes,
        file_name=f"asistencia_completa_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="btn_descarga_masiva_limpia"
    )
else:
    st.info("💡 Aún no hay registros guardados en el archivo Excel de Asistencia.")

# --- SECCIÓN D: ADMINISTRACIÓN DE SEGURIDAD (BORRAR REGISTROS / TALLERES) ---
st.markdown("<br>", unsafe_allow_html=True)
with st.expander("🚨 Panel de Administración - Control de Historial", expanded=False):
    st.write("Módulo de alta seguridad exclusivo para la remoción precisa de información errónea en la base de datos de asistencia.")
    clave_input = st.text_input("Ingresa la clave de autorización:", type="password", key="clave_borrar")
    
    if clave_input == CLAVE_BORRADO:
        st.warning("⚠️ Clave correcta. Configura los filtros de abajo para buscar y seleccionar los registros a eliminar.")
        
        if not df_original.empty:
            fechas_unicas = sorted(df_original[COL_FECHA].dropna().unique(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"), reverse=True)
            fecha_seleccionada = st.selectbox("1. Selecciona la fecha donde se encuentra el error:", fechas_unicas, key="fecha_eliminar_individual")
            df_dia = df_original[df_original[COL_FECHA] == fecha_seleccionada].copy()
            
            if not df_dia.empty:
                st.markdown("### 2. Selecciona las casillas de las personas que deseas BORRAR:")
                registros_a_eliminar = []
                
                st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
                for idx, fila in df_dia.iterrows():
                    info_registro = f"👤 {fila[COL_ASISTENCIA]} | 📚 {fila[COL_TALLER]} ({fila[COL_HORAS]} hrs)"
                    marcado = st.checkbox(info_registro, value=False, key=f"del_{idx}")
                    if marcado:
                        registros_a_eliminar.append(idx)
                st.markdown('</div>', unsafe_allow_html=True)
                
                if registros_a_eliminar:
                    st.error(f"🚨 Alerta de seguridad: Has marcado {len(registros_a_eliminar)} registro(s) para ser eliminado(s).")
                    confirmar_clic_html = st.checkbox("👉 Marca esta casilla para activar el botón de guardado permanente", value=False, key="check_activar_html_personas")
                    st.markdown('<button class="boton-borrado-html">❌ CONFIRMAR ELIMINACIÓN DE REGISTROS</button>', unsafe_allow_html=True)
                    
                    if confirmar_clic_html:
                        df_resultado = df_original.drop(index=registros_a_eliminar)
                        if guardar_archivo_github(EXCEL_FILE, df_resultado, archivo_sha, f"Admin: Eliminación de {len(registros_a_eliminar)} registros individuales del día {fecha_seleccionada}"):
                            st.success("🎉 ¡Los registros seleccionados han sido removidos con éxito!")
                            st.rerun()
            else:
                st.info("No se encontraron registros.")
        else:
            st.info("💡 La base de datos está vacía.")
            
    elif clave_input != "":
        st.error("❌ Clave incorrecta.")

# --- PIE DE PÁGINA INSTITUCIONAL ---
st.markdown("""
    <div class="footer-web">
        <hr>
        Av. Ocampo 1797 ote. C.P. 27000, Col. Centro Torreón, Coah.<br>
        <b>MENTES CON ALAS ES DONATARIA AUTORIZADA</b><br>
        <a href="https://mentesconalas.org.mx" target="_blank">🌐 Visitar Sitio Web Oficial</a>
    </div>
""", unsafe_allow_html=True)

