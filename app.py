import streamlit as st
import pandas as pd
from datetime import datetime
from github import Github
import io

# Configuración estética de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

EXCEL_FILE = "asistencia.xlsx"
URL_LOGO = "logo-mentes.png"

# --- INYECTAR DISEÑO VISUAL INSPIRADO EN LA WEB OFICIAL ---
st.markdown("""
    <style>
        /* Tipografía general y fondo limpio como la web */
        html, body, [data-testid="stAppViewContainer"] {
            background-color: #FAFAFA;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        /* Encabezado principal estilo institucional */
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
        /* Estilo para los subtítulos de secciones */
        h3 {
            color: #0A2540 !important;
            font-weight: 600 !important;
        }
        /* Tarjetas o contenedores tipo bloques de la web */
        div[data-testid="stExpander"] {
            border: 1px solid #E4E7EB !important;
            border-radius: 8px !important;
            background-color: #FFFFFF !important;
            box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.02) !important;
        }
        /* FORZAR COLOR AZUL INSTITUCIONAL Y NEGRITA EN LOS MIEMBROS Y EN EL PANEL DE BORRADO */
        div[data-testid="stCheckbox"] label p, div[data-testid="stExpander"] div[data-testid="stCheckbox"] label p {
            color: #0A2540 !important;
            font-weight: 700 !important;
            font-size: 15px !important;
        }
        /* Asegurar que las etiquetas de texto descriptivo en el panel administrativo también sean visibles */
        div[data-testid="stExpander"] p, div[data-testid="stExpander"] span {
            color: #0A2540 !important;
            font-weight: 500 !important;
        }
        /* Contenedor blanco de contraste para la lista de asistencia y borrado */
        .contenedor-asistencia {
            background-color: #FFFFFF !important;
            padding: 20px !important;
            border-radius: 8px !important;
            border: 1px solid #E4E7EB !important;
            margin-top: 15px !important;
            margin-bottom: 15px !important;
        }
        /* Personalización de los botones principales (Azul Mentes con Alas) */
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
        /* Inputs y Selectores limpios */
        input, select, div[data-baseweb="select"] {
            border-radius: 6px !important;
        }
        /* Pie de página institucional */
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
    </style>
""", unsafe_allow_html=True)

# --- LISTA OFICIAL Y FIJA DE INTEGRANTES ---
INTEGRANTES_FIJOS = [
    "ANA DE LOS ÁNGELES TORRES REVELES",
    "CARLOS ALBERTO DE LA TORRE SANTELLANO",
    "CRISTABEL DE LA CRUZ MALDONADO",
    "EFRAÍN MAYNEZ PORRAS",
    "FERNANDO ÁVILA BERLANGA",
    "FLORINDA ESTEVANÉ PIZARRO",
    "GUADALUPE LÓPEZ TOVAR",
    "ISAAC IGNACIO GONZÁLEZ CRUZ",
    "JESÚS ALEJANDRO SIFUENTES ESPINO",
    "JESÚS SALCIDO ZAMORA",
    "JOSE REYNALDO ALCORTA BENAVIDES",
    "JUAN JOSÉ OCHOA GONZÁLEZ",
    "JUAN RAFAEL YAÑEZ SERNA",
    "JUAN SILVERIO LÓPEZ DE LA ROSA",
    "KARLA LISSETTE PEDROZA GONZÁLEZ",
    "LUIS FERNANDO ÁVILA BERLANGA",
    "LUIS JAVIER GARCÍA MARTÍNEZ",
    "MA. ELIZABETH GONZÁLEZ HIDROGO",
    "MARÍA DE LOS ÁNGELES ORDUÑA RODRÍGUEZ",
    "MARÍA GUADALUPE DE LA CONCEPCIÓN TORRES LIRA",
    "PEDRO ANTONIO DE LA CERDA TREVIÑO",
    "TERESITA DEL NIÑO JESÚS RODRÍGUEZ SALAZAR",
    "TOMASITA MARÍA ENRIQUETA RIVERA DEL BOSQUE"
]

# --- LISTA OFICIAL Y FIJA DE TALLERES / CURSOS ---
TALLERES_FIJOS = [
    "ACTIVIDAD FÍSICA Y DEPORTE",
    "ALIMENTACIÓN Y NUTRICIÓN",
    "ALTA DE TALLER SISTEMA",
    "ARTE Y PINTURA",
    "ARTES PLÁSTICAS",
    "ASAMBLEA DE BIENVENIDA",
    "ASISTENCIA SOCIAL Y COMUNITARIA",
    "AUTOESTIMA Y DESARROLLO PERSONAL",
    "AUTONOMÍA Y VIDA INDEPENDIENTE",
    "BAILE Y EXPRESIÓN CORPORAL",
    "CINE DEBATE Y REFLEXIÓN",
    "COCINA Y REPOSTERÍA",
    "COMPRENSIÓN LECTORA Y ESCRITURA",
    "COMPUTACIÓN Y TECNOLOGÍA",
    "COMUNICACIÓN Y LENGUAJE",
    "CONVIVENCIA INSTITUCIONAL",
    "CUIDADO PERSONAL E HIGIENE",
    "DERECHOS HUMANOS E INCLUSIÓN",
    "EXPRESIÓN EMOCIONAL Y PSICOLOGÍA",
    "HABILIDADES COGNITIVAS",
    "HABILIDADES SOCIALES",
    "HUERTO Y JARDINERÍA",
    "INGLÉS BÁSICO",
    "LECTOESCRITURA",
    "MANUALIDADES Y ARTESANÍAS",
    "MATEMÁTICAS FUNCIONALES",
    "MEDITACIÓN Y RELAJACIÓN",
    "MÚSICA Y CANTO",
    "ORIENTACIÓN ESPACIAL Y TIEMPO",
    "PSICOMOTRICIDAD FINA Y GRUESA",
    "SALIDAS RECREATIVAS Y CULTURALES",
    "TEATRO Y DRAMATIZACIÓN"
]

# Columnas definitivas para el archivo Excel
COL_FECHA = "FECHA"
COL_ASISTENCIA = "INTEGRANTE / TALLER"
COL_TALLER = "ACTIVIDAD"
COL_HORAS = "HORAS"
CLAVE_BORRADO = "LupitaMentes1978"

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
        file_content = repo.get_contents(EXCEL_FILE)
        df = pd.read_excel(io.BytesIO(file_content.decoded_content))
        sha = file_content.sha
    except Exception:
        df = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS])
        sha = None

    integrantes = INTEGRANTES_FIJOS
    talleres = sorted(list(set(TALLERES_FIJOS)))
    return integrantes, talleres, df, sha

def guardar_en_github(df_nuevo, sha_actual, mensaje_commit):
    repo = conectar_github()
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_nuevo.to_excel(writer, index=False)
        excel_bytes = output.getvalue()

        if sha_actual:
            repo.update_file(EXCEL_FILE, mensaje_commit, excel_bytes, sha_actual)
        else:
            repo.create_file(EXCEL_FILE, mensaje_commit, excel_bytes)
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar en GitHub: {e}")
        return False

# Cargar datos iniciales
lista_integrantes, lista_talleres, df_original, archivo_sha = cargar_menus_y_datos()

# --- ENCABEZADO VISUAL INSTITUCIONAL ---
col_logo_1, col_logo_2, col_logo_3 = st.columns([1, 1.2, 1])
with col_logo_2:
    st.image(URL_LOGO, use_container_width=True)

st.markdown('<div class="titulo-web">Control de Asistencia Grupal</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo-web">Comunidad de Vida para Adultos con Parálisis Cerebral</div>', unsafe_allow_html=True)

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
                COL_FECHA: fecha_hoy, 
                COL_ASISTENCIA: "ALTA DE TALLER SISTEMA", 
                COL_TALLER: nuevo_taller_input, 
                COL_HORAS: 0.00
            }
            df_actualizado_taller = pd.concat([df_original, pd.DataFrame([nueva_fila_taller])], ignore_index=True)
            
            if guardar_en_github(df_actualizado_taller, archivo_sha, f"Sistema: Alta de taller {nuevo_taller_input}"):
                st.success(f"✨ ¡Taller '{nuevo_taller_input}' guardado en GitHub!")
                st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

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

    st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
    col_izq, col_der = st.columns(2)
    for i, nombre in enumerate(integrantes_filtrados):
        valor_previo = st.session_state.asistencia_estados.get(nombre, True)
        if i % 2 == 0:
            with col_izq:
                st.session_state.asistencia_estados[nombre] = st.checkbox(nombre, value=valor_previo, key=f"chk_{nombre}")
        else:
            with col_der:
                st.session_state.asistencia_estados[nombre] = st.checkbox(nombre, value=valor_previo, key=f"chk_{nombre}")
    st.markdown('</div>', unsafe_allow_html=True)
                
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
                COL_FECHA: fecha_formateada, 
                COL_ASISTENCIA: integrante, 
                COL_TALLER: taller, 
                COL_HORAS: float(horas)
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
    try:
        df_ordenado = df_original.copy()
        df_ordenado['FECHA_DATETIME'] = pd.to_datetime(df_ordenado[COL_FECHA], format="%d/%m/%Y", errors='coerce')
        df_ordenado = df_ordenado.sort_values(by='FECHA_DATETIME', ascending=True)
        df_ordenado = df_ordenado.drop(columns=['FECHA_DATETIME'])
    except Exception:
        df_ordenado = df_original

    output_descarga = io.BytesIO()
    with pd.ExcelWriter(output_descarga, engine='openpyxl') as writer:
        df_ordenado.to_excel(writer, index=False)
    excel_completo_bytes = output_descarga.getvalue()
    
    st.download_button(
        label="📥 Descargar Base de Datos Completa (Excel Ordenado)",
        data=excel_completo_bytes,
        file_name=f"asistencia_completa_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.dataframe(df_original.tail(15).iloc[::-1], use_container_width=True)
else:
    st.info("💡 Aún no hay registros guardados en el archivo Excel de GitHub.")

# --- SECCIÓN D: ADMINISTRACIÓN DE SEGURIDAD (BORRAR REGISTROS POR CASILLAS) ---
st.markdown("<br>", unsafe_allow_html=True)
with st.expander("🚨 Panel de Administración - Borrado Específico por Persona", expanded=False):
    st.write("Esta sección te permite seleccionar de forma exacta qué personas o registros deseas eliminar de la base de datos.")
    clave_input = st.text_input("Ingresa la clave de autorización:", type="password", key="clave_borrar")
    
    if clave_input == CLAVE_BORRADO:
        st.warning("⚠️ Clave correcta. Configura los filtros de abajo para buscar y seleccionar los registros a eliminar.")
        
        if not df_original.empty:
            # 1. Obtener todas las fechas únicas de la base de datos
            fechas_unicas = sorted(df_original[COL_FECHA].dropna().unique(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"), reverse=True)
            fecha_seleccionada = st.selectbox("1. Selecciona la fecha donde se encuentra el error:", fechas_unicas, key="fecha_eliminar_individual")
            
            # 2. Filtrar el DataFrame original para obtener solo los registros de esa fecha seleccionada
            df_dia = df_original[df_original[COL_FECHA] == fecha_seleccionada].copy()
            
            if not df_dia.empty:
                st.markdown("### 2. Selecciona las casillas de las personas que deseas BORRAR:")
                st.write("*(Los registros que dejes SIN MARCAR se conservarán intactos)*")
                
                registros_a_eliminar = []
                
                # Forzar contenedor blanco de contraste para que las casillas resalten en el celular
                st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
                for idx, fila in df_dia.iterrows():
                    info_registro = f"👤 {fila[COL_ASISTENCIA]} | 📚 {fila[COL_TALLER]} ({fila[COL_HORAS]} hrs)"
                    marcado = st.checkbox(info_registro, value=False, key=f"del_{idx}")
                    if marcado:
                        registros_a_eliminar.append(idx)
                st.markdown('</div>', unsafe_allow_html=True)
                
                # 3. Procesar la eliminación de las casillas seleccionadas
                if registros_a_eliminar:
                    st.error(f"🚨 Has seleccionado {len(registros_a_eliminar)} registro(s) para ser eliminado(s) de forma permanente.")
                    boton_ejecutar_borrado = st.button("❌ Confirmar: Eliminar registros seleccionados")
                    
                    if boton_ejecutar_borrado:
                        df_resultado = df_original.drop(index=registros_a_eliminar)
                        
                        if guardar_en_github(df_resultado, archivo_sha, f"Admin: Eliminación de {len(registros_a_eliminar)} registros individuales del día {fecha_seleccionada}"):
                            st.success("🎉 ¡Los registros seleccionados han sido removidos con éxito de GitHub!")
                            st.rerun()
                else:
                    st.info("💡 Selecciona una o más casillas de la lista superior cuando desees borrar a alguien.")
            else:
                st.info("No se encontraron registros para la fecha seleccionada.")
        else:
            st.info("💡 La base de datos está completamente vacía. No hay registros que borrar.")
            
    elif clave_input != "":
        st.error("❌ Clave incorrecta. No tienes autorización para realizar esta acción.")

# --- PIE DE PÁGINA INSTITUCIONAL ---
st.markdown("""
    <div class="footer-web">
        <hr>
        Av. Ocampo 1797 ote. C.P. 27000, Col. Centro Torreón, Coah.<br>
        <b>MENTES CON ALAS ES DONATARIA AUTORIZADA</b><br>
        <a href="https://mentesconalas.org.mx" target="_blank">🌐 Visitar Sitio Web Oficial</a>
    </div>
""", unsafe_allow_html=True)
