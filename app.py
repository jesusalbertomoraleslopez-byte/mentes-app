import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
from github import Github
import io
import calendar
import plotly.graph_objects as go

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE PÁGINA
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Mentes Con Alas — Asistencia",
    page_icon="🦅",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_FILE      = "asistencia.xlsx"
CALENDARIO_FILE = "calendario.xlsx"
INTEGRANTES_FILE = "integrantes.xlsx"
NO_LABORABLES_FILE = "no_laborables.xlsx"
ACTIVIDADES_FILE = "actividades.xlsx"
USUARIOS_FILE    = "usuarios.xlsx"
MAESTROS_FILE    = "maestros.xlsx"
URL_LOGO        = "logo-mentes.png"

# ══════════════════════════════════════════════════════════════
# CSS INSTITUCIONAL — paleta mentesconalas.org.mx
# ══════════════════════════════════════════════════════════════
_CSS = """
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&family=Open+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">
<style>

/* ── BASE ─────────────────────────────────────────── */
html, body, [data-testid="stAppViewContainer"] {
    background-color: #F2F2EF;
    font-family: 'Open Sans', sans-serif;
    color: #555;
}
#MainMenu, footer, header { visibility: hidden; }

/* ══ SIDEBAR ════════════════════════════════════════ */
[data-testid="stSidebar"] {
    background: #1E1F23 !important;
    border-right: 3px solid #EAB519 !important;
    min-width: 240px !important;
    max-width: 260px !important;
}
[data-testid="stSidebar"] > div { padding: 0 !important; }

/* logo área */
.sb-logo-wrap {
    background: #24252A;
    padding: 22px 18px 16px 18px;
    border-bottom: 2px solid #EAB519;
    text-align: center;
}
.sb-logo-wrap img { width: 110px; margin-bottom: 8px; }
.sb-org-name {
    font-family: 'Montserrat', sans-serif;
    font-size: 11px; font-weight: 800;
    color: #EAB519; letter-spacing: 1.5px;
    text-transform: uppercase; line-height: 1.3;
}

/* ── LOGO BLANCO en sidebar ─────────────────────────── */
[data-testid="stSidebar"] [data-testid="stImage"] img {
    filter: brightness(0) invert(1) !important;
    opacity: 0.92;
}

/* ── OCULTAR botones nativos duplicados del sidebar ─── */
[data-testid="stSidebar"] .stButton > button {
    position: absolute !important;
    opacity: 0 !important;
    pointer-events: auto !important;
    width: 100% !important;
    height: 52px !important;
    top: 0 !important; left: 0 !important;
    cursor: pointer !important;
    z-index: 10 !important;
}
[data-testid="stSidebar"] .stButton {
    position: relative !important;
    margin-top: -54px !important;
    margin-bottom: 2px !important;
}

/* ══ MAIN CONTENT ══════════════════════════════════ */
.main-header {
    display: flex; align-items: center; gap: 14px;
    padding: 18px 0 0 0; margin-bottom: 4px;
}
.main-header-text h1 {
    font-family: 'Montserrat', sans-serif;
    font-size: 22px; font-weight: 800; color: #24252A; margin: 0; line-height: 1.1;
}
.main-header-text p {
    font-family: 'Open Sans', sans-serif;
    font-size: 12px; color: #999; margin: 3px 0 0 0;
}

.gold-divider {
    height: 3px;
    background: linear-gradient(90deg, #EAB519 0%, rgba(234,181,25,0.2) 70%, transparent 100%);
    border: none; margin: 10px 0 20px 0; border-radius: 2px;
}

/* ── KPI CARDS ────────────────────────────────────── */
.kpi-grid {
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 14px; margin-bottom: 24px;
}
.kpi-card {
    background: #FFFFFF; border-radius: 10px; padding: 18px 20px;
    border-left: 4px solid #EAB519;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    transition: transform 0.2s, box-shadow 0.2s;
}
.kpi-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.1); }
.kpi-icon { font-size: 24px; margin-bottom: 8px; }
.kpi-label {
    font-family: 'Open Sans', sans-serif; font-size: 10px; font-weight: 700;
    color: #AAA; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px;
}
.kpi-value {
    font-family: 'Montserrat', sans-serif; font-size: 26px; font-weight: 800;
    color: #24252A; line-height: 1; margin-bottom: 3px;
}

/* ── SECTION TITLE ────────────────────────────────── */
.section-title {
    font-family: 'Montserrat', sans-serif;
    font-size: 15px; font-weight: 800; color: #24252A;
    text-transform: uppercase; letter-spacing: 0.5px;
    border-left: 4px solid #EAB519; padding-left: 12px;
    margin: 24px 0 14px 0;
}

/* ── EXPANDERS ────────────────────────────────────── */
div[data-testid="stExpander"] {
    background: #FFFFFF !important; border: 1px solid #E6E6E2 !important;
    border-radius: 10px !important; box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
    margin-bottom: 10px !important; overflow: hidden !important;
}
div[data-testid="stExpander"] summary {
    font-family: 'Montserrat', sans-serif !important;
    font-size: 13px !important; font-weight: 700 !important; color: #24252A !important;
}

/* ── CONTENEDOR ASISTENCIA ────────────────────────── */
.contenedor-asistencia {
    background: #F8F8F5 !important; border: 1px solid #E6E6E2 !important;
    border-radius: 10px !important; padding: 18px !important; margin: 10px 0 !important;
}

/* ── BOTONES ──────────────────────────────────────── */
.stButton > button {
    background: #EAB519 !important; color: #24252A !important;
    border: none !important; border-radius: 6px !important;
    font-family: 'Montserrat', sans-serif !important; font-weight: 800 !important;
    font-size: 12px !important; letter-spacing: 0.5px !important;
    width: 100% !important; text-transform: uppercase !important;
    box-shadow: 0 3px 10px rgba(234,181,25,0.3) !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    background: #f5ca3a !important; transform: translateY(-1px) !important;
}
div[data-testid="stForm"] button[kind="primaryFormSubmit"] {
    background: #EAB519 !important; color: #24252A !important;
    border: none !important; border-radius: 6px !important;
    font-family: 'Montserrat', sans-serif !important; font-weight: 800 !important;
    width: 100% !important; text-transform: uppercase !important;
}

/* ── FORM ─────────────────────────────────────────── */
div[data-testid="stForm"] {
    background: #FFFFFF; border-radius: 10px;
    border: 1px solid #E6E6E2; padding: 18px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.04);
}

/* ── CHECKBOX ─────────────────────────────────────── */
div[data-testid="stCheckbox"] label p {
    font-family: 'Open Sans', sans-serif !important;
    color: #333 !important; font-weight: 600 !important; font-size: 13px !important;
}

/* ── RADIO ────────────────────────────────────────── */
div[data-testid="stRadio"] label {
    font-family: 'Open Sans', sans-serif !important;
    font-weight: 600 !important; color: #333 !important;
    font-size: 13px !important; text-transform: none !important; letter-spacing: 0 !important;
}

/* ── CALENDARIO ───────────────────────────────────── */
.tabla-calendario {
    width: 100%; border-collapse: collapse; background: #fff;
    border-radius: 10px; overflow: hidden;
    box-shadow: 0 2px 10px rgba(0,0,0,0.07); table-layout: fixed;
}
.tabla-calendario th {
    background: #24252A; color: #EAB519; padding: 10px 4px; font-weight: 800;
    text-align: center; font-size: 11px; border: 1px solid #333;
}
.tabla-calendario td {
    height: 90px; vertical-align: top; padding: 5px 4px;
    border: 1px solid #EAEAE6; background: #fff; overflow: hidden;
}
.num-dia { font-family: 'Montserrat', sans-serif; font-weight: 800; color: #555; font-size: 12px; }
.dia-vacio { background: #F8F8F5 !important; }
.dia-hoy { background: rgba(234,181,25,0.07) !important; border: 2px solid #EAB519 !important; }

/* ── CONFLICT PANEL ───────────────────────────────── */
.conflict-card {
    background: #FFF3CD; border-left: 5px solid #FFC107;
    border-radius: 8px; padding: 12px 15px; margin-bottom: 10px;
    font-size: 12px; color: #856404;
}
.conflict-title { font-weight: 800; margin-bottom: 3px; font-family: 'Montserrat', sans-serif; }

/* ── FOOTER ───────────────────────────────────────── */
.mca-footer {
    background: #24252A; color: rgba(255,255,255,0.4);
    text-align: center; padding: 16px 20px;
    margin: 40px -3rem -3rem -3rem;
    font-family: 'Open Sans', sans-serif; font-size: 11px;
    border-top: 3px solid #EAB519;
}
.mca-footer a { color: #EAB519; text-decoration: none; }
</style>
"""
st.html(_CSS)

# ══════════════════════════════════════════════════════════════
# INTEGRANTES INICIALES (PRECARGA SI NO EXISTE DATABASE)
# ══════════════════════════════════════════════════════════════
INTEGRANTES_PRECARGA = [
    "ANA DE LOS ÁNGELES TORRES REVELES", "CARLOS ALBERTO DE LA TORRE SANTELLANO",
    "CRISTABEL DE LA CRUZ MALDONADO", "EFRAÍN MAYNEZ PORRAS",
    "FERNANDO ÁVILA BERLANGA", "FLORINDA ESTEVANÉ PIZARRO",
    "GUADALUPE LÓPEZ TOVAR", "ISAAC IGNACIO GONZÁLEZ CRUZ",
    "JESÚS ALEJANDRO SIFUENTES ESPINO", "JESÚS SALCIDO ZAMORA",
    "JOSE REYNALDO ALCORTA BENAVIDES", "JUAN JOSÉ OCHOA GONZÁLEZ",
    "JUAN RAFAEL YAÑEZ SERNA", "JUAN SILVERIO LÓPEZ DE LA ROSA",
    "KARLA LISSETTE PEDROZA GONZÁLEZ", "LUIS FERNANDO ÁVILA BERLANGA",
    "LUIS JAVIER GARCÍA MARTÍNEZ", "MA. ELIZABETH GONZÁLEZ HIDROGO",
    "MARÍA DE LOS ÁNGELES ORDUÑA RODRÍGUEZ",
    "MARÍA GUADALUPE DE LA CONCEPCIÓN TORRES LIRA",
    "PEDRO ANTONIO DE LA CERDA TREVIÑO",
    "TERESITA DEL NIÑO JESÚS RODRÍGUEZ SALAZAR",
    "TOMASITA MARÍA ENRIQUETA RIVERA DEL BOSQUE",
]
# TALLERES_FIJOS y COLORES_TALLERES se cargan ahora de forma dinámica desde la base de datos de actividades.

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
MAPA_DIAS = {0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves", 4: "Viernes", 5: "Sábado", 6: "Domingo"}

COL_FECHA      = "FECHA"
COL_ASISTENCIA = "INTEGRANTE / TALLER"
COL_TALLER     = "ACTIVIDAD"
COL_HORAS      = "HORAS"
CLAVE_BORRADO  = "LupitaMentes1978"

PAGINAS = [
    ("📅", "Planeación",           "Programación semestral y festivos"),
    ("🗓️", "Agenda Visual",        "Calendario mensual de actividades"),
    ("📝", "Registro de Asistencia","Captura de asistencia real"),
    ("📊", "Revisión e Informes",   "Historial, estadísticas y reportes"),
    ("👥", "Integrantes",           "Registro y base de datos de alumnos"),
]

# ══════════════════════════════════════════════════════════════
# CONEXIÓN GITHUB
# ══════════════════════════════════════════════════════════════
def conectar_github():
    try:
        token = st.secrets.get("GITHUB_TOKEN", None)
        repo_name = st.secrets.get("GITHUB_REPO", None)
        if not token or not repo_name:
            st.error("⚙️ **Configuración incompleta**: faltan credenciales de GitHub en los secrets.")
            st.stop()
        return Github(token).get_repo(repo_name)
    except Exception as e:
        st.error(f"🌐 **Error de conexión GitHub:** {e}")
        st.stop()


@st.cache_data(ttl=60, show_spinner=False)
def cargar_datos_sistema(_ts=None):
    repo = conectar_github()
    
    # 1. Asistencia
    try:
        f = repo.get_contents(EXCEL_FILE)
        df_a = pd.read_excel(io.BytesIO(f.decoded_content))
        sha_a = f.sha
    except Exception:
        df_a, sha_a = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS]), None

    # 2. Calendario
    try:
        f2 = repo.get_contents(CALENDARIO_FILE)
        df_c = pd.read_excel(io.BytesIO(f2.decoded_content))
        sha_c = f2.sha
    except Exception:
        df_c, sha_c = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS, "HORARIO"]), None

    # 3. Integrantes
    try:
        f3 = repo.get_contents(INTEGRANTES_FILE)
        df_i = pd.read_excel(io.BytesIO(f3.decoded_content))
        sha_i = f3.sha
    except Exception:
        # Precargar si no existe
        df_i = pd.DataFrame([
            {
                "NOMBRE_COMPLETO": name,
                "ESTADO": "Activo",
                "FECHA_REGISTRO": datetime.now().strftime("%d/%m/%Y"),
                "TELEFONO_CONTACTO": "—",
                "DATOS_MEDICOS": "Parálisis Cerebral",
                "NOTAS": "Precargado del sistema anterior."
            } for name in INTEGRANTES_PRECARGA
        ])
        sha_i = None

    # 4. No laborables
    try:
        f4 = repo.get_contents(NO_LABORABLES_FILE)
        df_n = pd.read_excel(io.BytesIO(f4.decoded_content))
        sha_n = f4.sha
    except Exception:
        df_n = pd.DataFrame(columns=["FECHA", "TIPO", "DESCRIPCIÓN"])
        sha_n = None

    # 5. Usuarios
    try:
        f5 = repo.get_contents(USUARIOS_FILE)
        df_u = pd.read_excel(io.BytesIO(f5.decoded_content))
        sha_u = f5.sha
    except Exception:
        df_u = pd.DataFrame([
            {"USUARIO": "admin", "CONTRASEÑA": "mentes2026", "PERFIL": "Administrador"}
        ])
        sha_u = None

    # 6. Actividades
    try:
        f6 = repo.get_contents(ACTIVIDADES_FILE)
        df_act = pd.read_excel(io.BytesIO(f6.decoded_content))
        sha_act = f6.sha
    except Exception:
        t_list = [
            ("AJEDREZ", "—", "#24252A", "Cognitiva / Educativa"),
            ("ARTE", "—", "#8B5E3C", "Socio-Emocional"),
            ("DESARROLLO EDUCATIVO", "—", "#1B5E20", "Cognitiva / Educativa"),
            ("FELDENKRAIS", "—", "#4A148C", "Física / Motora"),
            ("GRUPO DE CRECIMIENTO", "—", "#BF360C", "Socio-Emocional"),
            ("MOVIMIENTO VITAL EXPRESIVO", "—", "#0D47A1", "Física / Motora"),
            ('TALLER "SÍ PUEDO"', "—", "#006064", "Socio-Emocional"),
            ("TALLER DE COMUNICACIÓN", "—", "#558B2F", "Cognitiva / Educativa"),
            ("TEATRO", "—", "#E65100", "Socio-Emocional"),
            ("VIDA DIARIA", "—", "#37474F", "Autonomía / Vida Diaria"),
            ("OTRA", "—", "#546E7A", "Socio-Emocional"),
            ("PSICOLOGÍA", "—", "#880E4F", "Socio-Emocional"),
            ("MOVIMIENTO", "—", "#1565C0", "Física / Motora"),
            ("ACTIVIDAD FUERA DE COMUNIDAD", "—", "#33691E", "Autonomía / Vida Diaria")
        ]
        df_act = pd.DataFrame(t_list, columns=["ACTIVIDAD", "MAESTRO", "COLOR", "ÁREA_IMPACTO"])
        sha_act = None

    # 7. Maestros
    try:
        f7 = repo.get_contents(MAESTROS_FILE)
        df_m = pd.read_excel(io.BytesIO(f7.decoded_content))
        sha_m = f7.sha
    except Exception:
        df_m = pd.DataFrame([
            {"NOMBRE_COMPLETO": "—", "TELÉFONO": "—", "ESPECIALIDAD": "General", "ESTADO": "Activo"}
        ])
        sha_m = None

    return df_a, sha_a, df_c, sha_c, df_i, sha_i, df_n, sha_n, df_u, sha_u, df_act, sha_act, df_m, sha_m


def guardar_github(nombre, df, sha, msg):
    repo = conectar_github()
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        data = buf.getvalue()
        if sha:
            repo.update_file(nombre, msg, data, sha)
        else:
            repo.create_file(nombre, msg, data)
        cargar_datos_sistema.clear()
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar `{nombre}`: {e}")
        return False

# ══════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ══════════════════════════════════════════════════════════════
with st.spinner("🔄 Conectando..."):
    df_original, archivo_sha, df_calendario, sha_calendario, df_integrantes, sha_integrantes, df_no_laborables, sha_no_laborables, df_usuarios, sha_usuarios, df_actividades, sha_actividades, df_maestros, sha_maestros = cargar_datos_sistema()

# Garantizar bases de datos en Github si fueron autogeneradas
if sha_integrantes is None:
    guardar_github(INTEGRANTES_FILE, df_integrantes, None, "Inicializar base de datos de integrantes")
if sha_usuarios is None:
    guardar_github(USUARIOS_FILE, df_usuarios, None, "Inicializar base de datos de usuarios")
if sha_actividades is None:
    guardar_github(ACTIVIDADES_FILE, df_actividades, None, "Inicializar base de datos de actividades")
if sha_maestros is None:
    guardar_github(MAESTROS_FILE, df_maestros, None, "Inicializar base de datos de maestros")

# ══════════════════════════════════════════════════════════════
# VARIABLES GLOBALES DINÁMICAS (CATÁLOGOS)
# ══════════════════════════════════════════════════════════════
TALLERES_FIJOS = df_actividades["ACTIVIDAD"].tolist() if not df_actividades.empty else []
COLORES_TALLERES = dict(zip(df_actividades["ACTIVIDAD"], df_actividades["COLOR"])) if not df_actividades.empty else {}

# ══════════════════════════════════════════════════════════════
# SESSION STATE — página y autenticación
# ══════════════════════════════════════════════════════════════
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = None
if "perfil_actual" not in st.session_state:
    st.session_state.perfil_actual = None

# Cierre de sesión por query params
try:
    qp = st.query_params
except AttributeError:
    qp = st.experimental_get_query_params()

if qp.get("logout") == "true" or qp.get("logout") == ["true"]:
    st.session_state.authenticated = False
    st.session_state.usuario_actual = None
    st.session_state.perfil_actual = None
    try:
        st.query_params.clear()
    except AttributeError:
        st.experimental_set_query_params()
    st.rerun()

if "pagina" not in st.session_state:
    st.session_state.pagina = "Planeación"

# Lista de integrantes activos
lista_integrantes_activos = df_integrantes[df_integrantes["ESTADO"] == "Activo"]["NOMBRE_COMPLETO"].dropna().tolist()

# ══════════════════════════════════════════════════════════════
# CALLBACKS PARA SELECCIÓN MASIVA
# ══════════════════════════════════════════════════════════════
def toggle_todos_cal():
    val = st.session_state.todos_cal
    for nombre in lista_integrantes_activos:
        st.session_state[f"cal_{nombre}"] = val

def toggle_todos_dir():
    val = st.session_state.todos_dir
    for nombre in lista_integrantes_activos:
        st.session_state[f"chk_d_{nombre}"] = val

# ══════════════════════════════════════════════════════════════
# CONTROL DE ACCESO (LOGIN)
# ══════════════════════════════════════════════════════════════
if not st.session_state.get("authenticated", False):
    # Ocultar barra lateral y controles nativos con CSS
    st.markdown("""
        <style>
            [data-testid="stSidebar"] { display: none !important; }
            [data-testid="stSidebarCollapsedControl"] { display: none !important; }
            [data-testid="stHeader"] { display: none !important; }
            html, body, [data-testid="stAppViewContainer"] {
                background-color: #1E1F23 !important;
            }
        </style>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)
        try:
            st.image(URL_LOGO, width=160)
        except Exception:
            st.markdown("<h1 style='text-align:center; color:#EAB519;'>🦅</h1>", unsafe_allow_html=True)
            
        st.markdown("""
        <div style="background:#FFFFFF; padding:30px 25px; border-radius:15px; border-top: 5px solid #EAB519; text-align:center; box-shadow: 0 4px 20px rgba(0,0,0,0.15); margin-bottom: 20px;">
            <h3 style="font-family:'Montserrat',sans-serif; font-weight:800; color:#24252A; margin:0 0 4px 0;">INICIAR SESIÓN</h3>
            <p style="font-family:'Open Sans',sans-serif; font-size:11px; color:#888; margin:0 0 16px 0;">SISTEMA DE ASISTENCIA — MENTES CON ALAS</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("form_login"):
            usr = st.text_input("👤 USUARIO", placeholder="Nombre de usuario").strip().lower()
            pwd = st.text_input("🔑 CONTRASEÑA", type="password", placeholder="Contraseña de acceso")
            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            entrar = st.form_submit_button("INGRESAR")
            
            if entrar:
                if not df_usuarios.empty and usr in df_usuarios["USUARIO"].astype(str).str.lower().tolist():
                    fila_u = df_usuarios[df_usuarios["USUARIO"].astype(str).str.lower() == usr].iloc[0]
                    if str(fila_u["CONTRASEÑA"]) == pwd:
                        st.session_state.authenticated = True
                        st.session_state.usuario_actual = fila_u["USUARIO"]
                        st.session_state.perfil_actual = fila_u["PERFIL"]
                        st.success("🎉 Acceso concedido.")
                        st.rerun()
                    else:
                        st.error("❌ Contraseña incorrecta.")
                else:
                    st.error("❌ Usuario no encontrado.")
        
        st.markdown("""
        <p style="text-align:center; font-family:'Open Sans',sans-serif; font-size:10px; color:rgba(255,255,255,0.4); margin-top:20px;">
            🦅 Mentes con Alas A.C. &copy; 2026
        </p>
        """, unsafe_allow_html=True)
    st.stop()

# Construir lista de páginas según el perfil
paginas_visibles = PAGINAS.copy()
if str(st.session_state.get("perfil_actual", "Staff")).strip().lower() == "administrador":
    paginas_visibles.append(("⚙️", "Mantenimiento", "Catálogo de talleres y usuarios"))

# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    try:
        st.image(URL_LOGO, width=120)
    except Exception:
        st.markdown("🦅")

    st.markdown("""
    <div style="text-align:center;padding:0 0 14px 0;border-bottom:2px solid #EAB519;margin-bottom:4px;">
        <div style="font-family:'Montserrat',sans-serif;font-size:12px;font-weight:800;
                    color:#EAB519;letter-spacing:1.5px;text-transform:uppercase;line-height:1.4;">
            Mentes Con Alas
        </div>
        <div style="font-family:'Open Sans',sans-serif;font-size:10px;
                    color:rgba(255,255,255,0.4);margin-top:3px;">
            Sistema de Asistencia
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Mostrar info del usuario firmado
    st.markdown(f"""
    <div style="background:rgba(255,255,255,0.05); padding:10px 12px; border-radius:6px; margin: 8px 0; border-left:3px solid #EAB519;">
        <div style="font-family:'Open Sans',sans-serif; font-size:11px; color:rgba(255,255,255,0.5);">Firmado como:</div>
        <div style="font-family:'Montserrat',sans-serif; font-size:13px; font-weight:700; color:#EAB519;">{st.session_state.usuario_actual.upper()}</div>
        <div style="font-family:'Open Sans',sans-serif; font-size:10px; color:rgba(255,255,255,0.4); font-style:italic;">{st.session_state.perfil_actual}</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="font-family:'Open Sans',sans-serif;font-size:9px;font-weight:700;
                color:rgba(255,255,255,0.3);text-transform:uppercase;letter-spacing:1.5px;
                padding:10px 0 6px 0;">
        Módulos
    </div>
    """, unsafe_allow_html=True)

    for icono, nombre, desc in paginas_visibles:
        activo = st.session_state.pagina == nombre
        bg     = "rgba(234,181,25,0.14)" if activo else "transparent"
        borde  = "#EAB519"              if activo else "transparent"
        color  = "#EAB519"              if activo else "rgba(255,255,255,0.72)"
        peso   = "700"                  if activo else "600"

        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;
                    cursor:pointer;border-left:3px solid {borde};
                    background:{bg};border-radius:0 6px 6px 0;margin-bottom:2px;">
            <span style="font-size:17px;width:24px;text-align:center;">{icono}</span>
            <div>
                <div style="font-family:'Open Sans',sans-serif;font-size:13px;
                            font-weight:{peso};color:{color};line-height:1.2;">{nombre}</div>
                <div style="font-family:'Open Sans',sans-serif;font-size:10px;
                            color:rgba(255,255,255,0.3);margin-top:1px;">{desc}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(f"{icono} {nombre}", key=f"nav_{nombre}", use_container_width=True):
            st.session_state.pagina = nombre
            st.rerun()

    st.markdown("""
    <div style="border-top:1px solid rgba(255,255,255,0.07);margin:12px 0;"></div>
    <div style="font-family:'Open Sans',sans-serif;font-size:9px;font-weight:700;
                color:rgba(255,255,255,0.3);text-transform:uppercase;letter-spacing:1.5px;">
        Información
    </div>
    """, unsafe_allow_html=True)

    hoy_sb = datetime.now().strftime("%d / %b / %Y")
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:8px;padding:8px 4px;">
        <span style="font-size:16px;">📅</span>
        <div>
            <div style="font-family:'Open Sans',sans-serif;font-size:11px;
                        color:rgba(255,255,255,0.6);font-weight:600;">{hoy_sb}</div>
            <div style="font-family:'Open Sans',sans-serif;font-size:9px;
                        color:rgba(255,255,255,0.3);">Torreón, Coahuila</div>
        </div>
    </div>
    <div style="display:flex;align-items:center;gap:8px;padding:4px 4px 12px 4px;">
        <span style="font-size:16px;">👥</span>
        <div>
            <div style="font-family:'Open Sans',sans-serif;font-size:11px;
                        color:rgba(255,255,255,0.6);font-weight:600;">{len(df_integrantes)} Integrantes</div>
            <div style="font-family:'Open Sans',sans-serif;font-size:9px;
                        color:rgba(255,255,255,0.3);">{len(lista_integrantes_activos)} activos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style='border-top:1px solid rgba(255,255,255,0.07);margin:12px 0;'></div>
    <a href="?logout=true" target="_self" style="
        display: block;
        text-align: center;
        background: #dc3545;
        color: #FFF !important;
        font-family: 'Montserrat', sans-serif;
        font-weight: 800;
        font-size: 11px;
        letter-spacing: 0.5px;
        padding: 10px 15px;
        border-radius: 6px;
        text-transform: uppercase;
        text-decoration: none;
        box-shadow: 0 3px 10px rgba(220,53,69,0.2);
        transition: all 0.2s;
        margin-top: 5px;
        margin-bottom: 25px;
    ">🚪 Cerrar Sesión</a>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ENCABEZADO PRINCIPAL Y BANNER
# ══════════════════════════════════════════════════════════════
try:
    st.image("banner-mentes.png", use_container_width=True)
except Exception:
    pass

pagina_actual = st.session_state.pagina
icono_actual  = next((i for i, n, _ in paginas_visibles if n == pagina_actual), "📋")
desc_actual   = next((d for _, n, d in paginas_visibles if n == pagina_actual), "")

col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(f"""
    <div class="main-header">
        <div><div style="font-size:28px;line-height:1;">{icono_actual}</div></div>
        <div class="main-header-text">
            <h1>{pagina_actual}</h1>
            <p>{desc_actual}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
with col_h2:
    st.markdown(f"""
    <div style="text-align:right;padding-top:18px;">
        <div style="font-family:'Montserrat',sans-serif;font-size:11px;
                    font-weight:800;color:#EAB519;letter-spacing:1px;">HOY</div>
        <div style="font-family:'Montserrat',sans-serif;font-size:14px;
                    font-weight:800;color:#24252A;">{datetime.now().strftime("%d %b %Y")}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<hr class="gold-divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# DETECCIÓN DE CONFLICTOS (FUNCIÓN DE AYUDA)
# ══════════════════════════════════════════════════════════════
def calcular_minutos(t):
    return t.hour * 60 + t.minute

def verificar_solapamiento(time1_start, time1_end, time2_start, time2_end):
    return max(time1_start, time2_start) < min(time1_end, time2_end)

def buscar_conflictos(integrantes, dias_sem, hora_ini, duracion_h, fecha_ini, fecha_fin):
    conflictos = []
    if not integrantes or not dias_sem or df_calendario.empty:
        return conflictos

    min_ini_prop = calcular_minutos(hora_ini)
    min_fin_prop = min_ini_prop + int(duracion_h * 60)
    
    # Crear set de fechas programadas en el rango propuesto que caen en los días seleccionados
    fechas_propuestas = []
    curr = fecha_ini
    while curr <= fecha_fin:
        dia_nombre = MAPA_DIAS[curr.weekday()]
        if dia_nombre in dias_sem:
            fecha_str = curr.strftime("%d/%m/%Y")
            if df_no_laborables.empty or fecha_str not in df_no_laborables["FECHA"].astype(str).tolist():
                fechas_propuestas.append(fecha_str)
        curr += timedelta(days=1)

    if not fechas_propuestas:
        return conflictos

    # Filtrar calendario por fechas y alumnos seleccionados
    df_filtrado = df_calendario[
        (df_calendario[COL_FECHA].isin(fechas_propuestas)) &
        (df_calendario[COL_ASISTENCIA].isin(integrantes))
    ]

    for _, fila in df_filtrado.iterrows():
        horario_str = str(fila.get("HORARIO", ""))
        if "-" in horario_str:
            try:
                ini_str, fin_str = horario_str.split("-")
                hi = datetime.strptime(ini_str.strip(), "%H:%M").time()
                hf = datetime.strptime(fin_str.strip(), "%H:%M").time()
                
                min_ini_exist = calcular_minutos(hi)
                min_fin_exist = calcular_minutos(hf)

                if verificar_solapamiento(min_ini_prop, min_fin_prop, min_ini_exist, min_fin_exist):
                    conflictos.append({
                        "integrante": fila[COL_ASISTENCIA],
                        "fecha": fila[COL_FECHA],
                        "actividad": fila[COL_TALLER],
                        "horario": horario_str
                    })
            except Exception:
                pass
    return conflictos

# ══════════════════════════════════════════════════════════════
# MÓDULO 1: PLANEACIÓN
# ══════════════════════════════════════════════════════════════
if pagina_actual == "Planeación":
    col_izq, col_der = st.columns([2, 1])

    with col_izq:
        st.markdown('<div class="section-title">📅 Programación Semestral / Recurrente</div>', unsafe_allow_html=True)
        
        # Determinar estados de validación dinámicamente
        taller_ok = st.session_state.get("taller_seleccionado", None) is not None
        dias_selec_actuales = [d for d in DIAS_SEMANA if st.session_state.get(f"day_{d}", False)]
        hora_ok = st.session_state.get("hora_inicio_input_sel", None) is not None
        citados_actuales = [n for n in lista_integrantes_activos if st.session_state.get(f"cal_{n}", False)]
        f_inicio_actual = st.session_state.get("f_ini_input", datetime.now().date())
        f_fin_actual = st.session_state.get("f_fin_input", (datetime.now() + timedelta(days=90)).date())

        # 1. Actividad
        tag_taller = " <span style='color:#2e7d32; font-weight:bold; margin-left:8px;'>✅ Validado</span>" if taller_ok else " <span style='color:#e65100; font-size:11px; margin-left:8px;'>(Selecciona un taller)</span>"
        st.markdown(f"<label style='font-weight:700;'>1. TALLER A PROGRAMAR{tag_taller}</label>", unsafe_allow_html=True)
        p_taller = st.selectbox("1. TALLER A PROGRAMAR", TALLERES_FIJOS, index=None, placeholder="— Selecciona un Taller —", key="taller_seleccionado", label_visibility="collapsed")
        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
        
        # 2. Días de la semana
        dias_ok = len(dias_selec_actuales) > 0
        tag_dias = " <span style='color:#2e7d32; font-weight:bold; margin-left:8px;'>✅ Validado</span>" if dias_ok else " <span style='color:#e65100; font-size:11px; margin-left:8px;'>(Selecciona al menos un día)</span>"
        st.markdown(f"<label style='font-weight:700;'>2. DÍAS DE LA SEMANA RECURRENTES{tag_dias}</label>", unsafe_allow_html=True)
        cols_dias = st.columns(7)
        dias_selec = []
        for idx, d in enumerate(DIAS_SEMANA):
            with cols_dias[idx]:
                if st.checkbox(d[:2], key=f"day_{d}"):
                    dias_selec.append(d)
        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
                    
        # 3. Horario
        tag_hora = " <span style='color:#2e7d32; font-weight:bold; margin-left:8px;'>✅ Validado</span>" if hora_ok else " <span style='color:#e65100; font-size:11px; margin-left:8px;'>(Ingresa la hora de inicio)</span>"
        st.markdown(f"<label style='font-weight:700;'>3. HORARIO DE CLASE{tag_hora}</label>", unsafe_allow_html=True)
        c_h1, c_h2 = st.columns(2)
        with c_h1:
            opciones_horas_lista = [f"{h:02d}:{m:02d}" for h in range(8, 19) for m in (0, 15, 30, 45)] + ["19:00"]
            p_ini_sel = st.selectbox("HORA DE INICIO", opciones_horas_lista, index=None, placeholder="— Selecciona una Hora —", key="hora_inicio_input_sel")
        with c_h2:
            p_dur = st.number_input("DURACIÓN (horas)", min_value=0.25, max_value=8.0, value=1.5, step=0.25)
            
        if p_ini_sel is not None:
            p_ini = datetime.strptime(p_ini_sel, "%H:%M").time()
            min_totales = int(p_dur * 60)
            dt_fin      = datetime.combine(datetime.today(), p_ini) + timedelta(minutes=min_totales)
            h_completo  = f"{p_ini.strftime('%H:%M')} - {dt_fin.strftime('%H:%M')}"
            st.info(f"⏰ Horario calculado: **{h_completo}**")
        else:
            p_ini = None
            h_completo = None
            st.info("⏰ Esperando hora de inicio...")
        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

        # 4. Integrantes
        citados_ok = len(citados_actuales) > 0
        tag_citados = " <span style='color:#2e7d32; font-weight:bold; margin-left:8px;'>✅ Validado</span>" if citados_ok else " <span style='color:#e65100; font-size:11px; margin-left:8px;'>(Selecciona al menos un integrante)</span>"
        st.markdown(f"<label style='font-weight:700;'>4. INTEGRANTES PARTICIPANTES{tag_citados}</label>", unsafe_allow_html=True)
        buscar_cal = st.text_input("🔍 Buscar integrante...", placeholder="Escribe un nombre para filtrar lista...")
        lista_filtrada = [n for n in lista_integrantes_activos if buscar_cal.upper() in n] if buscar_cal else lista_integrantes_activos
        
        todos_c = st.checkbox("✅ Seleccionar todos los activos", key="todos_cal", on_change=toggle_todos_cal)
        st.markdown('<div class="contenedor-asistencia" style="max-height: 200px; overflow-y: scroll;">', unsafe_allow_html=True)
        col_l, col_r = st.columns(2)
        estados = {}
        for i, nombre in enumerate(lista_filtrada):
            with col_l if i % 2 == 0 else col_r:
                estados[nombre] = st.checkbox(nombre, key=f"cal_{nombre}")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

        # 5. Rango de Fechas (Semestre)
        periodo_ok = False
        try:
            d_ini = f_inicio_actual.date() if isinstance(f_inicio_actual, datetime) else f_inicio_actual
            d_fin = f_fin_actual.date() if isinstance(f_fin_actual, datetime) else f_fin_actual
            periodo_ok = d_ini <= d_fin
        except Exception:
            periodo_ok = True
            
        tag_periodo = " <span style='color:#2e7d32; font-weight:bold; margin-left:8px;'>✅ Validado</span>" if periodo_ok else " <span style='color:#c62828; font-weight:bold; margin-left:8px;'>❌ Inválido (Fecha inicio mayor a fin)</span>"
        st.markdown(f"<label style='font-weight:700;'>5. PERÍODO DE REPETICIÓN (RANGO SEMESTRAL){tag_periodo}</label>", unsafe_allow_html=True)
        c_f1, c_f2 = st.columns(2)
        with c_f1:
            f_inicio = st.date_input("FECHA DE INICIO", datetime.now(), key="f_ini_input")
        with c_f2:
            f_fin = st.date_input("FECHA DE FINALIZACIÓN", datetime.now() + timedelta(days=90), key="f_fin_input")

        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
        enviar_prog = st.button("🗓️ GUARDAR PROGRAMACIÓN SEMESTRAL", use_container_width=True)

        # Registro de la programación
        if enviar_prog:
            citados = [n for n, v in estados.items() if v]
            if p_taller is None:
                st.warning("⚠️ Debes seleccionar un taller a programar.")
            elif not dias_selec:
                st.warning("⚠️ Debes seleccionar al menos un día de la semana.")
            elif p_ini is None:
                st.warning("⚠️ Debes ingresar una hora de inicio.")
            elif not citados:
                st.warning("⚠️ Debes seleccionar al menos un integrante.")
            elif f_inicio > f_fin:
                st.error("❌ La fecha de inicio no puede ser posterior a la de finalización.")
            else:
                # Generar fechas
                fechas_generadas = []
                curr = f_inicio
                dias_excluidos_festivos = 0
                
                # Lista de no laborables
                lista_no_lab = df_no_laborables["FECHA"].astype(str).tolist() if not df_no_laborables.empty else []

                while curr <= f_fin:
                    dia_nombre = MAPA_DIAS[curr.weekday()]
                    if dia_nombre in dias_selec:
                        fecha_str = curr.strftime("%d/%m/%Y")
                        if fecha_str in lista_no_lab:
                            dias_excluidos_festivos += 1
                        else:
                            fechas_generadas.append(fecha_str)
                    curr += timedelta(days=1)

                if not fechas_generadas:
                    st.warning("⚠️ No se generaron fechas válidas. Revisa que el rango abarque los días de la semana seleccionados y no sean inhábiles.")
                else:
                    # Crear registros
                    nuevos_registros = []
                    for fec in fechas_generadas:
                        for cit in citados:
                            nuevos_registros.append({
                                COL_FECHA: fec,
                                COL_ASISTENCIA: cit,
                                COL_TALLER: p_taller,
                                COL_HORAS: float(p_dur),
                                "HORARIO": h_completo
                            })
                    
                    df_cal_nuevo = pd.concat([df_calendario, pd.DataFrame(nuevos_registros)], ignore_index=True)
                    with st.spinner("Guardando programación..."):
                        if guardar_github(CALENDARIO_FILE, df_cal_nuevo, sha_calendario, 
                                          f"Programación Semestral: {p_taller} ({h_completo})"):
                            st.success(f"🎉 ¡Programación semestral exitosa! Se agendaron **{len(fechas_generadas)}** días, omitiendo **{dias_excluidos_festivos}** días festivos/vacaciones.")
                            st.balloons()
                            st.rerun()

    # Panel Derecho: Alertas de Conflicto en tiempo real
    with col_der:
        st.markdown('<div class="section-title">🚨 Conflictos de Horario</div>', unsafe_allow_html=True)
        st.markdown("<p style='font-size: 11px; color:#777;'>Validación en tiempo real para el rango y horario seleccionados:</p>", unsafe_allow_html=True)
        
        # Obtener los datos actuales del formulario de st.session_state dinámicamente
        citados_tmp = [n for n in lista_integrantes_activos if st.session_state.get(f"cal_{n}", False)]
        dias_tmp = [d for d in DIAS_SEMANA if st.session_state.get(f"day_{d}", False)]
        
        conflictos_encontrados = buscar_conflictos(
            integrantes=citados_tmp,
            dias_sem=dias_tmp,
            hora_ini=p_ini,
            duracion_h=p_dur,
            fecha_ini=f_inicio,
            fecha_fin=f_fin
        )
        
        if conflictos_encontrados:
            st.markdown(f"<div style='color:#b58105; font-weight:800; font-size:13px; margin-bottom:8px;'>⚠️ Se detectaron {len(conflictos_encontrados)} conflictos de agenda:</div>", unsafe_allow_html=True)
            for conf in conflictos_encontrados:
                st.markdown(f"""
                <div class="conflict-card">
                    <div class="conflict-title">👤 {conf['integrante']}</div>
                    📅 {conf['fecha']} · ⏱️ {conf['horario']}<br>
                    ❌ Ocupado en: <b>{conf['actividad']}</b>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.success("✅ ¡Sin conflictos detectados para los alumnos seleccionados en este período!")

    # ── Gestión de Días No Laborables y Vacaciones ──
    st.markdown('<hr class="gold-divider">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🏝️ Calendario de Días No Laborables y Vacaciones</div>', unsafe_allow_html=True)
    
    col_fest1, col_fest2 = st.columns([1, 2])
    with col_fest1:
        with st.form("form_festivo"):
            st.markdown("##### ➕ Registrar Día Inhábil / Vacaciones")
            fest_fecha = st.date_input("FECHA", datetime.now())
            fest_tipo  = st.selectbox("TIPO", ["Día Festivo / Inhábil", "Periodo Vacacional", "Suspensión Extraordinaria"])
            fest_desc  = st.text_input("MOTIVO / DESCRIPCIÓN", placeholder="Ej: Navidad, Junta Consejo...")
            
            save_fest = st.form_submit_button("💾 Guardar Día No Laborable")
            
        if save_fest:
            fest_str = fest_fecha.strftime("%d/%m/%Y")
            if not df_no_laborables.empty and fest_str in df_no_laborables["FECHA"].astype(str).tolist():
                st.warning("⚠️ Esta fecha ya está registrada como inhábil.")
            else:
                nueva_fila = pd.DataFrame([{"FECHA": fest_str, "TIPO": fest_tipo, "DESCRIPCIÓN": fest_desc}])
                df_no_lab_nuevo = pd.concat([df_no_laborables, nueva_fila], ignore_index=True)
                with st.spinner("Guardando festivo..."):
                    if guardar_github(NO_LABORABLES_FILE, df_no_lab_nuevo, sha_no_laborables, f"Festivo: {fest_desc} ({fest_str})"):
                        st.success(f"✅ Registrado exitosamente: {fest_str}")
                        st.rerun()
                        
    with col_fest2:
        st.markdown("##### 📋 Fechas Inhábiles Registradas")
        if df_no_laborables.empty:
            st.info("No hay fechas registradas como días no laborables.")
        else:
            st.dataframe(df_no_laborables, use_container_width=True, height=200)
            if st.checkbox("Habilitar panel de eliminación de festivos"):
                f_del = st.selectbox("Elige fecha a eliminar:", df_no_laborables["FECHA"].unique())
                if st.button("❌ Eliminar Fecha Inhábil"):
                    df_no_lab_nuevo = df_no_laborables[df_no_laborables["FECHA"] != f_del]
                    if guardar_github(NO_LABORABLES_FILE, df_no_lab_nuevo, sha_no_laborables, f"Eliminar festivo: {f_del}"):
                        st.success("Fecha eliminada correctamente.")
                        st.rerun()

    # ── Reprogramación Específica e Individual ──
    st.markdown('<hr class="gold-divider">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">✏️ Reprogramación / Modificación de Días Específicos</div>', unsafe_allow_html=True)
    
    if df_calendario.empty:
        st.info("No hay actividades programadas para modificar.")
    else:
        fechas_disponibles = sorted(df_calendario[COL_FECHA].unique(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"), reverse=True)
        c_mod1, c_mod2 = st.columns(2)
        with c_mod1:
            fecha_mod = st.selectbox("1. Selecciona la fecha específica a modificar:", fechas_disponibles)
        with c_mod2:
            df_dia_mod = df_calendario[df_calendario[COL_FECHA] == fecha_mod].copy()
            df_dia_mod["BLOQUE"] = df_dia_mod[COL_TALLER].astype(str) + " (" + df_dia_mod["HORARIO"].astype(str) + ")"
            bloque_mod = st.selectbox("2. Selecciona la actividad específica:", df_dia_mod["BLOQUE"].unique())
            
        df_bloque_sel = df_dia_mod[df_dia_mod["BLOQUE"] == bloque_mod]
        if not df_bloque_sel.empty:
            taller_mod_original = df_bloque_sel[COL_TALLER].iloc[0]
            horario_mod_original = df_bloque_sel["HORARIO"].iloc[0]
            horas_mod_original = df_bloque_sel[COL_HORAS].iloc[0]
            alumnos_mod_original = df_bloque_sel[COL_ASISTENCIA].unique().tolist()
            original_ini_str = horario_mod_original.split("-")[0].strip()
            opciones_horas_lista = [f"{h:02d}:{m:02d}" for h in range(8, 19) for m in (0, 15, 30, 45)] + ["19:00"]
            try:
                default_idx = opciones_horas_lista.index(original_ini_str)
            except ValueError:
                default_idx = 0

            with st.form("form_mod_especifica"):
                st.markdown(f"##### Modificando sesión de **{taller_mod_original}** el día **{fecha_mod}**")
                c_m_h1, c_m_h2 = st.columns(2)
                with c_m_h1:
                    nuevo_ini_sel = st.selectbox("NUEVA HORA INICIO", opciones_horas_lista, index=default_idx)
                with c_m_h2:
                    nuevas_horas = st.number_input("NUEVA DURACIÓN (horas)", min_value=0.25, max_value=8.0, value=float(horas_mod_original), step=0.25)
                
                nuevo_ini = datetime.strptime(nuevo_ini_sel, "%H:%M").time() if nuevo_ini_sel else datetime.strptime(original_ini_str, "%H:%M").time()
                min_tot_m = int(nuevas_horas * 60)
                dt_fin_m = datetime.combine(datetime.today(), nuevo_ini) + timedelta(minutes=min_tot_m)
                nuevo_horario_completo = f"{nuevo_ini.strftime('%H:%M')} - {dt_fin_m.strftime('%H:%M')}"
                st.info(f"Nuevo horario para el {fecha_mod}: **{nuevo_horario_completo}**")
                
                st.markdown("**👥 Modificar Alumnos Citados para esta sesión:**")
                col_ma_l, col_ma_r = st.columns(2)
                checks_mod = {}
                for idx, n in enumerate(lista_integrantes_activos):
                    with col_ma_l if idx % 2 == 0 else col_ma_r:
                        checks_mod[n] = st.checkbox(n, value=(n in alumnos_mod_original), key=f"ma_{n}")
                        
                btn_save_mod_esp = st.form_submit_button("💾 GUARDAR MODIFICACIÓN DE ESTE DÍA")
                
            if btn_save_mod_esp:
                alumnos_finales_mod = [n for n, v in checks_mod.items() if v]
                # Limpiar anterior
                df_cal_sin_registro = df_calendario[~(
                    (df_calendario[COL_FECHA] == fecha_mod) &
                    (df_calendario[COL_TALLER] == taller_mod_original) &
                    (df_calendario["HORARIO"] == horario_mod_original)
                )]
                if alumnos_finales_mod:
                    filas_actualizadas = [{
                        COL_FECHA: fecha_mod,
                        COL_ASISTENCIA: al,
                        COL_TALLER: taller_mod_original,
                        COL_HORAS: float(nuevas_horas),
                        "HORARIO": nuevo_horario_completo
                    } for al in alumnos_finales_mod]
                    df_cal_modificado = pd.concat([df_cal_sin_registro, pd.DataFrame(filas_actualizadas)], ignore_index=True)
                else:
                    df_cal_modificado = df_cal_sin_registro
                    
                with st.spinner("Guardando reprogramación..."):
                    if guardar_github(CALENDARIO_FILE, df_cal_modificado, sha_calendario, f"Reprogramación específica: {taller_mod_original} el {fecha_mod}"):
                        st.success(f"🎉 Sesión del {fecha_mod} modificada exitosamente.")
                        st.rerun()

    # ── Tabla de Registros Recientes de Planeación ──
    st.markdown('<hr class="gold-divider">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📋 Últimos Registros Guardados en Planeación</div>', unsafe_allow_html=True)
    if df_calendario.empty:
        st.info("No hay actividades registradas en el calendario.")
    else:
        df_ultimos = df_calendario.tail(20).copy()
        st.dataframe(df_ultimos, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# MÓDULO 2: AGENDA VISUAL (CALENDARIO MENSUAL)
# ══════════════════════════════════════════════════════════════
elif pagina_actual == "Agenda Visual":
    st.markdown('<div class="section-title">🗓️ Calendario Mensual de Actividades</div>', unsafe_allow_html=True)

    hoy = datetime.now()
    mes_lista = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                 "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
    cv1, cv2 = st.columns(2)
    with cv1:
        mes_sel = st.selectbox("Mes:", mes_lista, index=hoy.month - 1)
        mes_num = mes_lista.index(mes_sel) + 1
    with cv2:
        anio_sel = st.selectbox("Año:", [hoy.year - 1, hoy.year, hoy.year + 1], index=1)

    semanas    = calendar.Calendar(firstweekday=0).monthdayscalendar(anio_sel, mes_num)
    dic_eventos = {}
    
    if not df_calendario.empty:
        df_u = df_calendario.drop_duplicates(subset=[COL_FECHA, COL_TALLER, "HORARIO"])
        for _, fila in df_u.iterrows():
            try:
                fo = datetime.strptime(str(fila[COL_FECHA]), "%d/%m/%Y")
                if fo.month == mes_num and fo.year == anio_sel:
                    color_ev = COLORES_TALLERES.get(str(fila[COL_TALLER]), "#24252A")
                    dic_eventos.setdefault(fo.day, []).append(
                        f"<div style='background:{color_ev};color:#EAB519;padding:2px 5px;"
                        f"font-size:9px;font-weight:700;border-radius:3px;margin-top:2px;"
                        f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"
                        f"font-family:Open Sans,sans-serif'>"
                        f"<b>{str(fila[COL_TALLER])[:13]}</b><br>⏱ {fila.get('HORARIO','')}</div>"
                    )
            except Exception:
                pass

    html = "<table class='tabla-calendario'>"
    html += "<tr><th>Lun</th><th>Mar</th><th>Mié</th><th>Jue</th><th>Vie</th><th>Sáb</th><th>Dom</th></tr>"
    for semana in semanas:
        html += "<tr>"
        for dia in semana:
            if dia == 0:
                html += "<td class='dia-vacio'></td>"
            else:
                es_hoy = (dia == hoy.day and mes_num == hoy.month and anio_sel == hoy.year)
                cls    = "class='dia-hoy'" if es_hoy else ""
                html  += f"<td {cls}><span class='num-dia'>{dia}</span>"
                for ev in dic_eventos.get(dia, []):
                    html += ev
                html += "</td>"
        html += "</tr>"
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)

    # Leyenda
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title" style="font-size:12px;">Leyenda de Talleres</div>', unsafe_allow_html=True)
    cols_ley = st.columns(7)
    for i, (t, c) in enumerate(COLORES_TALLERES.items()):
        with cols_ley[i % 7]:
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:4px;margin-bottom:4px;'>"
                f"<div style='width:10px;height:10px;background:{c};border-radius:2px;'></div>"
                f"<span style='font-family:Open Sans;font-size:10px;color:#666;'>{t[:13]}</span></div>",
                unsafe_allow_html=True
            )

# ══════════════════════════════════════════════════════════════
# MÓDULO 3: REGISTRO DE ASISTENCIA
# ══════════════════════════════════════════════════════════════
elif pagina_actual == "Registro de Asistencia":
    st.markdown('<div class="section-title">📝 Captura de Asistencia Real</div>', unsafe_allow_html=True)

    metodo = st.radio(
        "Modo de registro:",
        ["📅 A partir del Calendario", "✍️ Registro Directo (Sin Programar)"],
        horizontal=True, key="radio_reg"
    )
    fecha_hoy = st.date_input("Fecha de la actividad:", datetime.now()).strftime("%d/%m/%Y")

    if "Calendario" in metodo:
        df_dia = df_calendario[df_calendario[COL_FECHA] == fecha_hoy].copy() if not df_calendario.empty else pd.DataFrame()
        if df_dia.empty:
            st.info(f"💡 No hay actividades agendadas en el calendario para el **{fecha_hoy}**.")
        else:
            df_dia["BLOQUE"] = df_dia[COL_TALLER].astype(str) + " (" + df_dia["HORARIO"].astype(str) + ")"
            bloque = st.selectbox("Actividad:", sorted(df_dia["BLOQUE"].dropna().unique()))
            df_cit = df_dia[df_dia["BLOQUE"] == bloque]
            taller_p  = df_cit[COL_TALLER].iloc[0]
            horas_p   = df_cit[COL_HORAS].iloc[0]
            horario_p = df_cit["HORARIO"].iloc[0]
            st.success(f"📚 **{taller_p}** · {horario_p} · {horas_p} hrs")

            with st.form("form_cal"):
                st.markdown("**👥 Marca los integrantes que asistieron:**")
                st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
                cr_l, cr_r = st.columns(2)
                checks = {}
                for i, (idx, fila) in enumerate(df_cit.drop_duplicates(COL_ASISTENCIA).iterrows()):
                    n = fila[COL_ASISTENCIA]
                    with cr_l if i % 2 == 0 else cr_r:
                        checks[n] = st.checkbox(n, value=True, key=f"rc_{idx}")
                st.markdown('</div>', unsafe_allow_html=True)
                guardar_btn = st.form_submit_button("💾 Guardar Asistencia")

            if guardar_btn:
                asistieron = [n for n, v in checks.items() if v]
                if not asistieron:
                    st.warning("⚠️ Debes marcar al menos una asistencia.")
                else:
                    act_h = f"{taller_p} ({horario_p})"
                    filas = [{COL_FECHA: fecha_hoy, COL_ASISTENCIA: n, COL_TALLER: act_h, COL_HORAS: float(horas_p)} for n in asistieron]
                    df_n  = pd.concat([df_original, pd.DataFrame(filas)], ignore_index=True)
                    with st.spinner("Guardando..."):
                        if guardar_github(EXCEL_FILE, df_n, archivo_sha, f"Asistencia: {act_h} el {fecha_hoy}"):
                            st.success(f"🎉 ¡{len(asistieron)} asistencias registradas con éxito!")
                            st.balloons()
                            st.rerun()
    else:
        cd1, cd2 = st.columns(2)
        with cd1: t_dir = st.selectbox("TALLER", TALLERES_FIJOS)
        with cd2: h_dir = st.number_input("HORAS", min_value=0.25, max_value=8.0, value=1.5, step=0.25)

        buscar_d = st.text_input("🔍 Buscar integrante...", key="buscar_dir", placeholder="Filtrar lista...")
        lista_d  = [n for n in lista_integrantes_activos if buscar_d.upper() in n] if buscar_d else lista_integrantes_activos
        todos_d  = st.checkbox("✅ Seleccionar todos los activos", key="todos_dir", on_change=toggle_todos_dir)
        st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
        cd_l, cd_r = st.columns(2)
        checks_d = {}
        for i, n in enumerate(lista_d):
            with cd_l if i % 2 == 0 else cd_r:
                checks_d[n] = st.checkbox(n, key=f"chk_d_{n}")
        st.markdown('</div>', unsafe_allow_html=True)
        guardar_d = st.button("💾 Registrar Asistencia Directa", use_container_width=True)

        if guardar_d:
            asistieron_d = [n for n, v in checks_d.items() if v]
            if not asistieron_d:
                st.warning("⚠️ Selecciona al menos un integrante.")
            else:
                filas_d = [{COL_FECHA: fecha_hoy, COL_ASISTENCIA: n, COL_TALLER: t_dir, COL_HORAS: float(h_dir)} for n in asistieron_d]
                df_nd = pd.concat([df_original, pd.DataFrame(filas_d)], ignore_index=True)
                with st.spinner("Guardando..."):
                    if guardar_github(EXCEL_FILE, df_nd, archivo_sha, f"Directo: {t_dir} el {fecha_hoy}"):
                        st.success(f"🎉 {len(asistieron_d)} asistencias registradas.")
                        st.balloons()
                        st.rerun()

# ══════════════════════════════════════════════════════════════
# MÓDULO 4: REVISIÓN E INFORMES
# ══════════════════════════════════════════════════════════════
elif pagina_actual == "Revisión e Informes":
    # 1. Calcular impacto ODS
    ods_acum = {
        "ODS 3: Salud y Bienestar": 0.0,
        "ODS 4: Educación de Calidad": 0.0,
        "ODS 10: Reducción de Desigualdades": 0.0,
        "ODS 11: Comunidades Sostenibles": 0.0
    }
    
    map_area_ods = {
        "Física / Motora": "ODS 3: Salud y Bienestar",
        "Cognitiva / Educativa": "ODS 4: Educación de Calidad",
        "Socio-Emocional": "ODS 10: Reducción de Desigualdades",
        "Autonomía / Vida Diaria": "ODS 11: Comunidades Sostenibles"
    }
    
    map_tal_area = dict(zip(df_actividades["ACTIVIDAD"], df_actividades["ÁREA_IMPACTO"])) if not df_actividades.empty else {}
    
    if not df_original.empty:
        for _, row in df_original.iterrows():
            tal_limpio = str(row[COL_TALLER]).split("(")[0].strip()
            area = map_tal_area.get(tal_limpio, "Socio-Emocional")
            ods_key = map_area_ods.get(area, "ODS 10: Reducción de Desigualdades")
            try:
                horas = float(row[COL_HORAS])
            except Exception:
                horas = 1.5
            ods_acum[ods_key] += horas

    # 2. Renderizar KPIs ODS
    st.markdown('<div class="section-title">📊 Impacto Social Acumulado — Objetivos de Desarrollo Sostenible (ODS)</div>', unsafe_allow_html=True)
    ods_cols = st.columns(4)
    with ods_cols[0]:
        st.markdown(f"""
        <div style="background:#FFF; border-radius:10px; padding:18px 15px; border-left:5px solid #4C9F38; box-shadow: 0 2px 8px rgba(0,0,0,0.06); text-align:center;">
            <div style="font-size:26px; margin-bottom:4px;">🩺</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:10px; font-weight:800; color:#4C9F38; letter-spacing:0.5px; text-transform:uppercase;">ODS 3: SALUD Y BIENESTAR</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:24px; font-weight:800; color:#24252A; margin: 8px 0 2px 0;">{ods_acum['ODS 3: Salud y Bienestar']:.1f} hrs</div>
            <div style="font-size:10px; color:#999; font-style:italic;">Física / Motora</div>
        </div>
        """, unsafe_allow_html=True)
    with ods_cols[1]:
        st.markdown(f"""
        <div style="background:#FFF; border-radius:10px; padding:18px 15px; border-left:5px solid #C7212F; box-shadow: 0 2px 8px rgba(0,0,0,0.06); text-align:center;">
            <div style="font-size:26px; margin-bottom:4px;">🎓</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:10px; font-weight:800; color:#C7212F; letter-spacing:0.5px; text-transform:uppercase;">ODS 4: EDUCACIÓN DE CALIDAD</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:24px; font-weight:800; color:#24252A; margin: 8px 0 2px 0;">{ods_acum['ODS 4: Educación de Calidad']:.1f} hrs</div>
            <div style="font-size:10px; color:#999; font-style:italic;">Cognitiva / Educativa</div>
        </div>
        """, unsafe_allow_html=True)
    with ods_cols[2]:
        st.markdown(f"""
        <div style="background:#FFF; border-radius:10px; padding:18px 15px; border-left:5px solid #DD1367; box-shadow: 0 2px 8px rgba(0,0,0,0.06); text-align:center;">
            <div style="font-size:26px; margin-bottom:4px;">⚖️</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:10px; font-weight:800; color:#DD1367; letter-spacing:0.5px; text-transform:uppercase;">ODS 10: REDUCCIÓN DESIGUALDAD</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:24px; font-weight:800; color:#24252A; margin: 8px 0 2px 0;">{ods_acum['ODS 10: Reducción de Desigualdades']:.1f} hrs</div>
            <div style="font-size:10px; color:#999; font-style:italic;">Socio-Emocional</div>
        </div>
        """, unsafe_allow_html=True)
    with ods_cols[3]:
        st.markdown(f"""
        <div style="background:#FFF; border-radius:10px; padding:18px 15px; border-left:5px solid #F99D26; box-shadow: 0 2px 8px rgba(0,0,0,0.06); text-align:center;">
            <div style="font-size:26px; margin-bottom:4px;">🏙️</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:10px; font-weight:800; color:#F99D26; letter-spacing:0.5px; text-transform:uppercase;">ODS 11: COMUNIDAD SOSTENIBLE</div>
            <div style="font-family:'Montserrat',sans-serif; font-size:24px; font-weight:800; color:#24252A; margin: 8px 0 2px 0;">{ods_acum['ODS 11: Comunidades Sostenibles']:.1f} hrs</div>
            <div style="font-size:10px; color:#999; font-style:italic;">Autonomía / Vida Diaria</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">📋 Historial General de Asistencia</div>', unsafe_allow_html=True)

    if df_original.empty:
        st.info("💡 El historial está vacío.")
    else:
        df_vis = df_original.copy()
        try:
            df_vis["_dt"] = pd.to_datetime(df_vis[COL_FECHA], format="%d/%m/%Y", errors="coerce")
            df_vis = df_vis.sort_values("_dt", ascending=False).drop(columns=["_dt"])
        except Exception:
            pass

        filtro = st.text_input("🔍 Filtrar por nombre, fecha o taller:", placeholder="Escribe para buscar…").strip().upper()
        if filtro:
            mask = (
                df_vis[COL_FECHA].astype(str).str.upper().str.contains(filtro, na=False) |
                df_vis[COL_ASISTENCIA].astype(str).str.upper().str.contains(filtro, na=False) |
                df_vis[COL_TALLER].astype(str).str.upper().str.contains(filtro, na=False)
            )
            df_tabla = df_vis[mask]
        else:
            df_tabla = df_vis

        st.markdown(f'<div style="font-size:12px;color:#AAA;margin-bottom:6px;">{len(df_tabla):,} registros encontrados</div>', unsafe_allow_html=True)
        st.dataframe(df_tabla, use_container_width=True, height=260)

        buf_e = io.BytesIO()
        with pd.ExcelWriter(buf_e, engine="openpyxl") as w:
            df_vis.to_excel(w, index=False)
        st.download_button(
            "📥 Descargar Historial Completo (Excel)", buf_e.getvalue(),
            f"asistencias_mentes_{datetime.now().strftime('%Y%m%d')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # Gráficas
        st.markdown('<div class="section-title">📈 Estadísticas y Distribución</div>', unsafe_allow_html=True)
        g1, g2, g3 = st.columns([1.1, 0.95, 0.95])

        with g1:
            st.markdown("**Asistencias por Alumno**")
            try:
                dp = (df_original[COL_ASISTENCIA].value_counts().reset_index().rename(columns={"index": "I", COL_ASISTENCIA: "N"}))
                dp["Nombre"] = dp[COL_ASISTENCIA].apply(lambda x: " ".join(str(x).split()[:2]))
                fig1 = go.Figure(go.Bar(
                    x=dp["N"], y=dp["Nombre"], orientation="h",
                    marker=dict(color=dp["N"], colorscale=[[0,"#F5EDCC"],[1,"#EAB519"]],
                                line=dict(color="#24252A", width=0.4)),
                    text=dp["N"], textposition="outside",
                ))
                fig1.update_layout(
                    plot_bgcolor="#FFF", paper_bgcolor="#FFF",
                    height=360, margin=dict(l=5, r=30, t=5, b=5),
                    xaxis=dict(showgrid=True, gridcolor="#F0F0EC", zeroline=False),
                    yaxis=dict(autorange="reversed"),
                    font=dict(family="Open Sans", size=10, color="#666"),
                )
                st.plotly_chart(fig1, use_container_width=True)
            except Exception as e:
                st.caption(f"No se pudo generar: {e}")

        with g2:
            st.markdown("**Distribución por Taller**")
            try:
                dt = (df_original[COL_TALLER].value_counts().reset_index().rename(columns={"index": "T", COL_TALLER: "S"}))
                colors_p = [COLORES_TALLERES.get(str(t).split("(")[0].strip(), "#EAB519") for t in dt[COL_TALLER]]
                fig2 = go.Figure(go.Pie(
                    labels=dt[COL_TALLER], values=dt["S"], hole=0.4,
                    marker=dict(colors=colors_p, line=dict(color="#FFF", width=1.5)),
                    textinfo="percent",
                    textfont=dict(family="Open Sans", size=9),
                ))
                fig2.update_layout(
                    plot_bgcolor="#FFF", paper_bgcolor="#FFF",
                    height=360, margin=dict(l=5, r=5, t=5, b=5),
                    showlegend=True,
                    legend=dict(font=dict(size=8), orientation="h", y=-0.1),
                )
                st.plotly_chart(fig2, use_container_width=True)
            except Exception as e:
                st.caption(f"No se pudo generar: {e}")

        with g3:
            st.markdown("**Distribución por Indicador ODS**")
            try:
                df_ods = pd.DataFrame(list(ods_acum.items()), columns=["ODS", "Horas"])
                df_ods = df_ods[df_ods["Horas"] > 0]
                if not df_ods.empty:
                    color_map_ods = {
                        "ODS 3: Salud y Bienestar": "#4C9F38",
                        "ODS 4: Educación de Calidad": "#C7212F",
                        "ODS 10: Reducción de Desigualdades": "#DD1367",
                        "ODS 11: Comunidades Sostenibles": "#F99D26"
                    }
                    colors_ods_chart = [color_map_ods[o] for o in df_ods["ODS"]]
                    
                    fig3 = go.Figure(go.Pie(
                        labels=df_ods["ODS"], values=df_ods["Horas"], hole=0.4,
                        marker=dict(colors=colors_ods_chart, line=dict(color="#FFF", width=1.5)),
                        textinfo="percent",
                        textfont=dict(family="Open Sans", size=9),
                    ))
                    fig3.update_layout(
                        plot_bgcolor="#FFF", paper_bgcolor="#FFF",
                        height=360, margin=dict(l=5, r=5, t=5, b=5),
                        showlegend=True,
                        legend=dict(font=dict(size=8), orientation="h", y=-0.1),
                    )
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("Registra asistencias para ver el impacto ODS.")
            except Exception as e:
                st.caption(f"No se pudo generar: {e}")

    # Panel de Borrado Protegido
    st.markdown('<div class="section-title">🚨 Administración</div>', unsafe_allow_html=True)
    with st.expander("Panel de Eliminación de Registros (Protegido)", expanded=False):
        st.write("Requiere clave de autorización para eliminar registros históricos.")
        clave = st.text_input("🔑 Clave:", type="password", key="clave_admin")

        if clave == CLAVE_BORRADO:
            st.warning("⚠️ **Acceso autorizado.** Selecciona los registros a eliminar.")
            if not df_original.empty:
                fechas_b = sorted(df_original[COL_FECHA].dropna().unique(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"), reverse=True)
                f_sel    = st.selectbox("Filtra por fecha:", fechas_b)
                df_dia_b = df_original[df_original[COL_FECHA] == f_sel]

                st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
                b_regs = []
                for idx, fila in df_dia_b.iterrows():
                    if st.checkbox(f"👤 {fila[COL_ASISTENCIA]}  |  📚 {fila[COL_TALLER]}", key=f"del_{idx}"):
                        b_regs.append(idx)
                st.markdown('</div>', unsafe_allow_html=True)

                if b_regs:
                    st.error(f"🚨 {len(b_regs)} registro(s) marcado(s) para eliminación.")
                    if st.checkbox("☑️ Confirmo la eliminación permanente"):
                        if st.button("❌ ELIMINAR REGISTROS"):
                            df_res = df_original.drop(index=b_regs)
                            with st.spinner("Eliminando..."):
                                if guardar_github(EXCEL_FILE, df_res, archivo_sha, f"Admin: Borrado {len(b_regs)} registros"):
                                    st.success(f"✅ {len(b_regs)} registros eliminados.")
                                    st.rerun()

# ══════════════════════════════════════════════════════════════
# MÓDULO 5: INTEGRANTES
# ══════════════════════════════════════════════════════════════
elif pagina_actual == "Integrantes":
    st.markdown('<div class="section-title">👥 Base de Datos y Registro de Alumnos</div>', unsafe_allow_html=True)

    col_int1, col_int2 = st.columns([1, 2])
    
    # Formulario para agregar / editar integrante
    with col_int1:
        st.markdown("##### ➕ Registrar / Editar Alumno")
        with st.form("form_integrante"):
            nombre_alumno = st.text_input("NOMBRE COMPLETO DEL ALUMNO", placeholder="Ej: JUAN PÉREZ GÓMEZ").strip().upper()
            estado_alumno = st.selectbox("ESTADO DEL ALUMNO", ["Activo", "Inactivo"])
            tel_alumno = st.text_input("TELÉFONO DE CONTACTO / TUTOR", placeholder="Ej: (871) 123-4567")
            datos_medicos = st.text_area("DATOS MÉDICOS / ALERGIAS", placeholder="Indicaciones médicas relevantes, tipo de parálisis...")
            notas_alumno = st.text_area("NOTAS / OBSERVACIONES adicionales")
            
            save_int = st.form_submit_button("💾 Guardar Ficha de Integrante")

        if save_int:
            if not nombre_alumno:
                st.warning("⚠️ El nombre completo es obligatorio.")
            else:
                # Comprobar si existe para actualizar o insertar
                if not df_integrantes.empty and nombre_alumno in df_integrantes["NOMBRE_COMPLETO"].astype(str).tolist():
                    # Modificar existente
                    df_integrantes.loc[df_integrantes["NOMBRE_COMPLETO"] == nombre_alumno, "ESTADO"] = estado_alumno
                    df_integrantes.loc[df_integrantes["NOMBRE_COMPLETO"] == nombre_alumno, "TELEFONO_CONTACTO"] = tel_alumno
                    df_integrantes.loc[df_integrantes["NOMBRE_COMPLETO"] == nombre_alumno, "DATOS_MEDICOS"] = datos_medicos
                    df_integrantes.loc[df_integrantes["NOMBRE_COMPLETO"] == nombre_alumno, "NOTAS"] = notas_alumno
                    msg_commit = f"Actualizar ficha: {nombre_alumno}"
                else:
                    # Crear nuevo
                    nueva_ficha = pd.DataFrame([{
                        "NOMBRE_COMPLETO": nombre_alumno,
                        "ESTADO": estado_alumno,
                        "FECHA_REGISTRO": datetime.now().strftime("%d/%m/%Y"),
                        "TELEFONO_CONTACTO": tel_alumno,
                        "DATOS_MEDICOS": datos_medicos,
                        "NOTAS": notas_alumno
                    }])
                    df_integrantes = pd.concat([df_integrantes, nueva_ficha], ignore_index=True)
                    msg_commit = f"Registrar nuevo integrante: {nombre_alumno}"

                with st.spinner("Guardando en GitHub..."):
                    if guardar_github(INTEGRANTES_FILE, df_integrantes, sha_integrantes, msg_commit):
                        st.success(f"🎉 Ficha guardada con éxito para: {nombre_alumno}")
                        st.rerun()

    # Tabla general con buscador y editor de datos
    with col_int2:
        st.markdown("##### 📋 Listado de Integrantes de la Comunidad")
        
        filtro_int = st.text_input("🔍 Escribe un nombre para buscar:", placeholder="Filtro rápido…").strip().upper()
        
        if not df_integrantes.empty:
            df_int_vis = df_integrantes.copy()
            if filtro_int:
                df_int_vis = df_int_vis[df_int_vis["NOMBRE_COMPLETO"].astype(str).str.upper().str.contains(filtro_int)]
                
            st.dataframe(df_int_vis, use_container_width=True, height=350)
            
            # Cargar datos en el formulario para editar
            st.markdown("<p style='font-size: 11px; color:#777;'>💡 Para editar los detalles de un alumno, escribe su nombre exacto en el formulario de la izquierda.</p>", unsafe_allow_html=True)
        else:
            st.info("La base de datos de integrantes se encuentra vacía.")


# ══════════════════════════════════════════════════════════════
# MÓDULO 6: MANTENIMIENTO
# ══════════════════════════════════════════════════════════════
elif pagina_actual == "Mantenimiento":
    st.markdown('<div class="section-title">⚙️ Panel de Mantenimiento y Configuración del Sistema</div>', unsafe_allow_html=True)
    
    t_usuarios, t_actividades, t_maestros, t_alumnos, t_herramientas = st.tabs([
        "👥 Usuarios del Sistema", 
        "📚 Catálogo de Actividades", 
        "🎓 Catálogo de Maestros Talleristas",
        "👤 Base de Datos Alumnos", 
        "🔧 Herramientas de Datos"
    ])

    # Pestaña 1: Usuarios del Sistema
    with t_usuarios:
        col_u1, col_u2 = st.columns([1, 2])
        with col_u1:
            st.markdown("##### ➕ Registrar / Modificar Usuario")
            with st.form("form_mantenimiento_usuario"):
                n_usr = st.text_input("USUARIO", placeholder="Ej: lupita").strip().lower()
                n_pwd = st.text_input("CONTRASEÑA", placeholder="Ej: clave123")
                n_perfil = st.selectbox("PERFIL DE ACCESO", ["Staff", "Administrador"])
                save_u = st.form_submit_button("💾 Guardar Usuario")
                
            if save_u:
                if not n_usr or not n_pwd:
                    st.warning("⚠️ Debes llenar todos los campos.")
                else:
                    if not df_usuarios.empty and n_usr in df_usuarios["USUARIO"].astype(str).str.lower().tolist():
                        df_usuarios.loc[df_usuarios["USUARIO"].astype(str).str.lower() == n_usr, "CONTRASEÑA"] = n_pwd
                        df_usuarios.loc[df_usuarios["USUARIO"].astype(str).str.lower() == n_usr, "PERFIL"] = n_perfil
                        msg_c = f"Mantenimiento: Actualizar usuario {n_usr}"
                    else:
                        nuevo_u = pd.DataFrame([{"USUARIO": n_usr, "CONTRASEÑA": n_pwd, "PERFIL": n_perfil}])
                        df_usuarios = pd.concat([df_usuarios, nuevo_u], ignore_index=True)
                        msg_c = f"Mantenimiento: Registrar usuario {n_usr}"
                        
                    with st.spinner("Guardando en GitHub..."):
                        if guardar_github(USUARIOS_FILE, df_usuarios, sha_usuarios, msg_c):
                            st.success(f"🎉 Cuenta guardada: {n_usr}")
                            st.rerun()
                            
        with col_u2:
            st.markdown("##### 📋 Cuentas Registradas")
            st.dataframe(df_usuarios, use_container_width=True)
            
            st.markdown("##### ❌ Eliminar Cuenta de Usuario")
            u_eliminar = st.selectbox("Selecciona la cuenta a borrar:", df_usuarios["USUARIO"].tolist() if not df_usuarios.empty else [])
            if st.button("❌ ELIMINAR USUARIO"):
                if u_eliminar == st.session_state.usuario_actual:
                    st.error("❌ No puedes eliminar tu propia cuenta en sesión.")
                elif len(df_usuarios[df_usuarios["PERFIL"] == "Administrador"]) <= 1 and df_usuarios.loc[df_usuarios["USUARIO"] == u_eliminar, "PERFIL"].values[0] == "Administrador":
                    st.error("❌ Debe quedar al menos un Administrador en el sistema.")
                else:
                    df_usuarios = df_usuarios[df_usuarios["USUARIO"] != u_eliminar]
                    with st.spinner("Guardando cambios..."):
                        if guardar_github(USUARIOS_FILE, df_usuarios, sha_usuarios, f"Mantenimiento: Borrar usuario {u_eliminar}"):
                            st.success(f"✅ Usuario {u_eliminar} eliminado.")
                            st.rerun()

    # Pestaña 2: Catálogo de Actividades
    with t_actividades:
        col_a1, col_a2 = st.columns([1, 2])
        
        # Cargar lista de maestros activos para el selector
        lista_maestros_activos = df_maestros[df_maestros["ESTADO"] == "Activo"]["NOMBRE_COMPLETO"].dropna().tolist() if not df_maestros.empty else []
        if "—" not in lista_maestros_activos:
            lista_maestros_activos.insert(0, "—")

        with col_a1:
            st.markdown("##### ➕ Registrar / Modificar Actividad")
            with st.form("form_mantenimiento_actividad"):
                n_act = st.text_input("NOMBRE DEL TALLER / ACTIVIDAD", placeholder="Ej: YOGA").strip().upper()
                n_maestro = st.selectbox("MAESTRO / INSTRUCTOR A CARGO", lista_maestros_activos)
                n_color = st.color_picker("COLOR EN LA AGENDA", "#EAB519")
                n_impacto = st.selectbox("ÁREA DE IMPACTO CLINICO / ODS", [
                    "Física / Motora", 
                    "Cognitiva / Educativa", 
                    "Socio-Emocional", 
                    "Autonomía / Vida Diaria"
                ])
                save_a = st.form_submit_button("💾 Guardar Actividad")
                
            if save_a:
                if not n_act:
                    st.warning("⚠️ Ingresa el nombre del taller.")
                else:
                    if not df_actividades.empty and n_act in df_actividades["ACTIVIDAD"].astype(str).str.upper().tolist():
                        df_actividades.loc[df_actividades["ACTIVIDAD"].astype(str).str.upper() == n_act, "MAESTRO"] = n_maestro
                        df_actividades.loc[df_actividades["ACTIVIDAD"].astype(str).str.upper() == n_act, "COLOR"] = n_color
                        df_actividades.loc[df_actividades["ACTIVIDAD"].astype(str).str.upper() == n_act, "ÁREA_IMPACTO"] = n_impacto
                        msg_c = f"Mantenimiento: Actualizar actividad {n_act}"
                    else:
                        nueva_act = pd.DataFrame([{"ACTIVIDAD": n_act, "MAESTRO": n_maestro, "COLOR": n_color, "ÁREA_IMPACTO": n_impacto}])
                        df_actividades = pd.concat([df_actividades, nueva_act], ignore_index=True)
                        msg_c = f"Mantenimiento: Registrar actividad {n_act}"
                        
                    with st.spinner("Guardando en GitHub..."):
                        if guardar_github(ACTIVIDADES_FILE, df_actividades, sha_actividades, msg_c):
                            st.success(f"🎉 Actividad guardada: {n_act}")
                            st.rerun()
                            
        with col_a2:
            st.markdown("##### 📋 Catálogo de Actividades")
            
            map_ods = {
                "Física / Motora": "ODS 3 (Salud y Bienestar)",
                "Cognitiva / Educativa": "ODS 4 (Educación de Calidad)",
                "Socio-Emocional": "ODS 10 (Reducción de Desigualdades)",
                "Autonomía / Vida Diaria": "ODS 11 (Comunidades Sostenibles)"
            }
            if not df_actividades.empty:
                df_act_vis = df_actividades.copy()
                df_act_vis["VINCULACIÓN ODS"] = df_act_vis["ÁREA_IMPACTO"].map(map_ods)
                st.dataframe(df_act_vis, use_container_width=True)
            else:
                st.info("No hay actividades registradas.")
                
            st.markdown("##### ❌ Eliminar Actividad del Catálogo")
            act_eliminar = st.selectbox("Selecciona la actividad a borrar:", df_actividades["ACTIVIDAD"].tolist() if not df_actividades.empty else [])
            if st.button("❌ ELIMINAR ACTIVIDAD"):
                if act_eliminar:
                    df_actividades = df_actividades[df_actividades["ACTIVIDAD"] != act_eliminar]
                    with st.spinner("Eliminando actividad..."):
                        if guardar_github(ACTIVIDADES_FILE, df_actividades, sha_actividades, f"Mantenimiento: Borrar actividad {act_eliminar}"):
                            st.success(f"✅ Actividad {act_eliminar} eliminada.")
                            st.rerun()

    # Pestaña 3: Catálogo de Maestros Talleristas
    with t_maestros:
        col_m1, col_m2 = st.columns([1, 2])
        with col_m1:
            st.markdown("##### ➕ Registrar / Modificar Maestro")
            with st.form("form_mantenimiento_maestro"):
                m_nombre = st.text_input("NOMBRE COMPLETO", placeholder="Ej: PROF. JUAN PÉREZ").strip().upper()
                m_tel = st.text_input("TELÉFONO DE CONTACTO", placeholder="Ej: (871) 123-4567")
                m_esp = st.text_input("ESPECIALIDAD / ÁREA", placeholder="Ej: Arte / Terapia Física")
                m_estado = st.selectbox("ESTADO", ["Activo", "Inactivo"])
                save_m = st.form_submit_button("💾 Guardar Ficha de Maestro")
                
            if save_m:
                if not m_nombre:
                    st.warning("⚠️ El nombre es obligatorio.")
                else:
                    if not df_maestros.empty and m_nombre in df_maestros["NOMBRE_COMPLETO"].astype(str).str.upper().tolist():
                        df_maestros.loc[df_maestros["NOMBRE_COMPLETO"].astype(str).str.upper() == m_nombre, "TELÉFONO"] = m_tel
                        df_maestros.loc[df_maestros["NOMBRE_COMPLETO"].astype(str).str.upper() == m_nombre, "ESPECIALIDAD"] = m_esp
                        df_maestros.loc[df_maestros["NOMBRE_COMPLETO"].astype(str).str.upper() == m_nombre, "ESTADO"] = m_estado
                        msg_c = f"Mantenimiento: Actualizar maestro {m_nombre}"
                    else:
                        nuevo_m = pd.DataFrame([{"NOMBRE_COMPLETO": m_nombre, "TELÉFONO": m_tel, "ESPECIALIDAD": m_esp, "ESTADO": m_estado}])
                        df_maestros = pd.concat([df_maestros, nuevo_m], ignore_index=True)
                        msg_c = f"Mantenimiento: Registrar maestro {m_nombre}"
                        
                    with st.spinner("Guardando en GitHub..."):
                        if guardar_github(MAESTROS_FILE, df_maestros, sha_maestros, msg_c):
                            st.success(f"🎉 Maestro guardado: {m_nombre}")
                            st.rerun()
                            
        with col_m2:
            st.markdown("##### 📋 Maestros Registrados")
            if not df_maestros.empty:
                st.dataframe(df_maestros, use_container_width=True)
            else:
                st.info("No hay maestros registrados.")
                
            st.markdown("##### ❌ Eliminar Maestro del Registro")
            m_eliminar = st.selectbox("Selecciona el maestro a borrar:", df_maestros["NOMBRE_COMPLETO"].tolist() if not df_maestros.empty else [])
            if st.button("❌ ELIMINAR MAESTRO"):
                if m_eliminar:
                    if m_eliminar == "—":
                        st.error("❌ No se puede eliminar el maestro por defecto '—'.")
                    else:
                        df_maestros = df_maestros[df_maestros["NOMBRE_COMPLETO"] != m_eliminar]
                        with st.spinner("Eliminando maestro..."):
                            if guardar_github(MAESTROS_FILE, df_maestros, sha_maestros, f"Mantenimiento: Borrar maestro {m_eliminar}"):
                                st.success(f"✅ Maestro {m_eliminar} eliminado.")
                                st.rerun()

    # Pestaña 3: Base de Datos Alumnos
    with t_alumnos:
        st.markdown("##### 👤 Edición y Eliminación Completa de Fichas de Integrantes")
        if not df_integrantes.empty:
            st.dataframe(df_integrantes, use_container_width=True)
            
            col_del1, col_del2 = st.columns([2, 1])
            with col_del2:
                st.markdown("<div style='background:rgba(220, 53, 69, 0.08); padding: 18px; border-radius: 10px; border: 1px solid rgba(220, 53, 69, 0.2);'>", unsafe_allow_html=True)
                st.markdown("##### ⚠️ Acción Crítica")
                st.write("Esta acción borrará de manera definitiva la ficha del alumno. No se borrará su historial de asistencia para preservar las estadísticas.")
                alumno_del = st.selectbox("Selecciona el alumno a eliminar:", sorted(df_integrantes["NOMBRE_COMPLETO"].tolist()))
                confirma_del = st.checkbox("☑️ Confirmo eliminar permanentemente a este integrante")
                if st.button("🚨 ELIMINAR INTEGRANTE DEFINITIVAMENTE", use_container_width=True):
                    if confirma_del:
                        df_integrantes = df_integrantes[df_integrantes["NOMBRE_COMPLETO"] != alumno_del]
                        with st.spinner("Eliminando ficha..."):
                            if guardar_github(INTEGRANTES_FILE, df_integrantes, sha_integrantes, f"Mantenimiento: Borrar alumno {alumno_del}"):
                                st.success(f"✅ Alumno {alumno_del} eliminado de la base de datos.")
                                st.rerun()
                    else:
                        st.error("⚠️ Debes marcar la casilla de confirmación.")
                st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.info("No hay alumnos registrados en la base de datos.")

    # Pestaña 4: Herramientas de Datos (Sembrado y Limpieza)
    with t_herramientas:
        col_tool1, col_tool2 = st.columns(2)
        
        with col_tool1:
            st.markdown("<div style='background:rgba(234, 181, 25, 0.08); padding:20px; border-radius:10px; border:1px solid rgba(234,181,25,0.2); height:100%;'>", unsafe_allow_html=True)
            st.markdown("### 🧬 Inyectar Registros de Prueba (Semillero)")
            st.write("Genera datos ficticios aleatorios y realistas correspondientes a los **últimos 30 días**.")
            st.write("Esto inyectará tanto citas programadas (`calendario.xlsx`) como registros de asistencia (`asistencia.xlsx`) para poder evaluar el comportamiento de los dashboards e informes.")
            
            dias_sembrar = st.slider("Días de historial a sembrar:", 5, 30, 20)
            if st.button("⚙️ INYECTAR REGISTROS DE PRUEBA"):
                fechas_sembrar = []
                for d in range(dias_sembrar):
                    fecha_val = (datetime.now() - timedelta(days=d))
                    if fecha_val.weekday() != 6:
                        fechas_sembrar.append(fecha_val.strftime("%d/%m/%Y"))
                
                import random
                registros_citas = []
                registros_asistencias = []
                
                talleres_disponibles = df_actividades["ACTIVIDAD"].tolist() if not df_actividades.empty else []
                alumnos_activos = df_integrantes[df_integrantes["ESTADO"] == "Activo"]["NOMBRE_COMPLETO"].tolist() if not df_integrantes.empty else []
                
                if not talleres_disponibles or not alumnos_activos:
                    st.error("❌ Necesitas tener al menos una actividad y un integrante activo en la base de datos para sembrar.")
                else:
                    for fec in fechas_sembrar:
                        talleres_dia = random.sample(talleres_disponibles, min(3, len(talleres_disponibles)))
                        for tal in talleres_dia:
                            h_propuesto = random.choice(["09:00 - 10:30", "11:00 - 12:30", "13:00 - 14:30"])
                            num_alumnos = random.randint(min(5, len(alumnos_activos)), min(12, len(alumnos_activos)))
                            alumnos_dia = random.sample(alumnos_activos, num_alumnos)
                            
                            for al in alumnos_dia:
                                registros_citas.append({
                                    COL_FECHA: fec,
                                    COL_ASISTENCIA: al,
                                    COL_TALLER: tal,
                                    COL_HORAS: 1.5,
                                    "HORARIO": h_propuesto
                                })
                                if random.random() < 0.85:
                                    registros_asistencias.append({
                                        COL_FECHA: fec,
                                        COL_ASISTENCIA: al,
                                        COL_TALLER: f"{tal} ({h_propuesto})",
                                        COL_HORAS: 1.5
                                    })
                    
                    df_cal_sembrado = pd.DataFrame(registros_citas)
                    df_asist_sembrado = pd.DataFrame(registros_asistencias)
                    
                    with st.spinner("Inyectando registros en GitHub..."):
                        exito_cal = guardar_github(CALENDARIO_FILE, df_cal_sembrado, sha_calendario, "Semillero: Inyectar registros de calendario")
                        st.spinner("Actualizando asistencia...")
                        df_original, archivo_sha, df_calendario, sha_calendario, df_integrantes, sha_integrantes, df_no_laborables, sha_no_laborables, df_usuarios, sha_usuarios, df_actividades, sha_actividades, df_maestros, sha_maestros = cargar_datos_sistema()
                        exito_asist = guardar_github(EXCEL_FILE, df_asist_sembrado, archivo_sha, "Semillero: Inyectar registros de asistencia")
                        
                        if exito_cal and exito_asist:
                            st.success(f"🎉 ¡Inyección finalizada con éxito! Sembrados {len(df_cal_sembrado)} citas y {len(df_asist_sembrado)} asistencias.")
                            st.balloons()
                            st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            
        with col_tool2:
            st.markdown("<div style='background:rgba(220, 53, 69, 0.08); padding:20px; border-radius:10px; border:1px solid rgba(220,53,69,0.2); height:100%;'>", unsafe_allow_html=True)
            st.markdown("### 🧹 Limpiar Registros Históricos")
            st.write("Esta herramienta **borra por completo** todo el historial de asistencias guardadas y programadas.")
            st.write("Regresa las bases de datos a cero:")
            st.write("* Limpia `asistencia.xlsx` (asistencias históricas).")
            st.write("* Limpia `calendario.xlsx` (citas planeadas).")
            st.markdown("<strong style='color:#dc3545;'>⚠️ SE PRESERVA INTACTO EL CATÁLOGO DE INTEGRANTES, ACTIVIDADES, DOCENTES Y USUARIOS.</strong>", unsafe_allow_html=True)
            
            confirma_limpieza = st.checkbox("☑️ Confirmo que deseo vaciar todo el historial del calendario y asistencias.")
            if st.button("🚨 EJECUTAR LIMPIEZA TOTAL"):
                if confirma_limpieza:
                    df_cal_vacio = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS, "HORARIO"])
                    df_asist_vacia = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS])
                    
                    with st.spinner("Limpiando base de datos..."):
                        exito_cal = guardar_github(CALENDARIO_FILE, df_cal_vacio, sha_calendario, "Limpieza total de calendario")
                        df_original, archivo_sha, df_calendario, sha_calendario, df_integrantes, sha_integrantes, df_no_laborables, sha_no_laborables, df_usuarios, sha_usuarios, df_actividades, sha_actividades, df_maestros, sha_maestros = cargar_datos_sistema()
                        exito_asist = guardar_github(EXCEL_FILE, df_asist_vacia, archivo_sha, "Limpieza total de asistencias")
                        
                        if exito_cal and exito_asist:
                            st.success("🧹 Sistema restablecido a cero con éxito. Catálogos preservados.")
                            st.rerun()
                else:
                    st.error("⚠️ Debes marcar la casilla de confirmación.")
            st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="mca-footer">
    🦅 &nbsp;<strong>MENTES CON ALAS</strong> — Comunidad de Vida para Adultos con Parálisis Cerebral
    &nbsp;|&nbsp; Av. Ocampo 1797 Ote., Torreón, Coahuila
    &nbsp;|&nbsp; <a href="mailto:comunidad@mentesconalas.org.mx">comunidad@mentesconalas.org.mx</a>
    &nbsp;|&nbsp; <a href="https://www.mentesconalas.org.mx" target="_blank">mentesconalas.org.mx</a>
    <br><span style="font-size:10px;opacity:0.5;">Sistema de Control de Asistencia · {datetime.now().year}</span>
</div>
""", unsafe_allow_html=True)
