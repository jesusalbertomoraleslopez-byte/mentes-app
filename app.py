import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
from github import Github
import io
import calendar

# Configuración estética de la pestaña del navegador
st.set_page_config(page_title="Mentes Con Alas - Asistencia", page_icon="🦅", layout="centered")

EXCEL_FILE = "asistencia.xlsx"
CALENDARIO_FILE = "calendario.xlsx"
URL_LOGO = "logo-mentes.png"

# --- INYECTAR DISEÑO VISUAL INSPIRADO EN LA WEB OFICIAL ---
st.markdown("""
    <style>
        html, body, [data-testid="stAppViewContainer"] {
            background-color: #FAFAFA;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        .titulo-web { text-align: center; color: #0A2540; font-size: 32px; font-weight: 700; margin-bottom: 5px; }
        .subtitulo-web { text-align: center; color: #627D98; font-size: 16px; margin-bottom: 20px; }
        h3 { color: #0A2540 !important; font-weight: 600 !important; margin-top: 25px !important; }
        div[data-testid="stExpander"] {
            border: 1px solid #E4E7EB !important;
            border-radius: 8px !important;
            background-color: #FFFFFF !important;
            box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.02) !important;
        }
        div[data-testid="stCheckbox"] label p { color: #0A2540 !important; font-weight: 700 !important; font-size: 15px !important; }
        div[data-testid="stExpander"] p, div[data-testid="stExpander"] span { color: #0A2540 !important; font-weight: 500 !important; }
        div[data-testid="stNotification"] { background-color: #FFEEEE !important; border: 2px solid #E53E3E !important; border-radius: 8px !important; }
        div[data-testid="stNotification"] p, div[data-testid="stNotification"] span { color: #9B1C1C !important; font-weight: 700 !important; font-size: 16px !important; }
        .contenedor-asistencia { background-color: #FFFFFF !important; padding: 20px !important; border-radius: 8px !important; border: 1px solid #E4E7EB !important; margin-top: 10px !important; margin-bottom: 10px !important; }
        .stButton > button, div[data-testid="stForm"] button { background-color: #0A2540 !important; color: white !important; border-radius: 6px !important; border: none !important; padding: 10px 24px !important; font-weight: 600 !important; width: 100% !important; }
        div[data-testid="stForm"] button p, .stButton button p { color: #FFFFFF !important; font-weight: 700 !important; }
        input, select, div[data-baseweb="select"] { border-radius: 6px !important; }
        .footer-web { text-align: center; font-size: 13px; color: #829AB1; margin-top: 50px; line-height: 1.6; width: 100%; }
        .footer-web a { color: #0A2540 !important; text-decoration: none; font-weight: bold; }
        .boton-borrado-html { background-color: #FFFFFF !important; color: #D32F2F !important; border: 3px solid #D32F2F !important; padding: 14px 20px !important; text-align: center !important; font-size: 16px !important; font-weight: 800 !important; border-radius: 8px !important; width: 100% !important; text-transform: uppercase !important; box-shadow: 0px 4px 6px rgba(0,0,0,0.1) !important; }
        
        /* 📅 CALENDARIO RESPONSIVE ADAPTADO A CELULARES 📅 */
        .tabla-calendario { width: 100%; border-collapse: collapse; background-color: #FFFFFF; border-radius: 8px; overflow: hidden; box-shadow: 0px 2px 8px rgba(0,0,0,0.05); table-layout: fixed; }
        .tabla-calendario th { background-color: #0A2540; color: #FFFFFF; padding: 10px 5px; font-weight: 700; text-align: center; font-size: 13px; border: 1px solid #1F3B5C; }
        .tabla-calendario td { height: 95px; vertical-align: top; padding: 6px 4px; border: 1px solid #E4E7EB; position: relative; background-color: #FFFFFF; overflow: hidden; }
        .num-dia { font-weight: 700; color: #627D98; font-size: 13px; margin-bottom: 3px; display: block; }
        .dia-vacio { background-color: #F8FAFC !important; }
        .dia-hoy { background-color: #EFF6FF !important; border: 2px solid #3B82F6 !important; }
        .evento-tag { background-color: #E0F2FE !important; color: #0369A1 !important; border-left: 3px solid #0284C7 !important; padding: 2px 4px !important; font-size: 10px !important; font-weight: 700 !important; border-radius: 3px !important; margin-top: 3px !important; line-height: 1.1 !important; white-space: nowrap !important; text-overflow: ellipsis !important; overflow: hidden !important; }
    </style>
""", unsafe_allow_html=True)

# --- LISTAS OFICIALES E INTEGRANTES ---
INTEGRANTES_FIJOS = [
    "ANA DE LOS ÁNGELES TORRES REVELES", "CARLOS ALBERTO DE LA TORRE SANTELLANO", "CRISTABEL DE LA CRUZ MALDONADO",
    "EFRAÍN MAYNEZ PORRAS", "FERNANDO ÁVILA BERLANGA", "FLORINDA ESTEVANÉ PIZARRO", "GUADALUPE LÓPEZ TOVAR",
    "ISAAC IGNACIO GONZÁLEZ CRUZ", "JESÚS ALEJANDRO SIFUENTES ESPINO", "JESÚS SALCIDO ZAMORA",
    "JOSE REYNALDO ALCORTA BENAVIDES", "JUAN JOSÉ OCHOA GONZÁLEZ", "JUAN RAFAEL YAÑEZ SERNA",
    "JUAN SILVERIO LÓPEZ DE LA ROSA", "KARLA LISSETTE PEDROZA GONZÁLEZ", "LUIS FERNANDO ÁVILA BERLANGA",
    "LUIS JAVIER GARCÍA MARTÍNEZ", "MA. ELIZABETH GONZÁLEZ HIDROGO", "MARÍA DE LOS ÁNGELES ORDUÑA RODRÍGUEZ",
    "MARÍA GUADALUPE DE LA CONCEPCIÓN TORRES LIRA", "PEDRO ANTONIO DE LA CERDA TREVIÑO",
    "TERESITA DEL NIÑO JESÚS RODRÍGUEZ SALAZAR", "TOMASITA MARÍA ENRIQUETA RIVERA DEL BOSQUE"
]

TALLERES_FIJOS = [
    "AJEDREZ", "ARTE", "DESARROLLO EDUCATIVO", "FELDENKRAIS", "GRUPO DE CRECIMIENTO",
    "MOVIMIENTO VITAL EXPRESIVO", "TALLER \"SÍ PUEDO\"", "TALLER DE COMUNICACIÓN", "TEATRO", "VIDA DIARIA"
]

COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS, CLAVE_BORRADO = "FECHA", "INTEGRANTE / TALLER", "ACTIVIDAD", "HORAS", "LupitaMentes1978"

def conectar_github():
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        return g.get_repo(st.secrets["GITHUB_REPO"])
    except Exception as e:
        st.error(f"❌ Error de autenticación con GitHub: {e}"); st.stop()

def cargar_datos_sistema():
    repo = conectar_github()
    try:
        f_asistencia = repo.get_contents(EXCEL_FILE)
        df_asistencia = pd.read_excel(io.BytesIO(f_asistencia.decoded_content))
        sha_asistencia = f_asistencia.sha
    except Exception:
        df_asistencia = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS])
        sha_asistencia = None
    try:
        f_calendario = repo.get_contents(CALENDARIO_FILE)
        df_calendario = pd.read_excel(io.BytesIO(f_calendario.decoded_content))
        sha_calendario = f_calendario.sha
    except Exception:
        df_calendario = pd.DataFrame(columns=[COL_FECHA, COL_ASISTENCIA, COL_TALLER, COL_HORAS, "HORARIO"])
        sha_calendario = None
    return df_asistencia, sha_asistencia, df_calendario, sha_calendario

df_original, archivo_sha, df_calendario, sha_calendario = cargar_datos_sistema()

def guardar_archivo_github(nombre_archivo, df_nuevo, sha_actual, mensaje_commit):
    repo = conectar_github()
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_nuevo.to_excel(writer, index=False)
        if sha_actual:
            repo.update_file(nombre_archivo, mensaje_commit, output.getvalue(), sha_actual)
        else:
            repo.create_file(nombre_archivo, mensaje_commit, output.getvalue())
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar {nombre_archivo}: {e}"); return False

# --- LOGOTIPO ---
col_logo_1, col_logo_2, col_logo_3 = st.columns([1, 0.4, 1])
with col_logo_2: st.image(URL_LOGO, use_container_width=True)

st.markdown('<div class="titulo-web">Control de Asistencia Grupal</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo-web">Comunidad de Vida para Adultos con Parálisis Cerebral</div>', unsafe_allow_html=True)

# 🚨 SISTEMA DE 4 PESTAÑAS INTEGRALES SOLICITADAS 🚨
tab_seccion1, tab_seccion2, tab_seccion3, tab_seccion4 = st.tabs([
    "📅 1. CONTROL DE CALENDARIO", 
    "🗺️ 2. AGENDA MENSUAL Y DASHBOARD", 
    "📝 3. PASE DE LISTA REAL", 
    "⚖️ 4. COMPARATIVA PROGRAMADO VS REAL"
])

# =======================================================
# PESTAÑA 1: PROGRAMAR Y MODIFICAR CALENDARIO (PLANEACIÓN)
# =======================================================
with tab_seccion1:
    st.markdown("### 🗓️ Programación de Nuevas Actividades")
    with st.form("form_crear_agenda"):
        p_fecha = st.date_input("FECHA PROGRAMADA", datetime.now())
        p_taller = st.selectbox("TALLER A IMPARTIR", TALLERES_FIJOS)
        c1, c2 = st.columns(2)
        with c1: p_ini = st.time_input("HORA DE INICIO", time(9, 0))
        with c2: p_dur = st.number_input("CANTIDAD DE HORAS (DURACIÓN)", min_value=0.25, max_value=8.0, value=1.5, step=0.25)
        
        min_totales = int(p_dur * 60)
        dt_fin = datetime.combine(p_fecha, p_ini) + timedelta(minutes=min_totales)
        h_completo = f"{p_ini.strftime('%H:%M')} - {dt_fin.strftime('%H:%M')}"
        st.info(f"⏰ Horario calculado: {h_completo}")
        
        st.markdown("### 👥 Selecciona los Integrantes Citados:")
        st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
        col_c_izq, col_c_der = st.columns(2)
        estados_cal = {}
        for i, nombre in enumerate(INTEGRANTES_FIJOS):
            with col_c_izq if i % 2 == 0 else col_c_der:
                estados_cal[nombre] = st.checkbox(nombre, value=False, key=f"cal_{nombre}")
        st.markdown('</div>', unsafe_allow_html=True)
        boton_add_cal = st.form_submit_button("🗓️ Guardar Programación de Actividad")

    if boton_add_cal:
        citados = [n for n, c in estados_cal.items() if c]
        if not citados: st.warning("⚠️ Elige al menos un integrante.")
        else:
            f_str = p_fecha.strftime("%d/%m/%Y")
            filas = [{COL_FECHA: f_str, COL_ASISTENCIA: n, COL_TALLER: p_taller, COL_HORAS: float(p_dur), "HORARIO": h_completo} for n in citados]
            df_n_cal = pd.concat([df_calendario, pd.DataFrame(filas)], ignore_index=True)
            if guardar_archivo_github(CALENDARIO_FILE, df_n_cal, sha_calendario, f"Agenda: {p_taller}"):
                st.success("✨ ¡Actividad agendada correctamente!"); st.rerun()

    st.markdown("---")
    st.markdown("### ✏️ Modificar o Corregir una Actividad Agendada (Abierto Libre)")
    if not df_calendario.empty:
        fechas_cal = sorted(df_calendario[COL_FECHA].dropna().unique().tolist(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"), reverse=True)
        m_fecha = st.selectbox("1. Elige la fecha que deseas corregir:", fechas_cal, key="m_fecha_sel")
        df_m_dia = df_calendario[df_calendario[COL_FECHA] == m_fecha].copy()
        df_m_dia['BLOQUE'] = df_m_dia[COL_TALLER].astype(str) + " (" + df_m_dia["HORARIO"].astype(str) + ")"
        bloques_m = sorted(df_m_dia['BLOQUE'].dropna().unique().tolist())
        m_bloque = st.selectbox("2. Elige la actividad específica a modificar:", bloques_m, key="m_bloque_sel")
        
        df_bloque_especifico = df_m_dia[df_m_dia['BLOQUE'] == m_bloque]
        if not df_bloque_especifico.empty:
            taller_m = df_bloque_especifico[COL_TALLER].iloc[0]
            horas_m = df_bloque_especifico[COL_HORAS].iloc[0]
            horario_m = df_bloque_especifico["HORARIO"].iloc[0]
            alumnos_citados_actualmente = df_bloque_especifico[COL_ASISTENCIA].dropna().tolist()
            
            st.markdown("### 3. Modifica los convocados (Marca para añadir, desmarca para quitar):")
            st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
            col_m_izq, col_m_der = st.columns(2)
            nuevos_estados_cal = {}
            for i, nombre in enumerate(INTEGRANTES_FIJOS):
                ya_citado = nombre in alumnos_citados_actualmente
                with col_m_izq if i % 2 == 0 else col_m_der:
                    nuevos_estados_cal[nombre] = st.checkbox(nombre, value=ya_citado, key=f"mod_c_{nombre}")
            st.markdown('</div>', unsafe_allow_html=True)
            
            boton_modificar_agenda = st.button("💾 Guardar Cambios en la Agenda")
            if boton_modificar_agenda:
                alumnos_finales = [n for n, c in nuevos_estados_cal.items() if c]
                df_cal_limpio = df_calendario[~((df_calendario[COL_FECHA] == m_fecha) & (df_calendario[COL_TALLER] == taller_m) & (df_calendario["HORARIO"] == horario_m))]
                if alumnos_finales:
                    nuevas_filas_m = [{COL_FECHA: m_fecha, COL_ASISTENCIA: n, COL_TALLER: taller_m, COL_HORAS: float(horas_m), "HORARIO": horario_m} for n in alumnos_finales]
                    df_cal_finalizado = pd.concat([df_cal_limpio, pd.DataFrame(nuevas_filas_m)], ignore_index=True)
                else:
                    df_cal_finalizado = df_cal_limpio
                if guardar_archivo_github(CALENDARIO_FILE, df_cal_finalizado, sha_calendario, f"Calendario: Modificación bloque {taller_m}"):
                    st.success("🎉 ¡Agenda actualizada con éxito!"); st.rerun()
    else: st.info("💡 No hay actividades planificadas en el calendario para editar.")

# =======================================================
# PESTAÑA 2: CALENDARIO MENSUAL GRÁFICO Y DASHBOARD
# =======================================================
with tab_seccion2:
    st.markdown("### 📅 Agenda Mensual de Actividades")
    hoy = datetime.now()
    c_m1, c_m2 = st.columns(2)
    with c_m1:
        mes_lista = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        mes_sel = st.selectbox("Ver Mes:", mes_lista, index=hoy.month - 1)
        mes_num = mes_lista.index(mes_sel) + 1
    with c_m2: anio_sel = st.selectbox("Ver Año:", [hoy.year, hoy.year + 1], index=0)
        
    cal_objeto = calendar.Calendar(firstweekday=0)
    semanas_matriz = cal_objeto.monthdayscalendar(anio_sel, mes_num)
    
    diccionario_eventos = {}
    if not df_calendario.empty:
        df_unicos_cal = df_calendario.drop_duplicates(subset=[COL_FECHA, COL_TALLER, "HORARIO"])
        for _, fila in df_unicos_cal.iterrows():
            try:
                f_obj = datetime.strptime(str(fila[COL_FECHA]), "%d/%m/%Y")
                if f_obj.month == mes_num and f_obj.year == anio_sel:
                    dia_key = f_obj.day
                    if dia_key not in diccionario_eventos: diccionario_eventos[dia_key] = []
                    diccionario_eventos[dia_key].append(f"<b>{fila[COL_TALLER]}</b><br>⏱️ {fila['HORARIO']}")
            except: pass

    html_codigo = f"<table class='tabla-calendario'><tr><th>Lun</th><th>Mar</th><th>Mié</th><th>Jue</th><th>Vie</th><th>Sáb</th><th>Dom</th></tr>"
    for semana in semanas_matriz:
        html_codigo += "<tr>"
        for dia in semana:
            if dia == 0: html_codigo += "<td class='dia-vacio'></td>"
            else:
                clase_celda = "class='dia-hoy'" if (dia == hoy.day and mes_num == hoy.month and anio_sel == hoy.year) else ""
                html_codigo += f"<td {clase_celda}><span class='num-dia'>{dia}</span>"
                if dia in diccionario_eventos:
                    for ev in diccionario_eventos[dia]: html_codigo += f"<div class='evento-tag'>{ev}</div>"
                html_codigo += "</td>"
        html_codigo += "</tr>"
    html_codigo += "</table>"
    st.markdown(html_codigo, unsafe_allow_html=True)

    # --- NUEVO DASHBOARD ESTADÍSTICO DE RENDIMIENTO (SOLICITADO) ---
    st.markdown("---")
    st.markdown("### 📊 Dashboard Analítico de Asistencias Realizadas")
    if not df_original.empty:
        df_stats = df_original.copy()
        # Sumar el total de horas reales acumuladas por cada participante
        df_horas_totales = df_stats.groupby(COL_ASISTENCIA)[COL_HORAS].sum().reset_index()
        df_horas_totales.columns = ["Integrante", "Horas Totales"]
        df_horas_totales = df_horas_totales.sort_values(by="Horas Totales", ascending=False)
        
        st.write("Visualiza la cantidad total de horas de terapia y talleres completadas por cada integrante:")
        st.bar_chart(data=df_horas_totales.set_index("Integrante"), y="Horas Totales", use_container_width=True)
    else: st.info("💡 El Dashboard se pintará automáticamente cuando guardes los primeros registros de asistencia real.")

            st.error("❌ Clave incorrecta de administración.")

# --- PIE DE PÁGINA
        if btn_save_cal:
            asistieron = [n for n, c in checks_p.items() if c]
            if not asistieron: st.warning("⚠️ Registra al menos una asistencia.")
            else:
                filas_asist = [{COL_FECHA: fecha_lista_buscar, COL_ASISTENCIA: n, COL_TALLER: act_con_h, COL_HORAS: float(horas_p)} for n in asistieron]
                df_f_asist = pd.concat([df_original, pd.DataFrame(filas_asist)], ignore_index=True)
                if guardar_archivo_github(EXCEL_FILE, df_f_asist, archivo_sha, f"Asistencia: {act_con_h}"):
                    st.success("🎉 ¡Asistencia del bloque guardada con éxito!"); st.rerun()
else:
    with st.form("form_real_dir"):
        c_d1, c_d2 = st.columns(2)
        with c_d1: t_dir = st.selectbox("TALLER IMPARTIDO", TALLERES_FIJOS)
        with c_d2: h_dir = st.number_input("HORAS DEL TALLER", min_value=0.25, max_value=8.0, value=1.5, step=0.25)
        st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
        col_rd_izq, col_rd_der = st.columns(2)
        checks_d = {}
        for i, nombre in enumerate(INTEGRANTES_FIJOS):
            with col_rd_izq if i % 2 == 0 else col_rd_der:
                checks_d[nombre] = st.checkbox(nombre, value=True, key=f"chk_dir_{nombre}")
        st.markdown('</div>', unsafe_allow_html=True)
        btn_save_dir = st.form_submit_button("💾 Registrar Asistencia (Directo)")
        
    if btn_save_dir:
        asistieron_d = [n for n, c in checks_d.items() if c]
        if not asistieron_d: st.warning("⚠️ Selecciona al menos un integrante.")
        else:
            filas_dir = [{COL_FECHA: fecha_lista_buscar, COL_ASISTENCIA: n, COL_TALLER: t_dir, COL_HORAS: float(h_dir)} for n in asistieron_d]
            df_f_dir = pd.concat([df_original, pd.DataFrame(filas_dir)], ignore_index=True)
            if guardar_archivo_github(EXCEL_FILE, df_f_dir, archivo_sha, f"Directo: {t_dir}"):
                st.success("🎉 ¡Asistencia guardada directamente!"); st.rerun()

st.markdown("---")
st.markdown("### 📋 Historial y Buscador de Asistencias Realizadas")
if not df_original.empty:
    df_vis = df_original.copy()
    try:
        df_vis['FECHA_DT'] = pd.to_datetime(df_vis[COL_FECHA], format="%d/%m/%Y", errors='coerce')
        df_vis = df_vis.sort_values(by='FECHA_DT', ascending=False).drop(columns=['FECHA_DT'])
    except Exception: pass
    
    filtro = st.text_input("🔍 Filtro en tiempo real (Fecha, Nombre, Actividad):").strip().upper()
    if filtro:
        mask = df_vis[COL_FECHA].astype(str).str.upper().str.contains(filtro) | df_vis[COL_ASISTENCIA].astype(str).str.upper().str.contains(filtro) | df_vis[COL_TALLER].astype(str).str.upper().str.contains(filtro)
        df_f_tabla = df_vis[mask]
    else: df_f_tabla = df_vis
    
    st.dataframe(df_f_tabla, use_container_width=True)
    output_d = io.BytesIO()
    with pd.ExcelWriter(output_d, engine='openpyxl') as writer:
        df_vis.sort_values(by=COL_FECHA, ascending=True).to_excel(writer, index=False)
    st.download_button(label="📥 Descargar Base de Datos Completa de Asistencias (Excel)", data=output_d.getvalue(), file_name=f"asistencia_completa.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
else: st.info("💡 Archivo de asistencia vacío.")

with st.expander("🚨 Panel de Administración - Control de Historial", expanded=False):
    clave = st.text_input("Contraseña del Administrador:", type="password")
    if clave == CLAVE_BORRADO:
        if not df_original.empty:
            fechas_u = sorted(df_original[COL_FECHA].dropna().unique(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"), reverse=True)
            f_sel = st.selectbox("Fecha error:", fechas_u)
            df_dia_b = df_original[df_original[COL_FECHA] == f_sel].copy()
            
            st.markdown('<div class="contenedor-asistencia">', unsafe_allow_html=True)
            b_regs = []
            for idx, fila in df_dia_b.iterrows():
                if st.checkbox(f"👤 {fila[COL_ASISTENCIA]} | 📚 {fila[COL_TALLER]}", value=False, key=f"del_{idx}"): b_regs.append(idx)
            st.markdown('</div>', unsafe_allow_html=True)
            
            if b_regs:
                st.error(f"Se eliminarán {len(b_regs)} filas.")
                if st.checkbox("Confirmar acción", value=False):
                    st.markdown('<button class="boton-borrado-html">❌ CONFIRMAR ELIMINACIÓN DE REGISTROS</button>', unsafe_allow_html=True)
                    df_res = df_original.drop(index=b_regs)
                    if guardar_archivo_github(EXCEL_FILE, df_res, archivo_sha, "Admin: Borrado"):
                        st.success("Eliminado con éxito!"); st.rerun()

# ==========================================================
# SECCIÓN 5: COMPARATIVA ANALÍTICA (PROGRAMADO VS REAL)
# ==========================================================
with tab5:
    st.markdown("### ⚖️ Conciliación: Planeación del Calendario vs Asistencia Real")
    st.write("Esta herramienta cruza los datos de tus planes contra lo registrado para identificar de forma exacta asistencias y ausencias injustificadas.")
    
    if not df_calendario.empty:
        fechas_disponibles_cal = sorted(df_calendario[COL_FECHA].dropna().unique().tolist(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"), reverse=True)
        fecha_evaluar_comp = st.selectbox("Selecciona la fecha que deseas conciliar:", fechas_disponibles_cal, key="fecha_evaluar_comp_key")
        
        # Filtrar planeación del día
        df_cal_dia = df_calendario[df_calendario[COL_FECHA] == fecha_evaluar_comp].copy()
        df_cal_dia['BLOQUE_FUSION'] = df_cal_dia[COL_TALLER].astype(str) + " (" + df_cal_dia["HORARIO"].astype(str) + ")"
        bloques_dia_cal = sorted(df_cal_dia['BLOQUE_FUSION'].unique().tolist())
        
        bloque_evaluar_comp = st.selectbox("Selecciona el bloque específico programado para auditar:", bloques_dia_cal, key="bloque_evaluar_comp_key")
        
        # Alumnos que debieron ir según calendario
        citados_lista = df_cal_dia[df_cal_dia['BLOQUE_FUSION'] == bloque_evaluar_comp][COL_ASISTENCIA].dropna().unique().tolist()
        
        # Alumnos que realmente asistieron según historial
        asistidos_lista = []
        if not df_original.empty:
            df_asist_dia = df_original[df_original[COL_FECHA] == fecha_evaluar_comp]
            # Filtrar por actividad aproximada o exacta
            df_asist_bloque = df_asist_dia[df_asist_dia[COL_TALLER] == bloque_evaluar_comp]
            asistidos_lista = df_asist_bloque[COL_ASISTENCIA].dropna().unique().tolist()
            
        # Realizar el cruce de datos analítico
        logrados = [alumno for alumno in citados_lista if alumno in asistidos_lista]
        ausentes = [alumno for alumno in citados_lista if alumno not in asistidos_lista]
        
        st.markdown("---")
        c_comp1, c_comp2 = st.columns(2)
        
        with c_comp1:
            st.markdown(f"#### ✅ Integrantes Cumplidos ({len(logrados)})")
            st.write("Alumnos programados que cuentan con pase de lista registrado:")
            if logrados:
                for al in logrados: st.write(f"• {al}")
            else:
                st.caption("Ningún alumno citado cuenta con asistencia asentada.")
                
        with c_comp2:
            st.markdown(f"#### ❌ Integrantes Ausentes / Faltas ({len(ausentes)})")
            st.write("Alumnos programados que NO tienen registro de asistencia en este bloque:")
            if ausentes:
                for al in ausentes:
                    st.markdown(f"<span style='color:#CC0000; font-weight:bold;'>• {al}</span>", unsafe_allow_html=True)
            else:
                st.caption("🎉 ¡Asistencia perfecta! Todos los programados asistieron.")
    else:
        st.info("💡 El archivo de calendario está vacío. No hay datos que conciliar.")

# --- PIE DE PÁGINA INSTITUCIONAL ---
st.markdown("""
    <div class="footer-web">
        <hr>
        Av. Ocampo 1797 ote. C.P. 27000, Col. Centro Torreón, Coah.<br>
        <b>MENTES CON ALAS ES DONATARIA AUTORIZADA</b><br>
        <a href="https://mentesconalas.org.mx" target="_blank">🌐 Visitar Sitio Web Oficial</a>
    </div>
""", unsafe_allow_html=True)

