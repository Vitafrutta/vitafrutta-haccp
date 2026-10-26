import streamlit as st
import pandas as pd
import os
from datetime import datetime

# 1. Obter caminho absoluto do diretório do script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "Sistema_HACCP_Master_11Modulos.xlsx")

# 2. Configuração de Página Executiva
st.set_page_config(
    page_title="Vitafrutta Quality Suite | HACCP",
    page_icon="🍋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 3. Estilização CSS com Alto Contraste e Design Responsivo
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    .stApp {
        background-color: #f3f7f5 !important;
    }

    /* Forçar contraste de texto escuro */
    p, span, label, li, ul, div, h1, h2, h3, h4, h5, h6 {
        color: #11231E !important;
    }

    /* Painel Lateral Esquerdo */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #00382B 0%, #00221B 100%) !important;
        border-right: 3px solid #76BC21 !important;
    }

    .sidebar-logo-card {
        background: #ffffff !important;
        padding: 16px 14px;
        border-radius: 16px;
        border-bottom: 4px solid #76BC21;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.25);
        text-align: center;
        margin-bottom: 20px;
    }

    section[data-testid="stSidebar"] .stSelectbox label p,
    section[data-testid="stSidebar"] .stSelectbox label {
        color: #a7f36a !important;
        font-weight: 800 !important;
        font-size: 0.85rem !important;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }

    div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border: 2px solid #76BC21 !important;
        border-radius: 12px !important;
        color: #004D3D !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
    }

    div[data-baseweb="select"] * {
        color: #004D3D !important;
        white-space: normal !important;
        word-wrap: break-word !important;
    }

    .plant-status-card {
        background: rgba(255, 255, 255, 0.08) !important;
        backdrop-filter: blur(8px);
        padding: 14px;
        border-radius: 14px;
        border: 1px solid rgba(118, 188, 33, 0.3);
        border-left: 5px solid #76BC21;
        margin-top: 25px;
    }
    .plant-status-title {
        font-size: 0.8rem;
        font-weight: 800;
        color: #76BC21 !important;
        text-transform: uppercase;
    }
    .plant-status-text {
        font-size: 0.76rem;
        color: #e2f5d5 !important;
        margin-top: 4px;
    }

    /* Banner Principal */
    .main-banner {
        background: linear-gradient(135deg, #004D3D 0%, #002B22 100%) !important;
        padding: 20px 26px;
        border-radius: 18px;
        color: white !important;
        box-shadow: 0 8px 24px rgba(0, 77, 61, 0.18);
        margin-bottom: 22px;
        border-bottom: 4px solid #76BC21;
    }
    .main-banner-title {
        color: #ffffff !important;
        font-size: 1.6rem;
        font-weight: 800;
        margin: 0;
    }
    .main-banner-subtitle {
        color: #a7f36a !important;
        margin-top: 4px;
        font-size: 0.98rem;
        font-weight: 600;
    }
    .live-badge {
        background: rgba(118, 188, 33, 0.2) !important;
        color: #ffffff !important;
        padding: 5px 14px;
        border-radius: 50px;
        font-size: 0.74rem;
        font-weight: 700;
        border: 1px solid #76BC21;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        margin-top: 10px;
    }
    .pulse-green {
        width: 7px;
        height: 7px;
        background-color: #76BC21;
        border-radius: 50%;
        box-shadow: 0 0 8px #76BC21;
    }

    /* Cartões KPI */
    div[data-testid="stMetric"] {
        background-color: #ffffff !important;
        padding: 12px 14px !important;
        border-radius: 14px !important;
        border: 1px solid #e1efe9 !important;
        border-top: 4px solid #76BC21 !important;
        box-shadow: 0 4px 14px rgba(0, 77, 61, 0.05) !important;
    }
    div[data-testid="stMetricLabel"], 
    div[data-testid="stMetricLabel"] p {
        font-size: 0.76rem !important;
        font-weight: 800 !important;
        color: #004D3D !important;
        text-transform: uppercase;
        white-space: normal !important;
        word-wrap: break-word !important;
    }
    div[data-testid="stMetricValue"], 
    div[data-testid="stMetricValue"] div {
        font-size: 1.25rem !important;
        font-weight: 800 !important;
        color: #0f172a !important;
        white-space: normal !important;
        word-wrap: break-word !important;
    }

    /* Formulários e Rádios */
    label[data-testid="stWidgetLabel"] p, 
    .stMarkdown p, .stMarkdown li, .stMarkdown strong {
        color: #004D3D !important;
        font-weight: 700 !important;
    }

    div[data-testid="stRadio"] label p {
        color: #004D3D !important;
        font-weight: 800 !important;
        font-size: 0.9rem !important;
    }

    div[data-testid="stForm"], 
    div[data-testid="stVerticalBlock"] > div[style*="border"] {
        background-color: #ffffff !important;
        border-radius: 14px !important;
    }

    /* Abas em Cápsula (Sem Linha Vermelha) */
    div[data-baseweb="tab-list"] {
        gap: 8px !important;
        background-color: #e2ede7 !important;
        padding: 6px !important;
        border-radius: 14px !important;
        border: 1px solid #d1e2d8 !important;
    }

    button[data-baseweb="tab"] {
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        color: #004D3D !important;
        padding: 10px 20px !important;
        border-radius: 10px !important;
        border: none !important;
        background-color: transparent !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color: #ffffff !important;
        background-color: #004D3D !important;
        box-shadow: 0 4px 12px rgba(0, 77, 61, 0.25) !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] p {
        color: #ffffff !important;
    }

    div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] {
        display: none !important;
    }

    /* Botão Principal */
    .stButton > button {
        background: linear-gradient(135deg, #76BC21 0%, #5d9818 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 800 !important;
        font-size: 0.92rem !important;
        padding: 12px 24px !important;
        box-shadow: 0 5px 14px rgba(118, 188, 33, 0.35) !important;
    }
    .stButton > button p {
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

# 4. Funções do Banco de Dados Excel
def cargar_hoja(sheet_name):
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def guardar_registro_excel(sheet_name, nuevo_registro):
    if os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_existente = cargar_hoja(sheet_name)
            df_nuevo = pd.DataFrame([nuevo_registro])
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
            df_final.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        df_nuevo = pd.DataFrame([nuevo_registro])
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df_nuevo.to_excel(writer, sheet_name=sheet_name, index=False)

# 5. Painel Esquerdo (Sidebar)
with st.sidebar:
    logo_candidates = [
        "logo.png",
        "logo.png.png",
        "Isologo Vitafrutta PNG Color.png",
        "Isologo Vitafrutta PNG Color.png.png",
        "isologo_vitafruta_png_color.png",
        "isologo_vitafrutta_png_color.png",
        "logo.jpg",
        "logo.jpeg"
    ]
    
    logo_path = None
    for name in logo_candidates:
        full_path = os.path.join(BASE_DIR, name)
        if os.path.exists(full_path):
            logo_path = full_path
            break
            
    if not logo_path and os.path.exists(BASE_DIR):
        for file_name in os.listdir(BASE_DIR):
            if file_name.lower().endswith(('.png', '.jpg', '.jpeg')) and 'logo' in file_name.lower():
                logo_path = os.path.join(BASE_DIR, file_name)
                break
            
    st.markdown('<div class="sidebar-logo-card">', unsafe_allow_html=True)
    if logo_path:
        st.image(logo_path, use_container_width=True)
    else:
        st.markdown("<h2 style='color:#004D3D; font-weight:800;'>🍋 VITAFRUTTA</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#004D3D; font-size:0.75rem; font-weight:800; margin-top:2px;'>SYSTEM SUITE HACCP</p>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    modulos = [
        "1. Higiene y Sanitización de Planta",
        "2. Higiene y Sanitización del Personal",
        "3. Control de Plagas (MIP)",
        "4. Calidad y Control de Agua Potable",
        "5. Control de Proveedores y M.P.",
        "6. Mantenimiento y Calibración",
        "7. Control de Cadena de Frío",
        "8. Gestión de Acciones Correctivas",
        "9. Sistema de Trazabilidad",
        "10. Retiro de Productos (Recall)",
        "11. Manejo de Basura y Residuos"
    ]
    
    opcion_modulo = st.selectbox("Módulo Activo:", modulos)
    
    st.markdown("""
        <div class="plant-status-card">
            <div class="plant-status-title">📍 Planta Melipilla, Chile</div>
            <div class="plant-status-text">Sistema Unificado de Control de Calidad</div>
            <div class="plant-status-text"><b>Estado:</b> Servidor Local Activo</div>
        </div>
    """, unsafe_allow_html=True)

# 6. Banner Principal da Área de Trabalho
st.markdown(f"""
    <div class="main-banner">
        <div class="main-banner-title">🛡️ Sistema Integrado HACCP</div>
        <div class="main-banner-subtitle">{opcion_modulo}</div>
        <div class="live-badge">
            <div class="pulse-green"></div>
            MONITOREO EN TIEMPO REAL &bull; BASE EXCEL MASTER CONECTADA
        </div>
    </div>
""", unsafe_allow_html=True)

# 7. Roteamento de Módulos (1 a 11 Totalmente Funcionais)

# MÓDULO 1: HIGIENE Y SANITIZACIÓN DE PLANTA
if "1. Higiene y Sanitización de Planta" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Estado Sistema CIP", "OPERATIVO", "Líneas 1 & 2")
    with k2: st.metric("Concentración Target", "200 PPM", "Cloro / Peracético")
    with k3: st.metric("Temp. Norma CIP", "65 °C", "Mín. 60 °C Exigido")
    with k4: st.metric("Frecuencia Monitoreo", "Por Turno", "Frecuencia Diaria")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Registrar Inspección", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Formulario de Higiene de Planta</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha", datetime.today(), key="m1_fecha")
                    turno = st.selectbox("Turno", ["Mañana", "Tarde", "Noche"], key="m1_turno")
                    equipo = st.selectbox("Área/Maquinaria", ["Tolva y Lavadora", "Extractoras", "Línea CIP", "Envasadora", "IQF y Cámaras"], key="m1_eq")
                with col2:
                    ppm = st.number_input("Concentración (PPM)", 0, 1000, 200, key="m1_ppm")
                    temp = st.number_input("Temp CIP (°C)", 10, 90, 65, key="m1_temp")
                    resp = st.selectbox("Inspector", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas", "Noemi Alvarado"], key="m1_resp")
                estado = st.radio("Conformidad", ["CONFORME", "NO CONFORME"], horizontal=True, key="m1_est")
                
                if st.button("💾 Guardar Registro POES 01", use_container_width=True, key="m1_btn"):
                    reg = {"Fecha": str(fecha), "Turno": turno, "Equipo/Área": equipo, "PPM": ppm, "Temp_CIP": temp, "Inspector": resp, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES01", reg)
                    st.success("✅ Registro guardado en POES 01.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Criterios Normativos</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* **Extractoras:** 130 - 170 PPM\n* **Línea CIP:** 180 - 220 PPM\n* **Pisos/Pediluvios:** 450 - 550 PPM")

    with t2:
        df = cargar_hoja("Registros_POES01")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 01.")

# MÓDULO 2: HIGIENE Y SANITIZACIÓN DEL PERSONAL
elif "2. Higiene y Sanitización del Personal" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Cumplimiento Personal", "98.5%", "Meta > 95%")
    with k2: st.metric("Control Lavado Manos", "100%", "Verificación Diaria")
    with k3: st.metric("Estado Salud Operarios", "APTO", "Sin Sintomatología")
    with k4: st.metric("Auditoría Vestuarios", "CONFORME", "Turno Actual")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Registrar Inspección", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Control de Higiene del Personal</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha", datetime.today(), key="m2_fecha")
                    operario = st.text_input("Nombre / RUT del Operario", key="m2_op")
                    area = st.selectbox("Área de Trabajo", ["Procesamiento / Extracción", "Envasado y Empaque", "Cámaras y Bodega", "Mantenimiento"], key="m2_area")
                with col2:
                    uniforme = st.selectbox("Estado Uniforme y Cofia", ["Limpio y Completo", "Deficiente / Incompleto"], key="m2_uni")
                    manos_unas = st.selectbox("Lavado de Manos y Uñas", ["Correcto / Sin Joyas", "No Conforme"], key="m2_manos")
                    inspector = st.selectbox("Supervisor Responsable", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas", "Noemi Alvarado"], key="m2_resp")
                
                estado = st.radio("Evaluación General", ["APTO PARA OPERAR", "RECHAZADO / CORREGIR"], horizontal=True, key="m2_est")
                
                if st.button("💾 Guardar Registro POES 02", use_container_width=True, key="m2_btn"):
                    reg = {"Fecha": str(fecha), "Operario": operario, "Área": area, "Uniforme": uniforme, "Manos_Uñas": manos_unas, "Inspector": inspector, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES02", reg)
                    st.success("✅ Registro guardado en POES 02.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Normas de Higiene Personal</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Uso obligatorio de cofia y mascarilla.\n* Prohibido uso de relojes, anillos e alhajas.\n* Uñas cortas, limpos e sem esmalte.")

    with t2:
        df = cargar_hoja("Registros_POES02")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 02.")

# MÓDULO 3: CONTROL DE PLAGAS (MIP)
elif "3. Control de Plagas" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Estaciones Monitoreadas", "24 / 24", "100% Cobertura")
    with k2: st.metric("Captura / Hallazgos", "0 Unidades", "Límite Crítico: 0")
    with k3: st.metric("Consumo Cebo Rodenticida", "0%", "Sin Actividad")
    with k4: st.metric("Lamp. UV Atrapa-Insectos", "OPERATIVAS", "Inspección Semanal")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Registrar Control MIP", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Registro de Monitoreo de Plagas</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha de Inspección", datetime.today(), key="m3_fecha")
                    estacion = st.text_input("N° Estación / Trampa UV", "Estación 01 - Périmetro Ext.", key="m3_estacion")
                    tipo = st.selectbox("Tipo de Trampa", ["Caja Cebadera Exterior", "Lámpara UV Atrapa-Insectos", "Trampa de Impacto / Pegajosa"], key="m3_tipo")
                with col2:
                    hallazgo = st.selectbox("Evidencia / Hallazgo", ["Ninguna (Sin Actividad)", "Insecto Volador Capturado", "Consumo Parcial Cebo", "Consumo Total Cebo"], key="m3_hall")
                    tecnico = st.text_input("Técnico / Empresa Controladora", "Rentokil / Ecolab", key="m3_tec")
                
                estado = st.radio("Condición de la Trampa", ["EN BUEN ESTADO", "REQUIERE REPOSICIÓN / MANTENCIÓN"], horizontal=True, key="m3_est")
                
                if st.button("💾 Guardar Registro POES 03", use_container_width=True, key="m3_btn"):
                    reg = {"Fecha": str(fecha), "Estación": estacion, "Tipo": tipo, "Hallazgo": hallazgo, "Técnico": tecnico, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES03", reg)
                    st.success("✅ Registro guardado en POES 03.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Criterios Manejo Integrado</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Prohibido uso de cebos tóxicos dentro de la sala de proceso.\n* Revisión de perímetro cada 15 días.\n* Limpieza semanal de bandejas UV.")

    with t2:
        df = cargar_hoja("Registros_POES03")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 03.")

# MÓDULO 4: CALIDAD Y CONTROL DE AGUA POTABLE
elif "4. Calidad y Control de Agua Potable" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Cloro Libre Residual", "1.2 PPM", "Norma: 0.2 - 2.0 PPM")
    with k2: st.metric("pH Agua Potable", "7.1", "Rango: 6.5 - 8.5")
    with k3: st.metric("Turbidez", "< 1 NTU", "Agua Cristalina")
    with k4: st.metric("Análisis Microbiológico", "CONFORME", "Sin Coliformes")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Registrar Control de Agua", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Análisis Diario de Agua Potable</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha", datetime.today(), key="m4_fecha")
                    punto = st.selectbox("Punto de Muestra", ["Grifo Sala de Extracción", "Línea Entrada CIP", "Grifo Lavadero de Fruta", "Vestuarios / Servicios"], key="m4_punto")
                    cloro = st.number_input("Cloro Libre Residual (PPM)", 0.0, 5.0, 1.2, step=0.1, key="m4_cloro")
                with col2:
                    ph = st.number_input("Valor pH", 0.0, 14.0, 7.1, step=0.1, key="m4_ph")
                    inspector = st.selectbox("Analista / Inspector", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas", "Noemi Alvarado"], key="m4_resp")
                
                estado = st.radio("Evaluación Normativa", ["AGUA POTABLE CONFORME", "NO CONFORME (INYECTAR CLORO)"], horizontal=True, key="m4_est")
                
                if st.button("💾 Guardar Registro POES 04", use_container_width=True, key="m4_btn"):
                    reg = {"Fecha": str(fecha), "Punto": punto, "Cloro_PPM": cloro, "pH": ph, "Inspector": inspector, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES04", reg)
                    st.success("✅ Registro guardado en POES 04.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Límites Críticos NCh409</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* **Cloro Residual Mínimo:** 0.2 PPM\n* **Cloro Residual Máximo:** 2.0 PPM\n* Si Cloro < 0.2 PPM, detener proceso y recargar dosificador.")

    with t2:
        df = cargar_hoja("Registros_POES04")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 04.")

# MÓDULO 5: CONTROL DE PROVEEDORES Y M.P.
elif "5. Control de Proveedores y M.P." in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Fruta Recepcionada", "12,500 Kg", "Lote Hoy")
    with k2: st.metric("°Brix Promedio MP", "9.5 °Brix", "Norma Limón Cítrico")
    with k3: st.metric("Descarte de Fruta", "1.8%", "Tolva y Seleção")
    with k4: st.metric("Conformidad Guías", "100%", "Facturas y Guías Ok")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Recepción de Materia Prima", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Formulario de Recepción de Fruta e Insumos</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha Recepción", datetime.today(), key="m5_fecha")
                    proveedor = st.text_input("Nombre Proveedor / Agrícola", "Agrícola Melipilla SpA", key="m5_prov")
                    lote_mp = st.text_input("Lote Materia Prima / Guía", "G-88452", key="m5_lote")
                with col2:
                    kilos = st.number_input("Kilos Totales Recepcionados", 0, 100000, 12500, step=500, key="m5_kg")
                    brix = st.number_input("Grados °Brix Muestra", 0.0, 20.0, 9.5, step=0.1, key="m5_brix")
                    inspector = st.selectbox("Receptor Responsable", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas", "Noemi Alvarado"], key="m5_resp")
                
                estado = st.radio("Aceptación de Carga", ["LOTE ACEPTADO", "LOTE RECHAZADO / CON OBSERVACIÓN"], horizontal=True, key="m5_est")
                
                if st.button("💾 Guardar Registro POES 05", use_container_width=True, key="m5_btn"):
                    reg = {"Fecha": str(fecha), "Proveedor": proveedor, "Lote_Guia": lote_mp, "Kilos": kilos, "Brix": brix, "Inspector": inspector, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES05", reg)
                    st.success("✅ Registro guardado en POES 05.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Estándar de Calidad MP</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Fruta fresca sem podridão ou mofo.\n* Mínimo 8.5 °Brix para processamento.\n* Guía de despacho com rastreabilidade de origem.")

    with t2:
        df = cargar_hoja("Registros_POES05")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 05.")

# MÓDULO 6: MANTENIMIENTO Y CALIBRACIÓN
elif "6. Mantenimiento y Calibración" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Refratômetros Opticos", "CALIBRADOS", "Calibración con Agua Distilada")
    with k2: st.metric("Termómetros Digitales", "0.0 °C Error", "Baño de Hielo Patronal")
    with k3: st.metric("Mantenimiento Preventivo", "92%", "Cumplimiento Plan")
    with k4: st.metric("Bombas y Pasteurizador", "OPERATIVOS", "HTST en Servicio")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Registro de Calibración / Mantención", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Registro de Mantención y Calibración</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha", datetime.today(), key="m6_fecha")
                    equipo = st.selectbox("Equipo / Instrumento", ["Refratómetro Atago #1", "Refratómetro Atago #2", "Termómetro Infrarrojo IQF", "Sensor Temp Pasteurizador", "Bomba de Pasteurización"], key="m6_eq")
                    tipo_act = st.selectbox("Tipo de Actividad", ["Calibración de Rutina", "Mantenimiento Preventivo", "Mantenimiento Correctivo / Reparación"], key="m6_tipo")
                with col2:
                    lectura = st.text_input("Lectura Patrón / Observación", "0.0 °Brix con Agua Destilada", key="m6_lec")
                    tecnico = st.selectbox("Técnico / Responsable", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas", "Servicio Técnico Externo"], key="m6_tec")
                
                estado = st.radio("Resultado", ["CONFORME / CALIBRADO", "FORA DE ESPECIFICAÇÃO / MANUTENÇÃO"], horizontal=True, key="m6_est")
                
                if st.button("💾 Guardar Registro POES 06", use_container_width=True, key="m6_btn"):
                    reg = {"Fecha": str(fecha), "Equipo": equipo, "Tipo": tipo_act, "Lectura": lectura, "Técnico": tecnico, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES06", reg)
                    st.success("✅ Registro guardado en POES 06.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Plan de Calibraciones</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Refratômetros: Calibração diária a 0 °Brix.\n* Termômetros: Verificação semanal em água com gelo (0 °C).\n* Manômetros: Verificação semestral.")

    with t2:
        df = cargar_hoja("Registros_POES06")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 06.")

# MÓDULO 7: CONTROL DE CADENA DE FRÍO
elif "7. Control de Cadena de Frío" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Túnel IQF Temperatura", "-35 °C", "Set Point: -35 °C")
    with k2: st.metric("Cámara Conservación #1", "-18.5 °C", "Norma < -18 °C")
    with k3: st.metric("Cámara Conservación #2", "-19.1 °C", "Norma < -18 °C")
    with k4: st.metric("Descongelamiento / Defrost", "INACTIVO", "Ciclo Automático Ok")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Control de Temperatura Cámaras", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Monitoreo de Cadena de Frío y Cámaras</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha", datetime.today(), key="m7_fecha")
                    camara = st.selectbox("Unidad de Frío", ["Túnel IQF Rodajas", "Cámara de Productos Congelados #1", "Cámara de Productos Congelados #2", "Ante-cámara de Despacho"], key="m7_cam")
                    temp = st.number_input("Temperatura Medida (°C)", -50.0, 10.0, -18.5, step=0.5, key="m7_temp")
                with col2:
                    puertas = st.selectbox("Estado de Cortinas y Puertas", ["Sellado Correcto", "Fuga de Aire / Dejar Abierta"], key="m7_puer")
                    inspector = st.selectbox("Inspector de Frío", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas", "Noemi Alvarado"], key="m7_resp")
                
                estado = st.radio("Evaluación de Frío", ["DENTRO DE RANGO (-18°C a -35°C)", "DESVIACIÓN DE TEMPERATURA"], horizontal=True, key="m7_est")
                
                if st.button("💾 Guardar Registro POES 07", use_container_width=True, key="m7_btn"):
                    reg = {"Fecha": str(fecha), "Cámara": camara, "Temp_C": temp, "Puertas": puertas, "Inspector": inspector, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES07", reg)
                    st.success("✅ Registro guardado en POES 07.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Límites PCC Frío</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* **Cámaras de Mantención:** Mínimo -18 °C.\n* **Túnel IQF:** Operação entre -30 °C e -38 °C.\n* Se Temp > -15 °C por mais de 2 horas: Ação Corretiva Imediata.")

    with t2:
        df = cargar_hoja("Registros_POES07")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 07.")

# MÓDULO 8: GESTIÓN DE ACCIONES CORRECTIVAS
elif "8. Gestión de Acciones Correctivas" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Acciones Abiertas", "1 Caso", "En Proceso de Solución")
    with k2: st.metric("Acciones Cerradas", "12 Casos", "Resolución Conforme")
    with k3: st.metric("Eficacia de Solución", "100%", "Verificación Cierre")
    with k4: st.metric("Tiempo Promedio Cierre", "24 Horas", "Cumplimiento Meta")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Registrar Acción Correctiva", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Formulario de Acción Correctiva y Preventiva</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha Desvío", datetime.today(), key="m8_fecha")
                    mod_origen = st.selectbox("Módulo donde ocurrió el Desvío", ["Módulo 1: Higiene Planta", "Módulo 4: Agua Potable", "Módulo 5: Materia Prima", "Módulo 7: Cadena de Frío"], key="m8_mod")
                    desvio = st.text_area("Descripción de la No Conformidad", "Concentración de cloro en lavado de fruta bajo el límite mínimo (100 PPM).", key="m8_desv")
                with col2:
                    accion = st.text_area("Acción Correctiva Inmediata Tomada", "Se detuvo la línea, se ajustó la bomba dosificadora y se re-sanitizó el lote.", key="m8_acc")
                    responsable = st.selectbox("Responsable Cierre", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas"], key="m8_resp")
                
                estado = st.radio("Estado de la Acción", ["ACCIÓN CERRADA / EFECTIVA", "EN PROCESO / PENDIENTE DE REVISIÓN"], horizontal=True, key="m8_est")
                
                if st.button("💾 Guardar Registro POES 08", use_container_width=True, key="m8_btn"):
                    reg = {"Fecha": str(fecha), "Módulo_Origen": mod_origen, "Desvío": desvio, "Acción_Tomada": accion, "Responsable": responsable, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES08", reg)
                    st.success("✅ Registro guardado en POES 08.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Protocolo de Cierre</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Investigar Causa Raíz.\n* Bloqueo inmediato del producto afectado.\n* Re-inspección y firma de conformidad por Aseguramiento de Calidad.")

    with t2:
        df = cargar_hoja("Registros_POES08")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 08.")

# MÓDULO 9: SISTEMA DE TRAZABILIDAD
elif "9. Sistema de Trazabilidad" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Lotes Trazados Hoy", "3 Lotes", "Producción IQF y Jugo")
    with k2: st.metric("Precision Rastreabilidade", "100%", "Materia Prima a PT")
    with k3: st.metric("Tiempo de Búsqueda", "< 5 Minutos", "Estándar Auditoría")
    with k4: st.metric("Mermas de Proceso", "2.1%", "Rendimiento Normal")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Registrar Vínculo de Trazabilidad", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Asociación de Lotes de Producción</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha Producción", datetime.today(), key="m9_fecha")
                    lote_pt = st.text_input("Lote Producto Terminado (PT)", "PT-IQF-20261026-A", key="m9_pt")
                    lote_mp = st.text_input("Lote Materia Prima (MP)", "G-88452 (Agrícola Melipilla)", key="m9_mp")
                with col2:
                    kilos_pt = st.number_input("Kilos Producidos (Kg)", 0, 50000, 3500, step=100, key="m9_kgpt")
                    cliente = st.text_input("Cliente / Destino de Exportación", "Japón / Foodex 2027 Container #1", key="m9_cli")
                    responsable = st.selectbox("Responsable Trazabilidad", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas"], key="m9_resp")
                
                estado = st.radio("Estado de Trazabilidad", ["LOTE COMPLETAMENTE TRAZADO", "PENDIENTE DE DESPACHO"], horizontal=True, key="m9_est")
                
                if st.button("💾 Guardar Registro POES 09", use_container_width=True, key="m9_btn"):
                    reg = {"Fecha": str(fecha), "Lote_PT": lote_pt, "Lote_MP": lote_mp, "Kilos_PT": kilos_pt, "Cliente_Destino": cliente, "Responsable": responsable, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES09", reg)
                    st.success("✅ Registro guardado en POES 09.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Regla Un Paso Atrás / Adelante</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Identificar Proveedor e Guía de Origen.\n* Registrar insumos e envases utilizados.\n* Identificar Cliente e Guía de Despacho.")

    with t2:
        df = cargar_hoja("Registros_POES09")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 09.")

# MÓDULO 10: RETIRO DE PRODUCTOS (RECALL)
elif "10. Retiro de Productos (Recall)" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Simulacros Anuales", "2 / 2 Realizados", "100% Cumplido")
    with k2: st.metric("Eficacia Recuperación", "99.8%", "Simulacro Último")
    with k3: st.metric("Tiempo de Bloqueo", "15 Minutos", "Aviso a Distribuição")
    with k4: st.metric("Comité de Crisis", "LISTO", "Contactos Actualizados")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Simulado / Evento de Recall", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Registro de Simulacro / Evento Recall</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha Evento / Simulado", datetime.today(), key="m10_fecha")
                    cod_recall = st.text_input("Código de Ejercicio / Evento", "SIMULACRO-RECALL-2026-02", key="m10_cod")
                    lote_afectado = st.text_input("Lote de Producto Involucrado", "PT-IQF-20260915-B", key="m10_lote")
                with col2:
                    kg_recuperados = st.number_input("Kilos Localizados / Recuperados (Kg)", 0, 50000, 4800, step=100, key="m10_kg")
                    lider = st.selectbox("Líder Comité de Crisis", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz"], key="m10_resp")
                
                estado = st.radio("Evaluación del Simulacro", ["SIMULACRO EXITOSO (> 98%)", "REQUIERE AJUSTE EN TIEMPOS"], horizontal=True, key="m10_est")
                
                if st.button("💾 Guardar Registro POES 10", use_container_width=True, key="m10_btn"):
                    reg = {"Fecha": str(fecha), "Código": cod_recall, "Lote_Involucrado": lote_afectado, "Kg_Recuperados": kg_recuperados, "Líder": lider, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES10", reg)
                    st.success("✅ Registro guardado en POES 10.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Etapas de Recall</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Convocar Comité de Crisis.\n* Detener despacho e notificar clientes.\n* Recuperar e isolar produto na câmara de retenção.")

    with t2:
        df = cargar_hoja("Registros_POES10")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 10.")

# MÓDULO 11: MANEJO DE BASURA Y RESIDUOS
elif "11. Manejo de Basura y Residuos" in opcion_modulo:
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Bagaço de Limão (Orgânico)", "8.2 Toneladas", "Retirada Agrícola")
    with k2: st.metric("Lixeiras de Processo", "100% Sanitizadas", "Limpeza Diária")
    with k3: st.metric("Plásticos e Embalagens", "RECICLADO", "Destino Certificado")
    with k4: st.metric("Área de Acúmulo Ext.", "CONFORME", "Sem Odores / Vetores")

    st.markdown("<br>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 Control de Evacuación de Residuos", "📊 Histórico Registrado (Excel)"])

    with t1:
        c_form, c_info = st.columns([2, 1])
        with c_form:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Registro de Evacuación de Residuos</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                col1, col2 = st.columns(2)
                with col1:
                    fecha = st.date_input("Fecha", datetime.today(), key="m11_fecha")
                    tipo_res = st.selectbox("Tipo de Residuo", ["Bagaço e Casca de Limão (Orgânico)", "Plásticos / Cartões (Reciclável)", "Residuos Domésticos / Lixo Comum", "Lodo / Sedimentos Tratamento"], key="m11_tipo")
                    kilos = st.number_input("Cantidad Estimada (Kg / Ton)", 0, 20000, 2500, step=100, key="m11_kg")
                with col2:
                    empresa = st.text_input("Destino / Empresa Retiro", "Agrícola Compost / Retiro Municipal", key="m11_emp")
                    inspector = st.selectbox("Inspector Responsable", ["Dante Refosco", "Carolina Alvarado", "Cristina Ortiz", "Ana Rosa Rojas", "Noemi Alvarado"], key="m11_resp")
                
                estado = st.radio("Estado da Área de Resíduos", ["ÁREA LIMPA E SANITIZADA", "REQUIERE RETIRO URGENTE"], horizontal=True, key="m11_est")
                
                if st.button("💾 Guardar Registro POES 11", use_container_width=True, key="m11_btn"):
                    reg = {"Fecha": str(fecha), "Tipo_Residuo": tipo_res, "Cantidad": kilos, "Destino": empresa, "Inspector": inspector, "Estado": estado, "Hora": datetime.now().strftime("%H:%M:%S")}
                    guardar_registro_excel("Registros_POES11", reg)
                    st.success("✅ Registro guardado en POES 11.")
                    st.rerun()

        with c_info:
            st.markdown("<h4 style='color:#004D3D; font-weight:800;'>Regras de Resíduos</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("* Lixeiras com tampa acionada por pedal.\n* Evacuação total de resíduos ao final de cada turno.\n* Proibido acúmulo na área limpa de processo.")

    with t2:
        df = cargar_hoja("Registros_POES11")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
        else: st.info("No hay registros en POES 11.")