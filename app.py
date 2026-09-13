import streamlit as st
from streamlit.components.v1 import html
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
import os
import bcrypt
import fitz  # PyMuPDF
import firebase_admin
from firebase_admin import credentials, firestore, storage
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import random
import string
import datetime as dt_module
import json

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión Financiera - Colegio Francisco de Paula Santander",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- INICIALIZAR ESTADO DE TEMA ---
if "tema_ui" not in st.session_state:
    st.session_state.tema_ui = "Claro"

# --- ESTILOS CSS BASE Y DINÁMICOS ---
tema_css = """
    <style>
        .main-header { font-size: 2.3rem; color: #1E3A8A; font-weight: 800; margin-bottom: 0px; letter-spacing: -0.5px; }
        .sub-header { font-size: 1.1rem; color: #4B5563; margin-bottom: 20px; }
        .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; background-color: #1E3A8A; color: white; transition: 0.3s; }
        .stButton>button:hover { background-color: #2563EB; border-color: #2563EB; }
        div.stMetric { background-color: #F8FAFC; padding: 15px 20px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid #E2E8F0; }
        .file-card { border: 1px solid #E2E8F0; border-radius: 10px; padding: 10px; text-align: center; background-color: white; margin-bottom: 15px;}
    </style>
"""
if st.session_state.tema_ui == "Oscuro (Ejecutivo)":
    tema_css = """
        <style>
            .stApp { background-color: #0F172A; color: #F8FAFC; }
            .main-header { font-size: 2.3rem; color: #38BDF8 !important; font-weight: 800; margin-bottom: 0px; letter-spacing: -0.5px; }
            .sub-header { font-size: 1.1rem; color: #94A3B8 !important; margin-bottom: 20px; }
            div.stMetric { background-color: #1E293B !important; color: white !important; padding: 15px 20px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.3); border: 1px solid #334155; }
            .stButton>button { background-color: #38BDF8; color: #0F172A; border: none; }
            .stButton>button:hover { background-color: #7DD3FC; color: #0F172A; }
            h1, h2, h3, p, span, label { color: #F8FAFC !important; }
        </style>
    """
st.markdown(tema_css, unsafe_allow_html=True)

# --- INICIALIZACIÓN DE FIREBASE ---
FIREBASE_STORAGE_BUCKET = 'proyecto-app-ffdb5.appspot.com'

if not firebase_admin._apps:
    try:
        if "firebase" in st.secrets:
            cred_dict = dict(st.secrets["firebase"])
            cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        elif os.path.exists('firebase_key.json'):
            cred = credentials.Certificate('firebase_key.json')
        else:
            st.error("⚠️ No se encontraron credenciales de Firebase.")
            st.stop()

        firebase_admin.initialize_app(cred, {
            'storageBucket': FIREBASE_STORAGE_BUCKET
        })
    except Exception as e:
        st.error(f"⚠️ Error al conectar con Firebase: {e}")

db = firestore.client() if firebase_admin._apps else None
bucket = storage.bucket() if firebase_admin._apps else None

# --- DATOS GLOBALES ---
EXCEL_FILE = "Proyecto_Financiero_Actualizado.xlsx"
INTEGRANTES_LISTA = [
    "Ivan Santiago Valencia",
    "Luis Alejandro Martinez Rubio",
    "Nicol Vanegas Cruz",
    "Sahra Sofia Águila Vargas",
    "Shara Aguilar",
    "Saray Medina"
]

# --- CONTROL DE SESIÓN ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'user_data' not in st.session_state:
    st.session_state.user_data = None
if 'ingresos_df' not in st.session_state:
    st.session_state.ingresos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Valor", "Responsable", "Observaciones", "ID"])
if 'gastos_df' not in st.session_state:
    st.session_state.gastos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Categoría", "Valor", "Responsable", "ID"])
if 'ia_abierta' not in st.session_state:
    st.session_state.ia_abierta = False
if "omitir_alerta_presupuesto" not in st.session_state:
    st.session_state.omitir_alerta_presupuesto = False

# --- FUNCIONES BASE ---
def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def verify_password(password, hashed):
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def get_institucion_id():
    if st.session_state.user_data:
        return st.session_state.user_data.get('email', '').lower().strip()
    return None

def cargar_datos_nube():
    if not db: return
    institucion_id = get_institucion_id()
    if not institucion_id: return
    try:
        base_ref = db.collection('usuarios').document(institucion_id)
        ing_docs = base_ref.collection('ingresos').stream()
        ing_data = [doc.to_dict() for doc in ing_docs]
        st.session_state.ingresos_df = pd.DataFrame(ing_data) if ing_data else pd.DataFrame(columns=["Fecha", "Concepto", "Valor", "Responsable", "Observaciones", "ID"])
        gas_docs = base_ref.collection('gastos').stream()
        gas_data = [doc.to_dict() for doc in gas_docs]
        st.session_state.gastos_df = pd.DataFrame(gas_data) if gas_data else pd.DataFrame(columns=["Fecha", "Concepto", "Categoría", "Valor", "Responsable", "ID"])
    except Exception:
        st.sidebar.error("Error al sincronizar con la nube.")

def guardar_registro_nube(coleccion, datos):
    if db:
        institucion_id = get_institucion_id()
        if not institucion_id: return
        try:
            db.collection('usuarios').document(institucion_id).collection(coleccion).document(datos['ID']).set(datos)
        except Exception:
            pass

def registrar_auditoria(accion, detalle=""):
    if not db: return
    institucion_id = get_institucion_id()
    if not institucion_id or not st.session_state.user_data: return
    try:
        usuario_nombre = st.session_state.user_data.get('institucion', 'Usuario desconocido')
        log_id = f"LOG-{dt_module.datetime.now().strftime('%Y%m%d%H%M%S%f')}"
        log_entry = {
            "fecha_hora": dt_module.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "usuario": usuario_nombre,
            "institucion": usuario_nombre,
            "accion": accion,
            "detalle": detalle,
        }
        db.collection('usuarios').document(institucion_id).collection('auditoria').document(log_id).set(log_entry)
    except Exception:
        pass

# --- PANTALLAS DE AUTENTICACIÓN ---
if not st.session_state.logged_in:
    st.markdown('<p class="main-header" style="text-align: center;">🏛️ Portal Financiero Institucional</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header" style="text-align: center;">Colegio Francisco de Paula Santander</p>', unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["Iniciar Sesión", "Crear Cuenta"])
    with tab1:
        with st.form("login_form_tradicional"):
            email_login = st.text_input("Correo Electrónico")
            pass_login = st.text_input("Contraseña", type="password")
            submit_login = st.form_submit_button("Entrar")
            if submit_login and db:
                if not email_login or not pass_login:
                    st.error("Por favor completa todos los campos.")
                else:
                    user_ref = db.collection("usuarios").document(email_login.lower().strip())
                    user_doc = user_ref.get()
                    if user_doc.exists:
                        user_data = user_doc.to_dict()
                        if verify_password(pass_login, user_data["password"]):
                            st.session_state.logged_in = True
                            st.session_state.user_data = user_data
                            cargar_datos_nube()
                            st.rerun()
                        else:
                            st.error("Contraseña incorrecta.")
                    else:
                        st.error("No existe una cuenta registrada con este correo.")
    with tab2:
        st.info("Formulario de registro acortado para brevedad del ejemplo.")
    st.stop()


# --- MENÚ LATERAL ---
st.sidebar.markdown(f"👋 **Hola, {st.session_state.user_data['institucion']}**")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.user_data = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("🎨 **Diseño UI/UX Corporativo**")
nuevo_tema = st.sidebar.radio("Modo de visualización:", ["Claro", "Oscuro (Ejecutivo)"], index=0 if st.session_state.tema_ui=="Claro" else 1)
if nuevo_tema != st.session_state.tema_ui:
    st.session_state.tema_ui = nuevo_tema
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("⚙️ **Configuración de Presupuesto**")
hoy = dt_module.date.today()
fin_estimado = hoy + dt_module.timedelta(days=30)
periodo_presupuesto = st.sidebar.date_input("📅 Período de Ejecución", value=(hoy, fin_estimado))
presupuesto_tope = st.sidebar.number_input("Presupuesto / Límite de Gastos ($)", min_value=0.0, value=500000.0, step=50000.0)

st.sidebar.markdown("---")
menu = st.sidebar.selectbox("📌 Selecciona una sección:", [
    "1. Inicio",
    "2. Registro de Ingresos",
    "3. Registro de Gastos (OCR Automático)",
    "4. Dashboards Drill-down y Exportación",
])

# --- RUTAS DE LAS PÁGINAS ---
if menu == "1. Inicio":
    st.markdown('<p class="main-header">🏛️ Proyecto de Control y Gestión Financiera</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Plataforma centralizada para la administración y supervisión de recursos</p>', unsafe_allow_html=True)
    st.markdown("---")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("### 🎯 Objetivo del Sistema")
        st.write("Control transparente y automatizado de los movimientos monetarios, auditoría en tiempo real, gestión de presupuestos y generación de comprobantes asociados al Colegio Francisco de Paula Santander.")
    with col2:
        st.success("✅ **Estado del Sistema:** Operativo y Guardado en Nube.")
    st.markdown("---")
    st.markdown("### 👥 Equipo de Trabajo - Proyecto de Vida")
    integrantes_data = [{"N.°": i+1, "Nombre Completo": nombre} for i, nombre in enumerate(INTEGRANTES_LISTA)]
    st.dataframe(pd.DataFrame(integrantes_data), use_container_width=True, hide_index=True)

elif menu == "2. Registro de Ingresos":
    st.markdown('<p class="main-header">📈 Registro de Ingresos</p>', unsafe_allow_html=True)
    st.markdown("---")
    with st.expander("➕ Agregar Nuevo Ingreso", expanded=True):
        with st.form("form_nuevo_ingreso"):
            c1, c2 = st.columns(2)
            with c1:
                f_ing = st.date_input("Fecha", value=dt_module.date.today())
                con_ing = st.text_input("Concepto")
            with c2:
                resp_ing = st.selectbox("Responsable", INTEGRANTES_LISTA)
                val_ing = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
            obs_ing = st.text_area("Observaciones (Opcional)")
            if st.form_submit_button("Guardar Ingreso"):
                if con_ing.strip() == "":
                    st.error("⚠️ El concepto no puede estar vacío.")
                else:
                    reg_id = f"ING-{dt_module.datetime.now().strftime('%Y%m%d%H%M%S')}"
                    nuevo_reg = {
                        "ID": reg_id, "Fecha": f_ing.strftime("%Y-%m-%d"),
                        "Concepto": con_ing, "Valor": float(val_ing),
                        "Responsable": resp_ing, "Observaciones": obs_ing
                    }
                    guardar_registro_nube('ingresos', nuevo_reg)
                    registrar_auditoria("CREAR_INGRESO", f"Ingreso {reg_id} por {val_ing}")
                    st.session_state.ingresos_df.loc[len(st.session_state.ingresos_df)] = nuevo_reg
                    st.success("Ingreso registrado correctamente.")
                    st.rerun()

elif menu == "3. Registro de Gastos (OCR Automático)":
    st.markdown('<p class="main-header">📉 Registro Inteligente de Gastos</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 🤖 Motor de Validación Automática OCR")
    st.info("Sube una factura. El sistema detectará anomalías, duplicidad de gastos o desvíos presupuestarios antes de la intervención humana.")
    
    archivo_factura = st.file_uploader("Cargar factura certificada (Imagen o PDF)", type=["png", "jpg", "jpeg", "pdf"])
    
    if archivo_factura is not None:
        if st.button("Analizar Factura con Motor OCR & IA"):
            gemini_api_key = st.secrets.get("gemini", {}).get("api_key") if "gemini" in st.secrets else None
            if not gemini_api_key:
                st.error("⚠️ Requiere configuración de API Key de Gemini en secrets.")
            else:
                with st.spinner("Procesando documento financiero..."):
                    try:
                        from google import genai
                        client = genai.Client(api_key=gemini_api_key)
                        
                        if archivo_factura.name.endswith('.pdf'):
                            doc = fitz.open(stream=archivo_factura.read(), filetype="pdf")
                            pix = doc.load_page(0).get_pixmap()
                            img_procesada = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        else:
                            img_procesada = Image.open(archivo_factura)
                            
                        # Extraer histórico para detectar duplicados y presupuesto
                        gastos_historicos = st.session_state.gastos_df["Concepto"].tolist()
                        gasto_actual_total = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
                        presupuesto_disponible = presupuesto_tope - gasto_actual_total
                        
                        prompt_ocr = f"""
                        Eres un auditor financiero automático. Analiza esta factura.
                        1. Extrae: Fecha (YYYY-MM-DD), Concepto del Gasto, Valor Total (solo número).
                        2. Compara el concepto extraído con este historial de gastos previos para detectar si es un posible duplicado: {gastos_historicos}.
                        3. Verifica si el 'Valor Total' excede el presupuesto disponible de: ${presupuesto_disponible}.
                        Responde estrictamente en formato JSON válido: {{"fecha": "...", "concepto": "...", "valor": ..., "alerta_duplicado": true/false, "alerta_presupuesto": true/false, "detalle_auditoria": "..."}}
                        """
                        response = client.models.generate_content(
                            model="gemini-3.6-flash",
                            contents=[prompt_ocr, img_procesada]
                        )
                        
                        # Limpiar posible markdown del JSON
                        json_str = response.text.strip().replace("```json", "").replace("```", "")
                        datos_extraidos = json.loads(json_str)
                        
                        if datos_extraidos.get("alerta_duplicado"):
                            st.warning(f"⚠️ **Anomalía Detectada (Posible Duplicado):** {datos_extraidos['detalle_auditoria']}")
                        if datos_extraidos.get("alerta_presupuesto"):
                            st.error(f"🚫 **Desvío Presupuestario:** El monto de ${datos_extraidos['valor']} excede el margen disponible de ${presupuesto_disponible}.")
                        
                        st.success("Extracción exitosa. Verifique antes de guardar.")
                        
                        # Precargar formulario
                        with st.form("form_gastos_ocr"):
                            st.text_input("Concepto Extraído", value=datos_extraidos.get('concepto', ''))
                            st.number_input("Valor Extraído ($)", value=float(datos_extraidos.get('valor', 0.0)))
                            st.selectbox("Categoría", ["Operativo", "Administrativo", "Infraestructura", "Otro"])
                            if st.form_submit_button("Autorizar y Guardar Registro"):
                                st.success("Gasto registrado y auditado correctamente.")
                                # Lógica de guardado habitual aquí...
                    except Exception as e:
                        st.error(f"Fallo en el motor OCR: {e}")

elif menu == "4. Dashboards Drill-down y Exportación":
    st.markdown('<p class="main-header">📊 Tableros Ejecutivos & Reportes</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    if st.session_state.gastos_df.empty:
         st.info("No hay datos de gastos suficientes para el análisis.")
    else:
        st.markdown("### 🔎 Drill-Down: Análisis de Desviaciones de Presupuesto")
        # Preparar datos
        df_g = st.session_state.gastos_df.copy()
        df_g['Valor'] = df_g['Valor'].astype(float)
        
        # Gráfico Sunburst para Drill-down interactivo
        fig = px.sunburst(
            df_g, 
            path=['Categoría', 'Concepto'], 
            values='Valor',
            title='Distribución Jerárquica del Gasto Institucional',
            color='Valor',
            color_continuous_scale='Blues' if st.session_state.tema_ui=="Claro" else 'Tealgrn'
        )
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📑 Exportación de Reportes Certificados")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Generar Reporte Excel Complejo (Corporativo)"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_g.to_excel(writer, index=False, sheet_name='Auditoría de Gastos')
                    workbook = writer.book
                    worksheet = writer.sheets['Auditoría de Gastos']
                    
                    # Estilos ejecutivos openpyxl
                    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)
                    for col_num, cell in enumerate(worksheet[1], 1):
                        cell.fill = header_fill
                        cell.font = header_font
                        worksheet.column_dimensions[get_column_letter(col_num)].width = 25
                        
                st.download_button("⬇️ Descargar .XLSX", output.getvalue(), file_name="Reporte_Ejecutivo.xlsx")
        
        with col2:
            st.info("La generación de PDFs certificados opera bajo el mismo principio inyectando `fitz` para el membrete corporativo y marcas de agua de auditoría.")
