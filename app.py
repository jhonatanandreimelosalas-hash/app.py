import streamlit as st
from streamlit.components.v1 import html
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
from datetime import datetime
import plotly.express as px
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

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión Financiera - Colegio Francisco de Paula Santander",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS ---
st.markdown("""
    <style>
        .main-header { font-size: 2.3rem; color: #1E3A8A; font-weight: 800; margin-bottom: 0px; letter-spacing: -0.5px; }
        .sub-header { font-size: 1.1rem; color: #4B5563; margin-bottom: 20px; }
        .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; background-color: #1E3A8A; color: white; transition: 0.3s; }
        .stButton>button:hover { background-color: #2563EB; border-color: #2563EB; }
        div.stMetric { background-color: #F8FAFC; padding: 15px 20px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid #E2E8F0; }
        .file-card { border: 1px solid #E2E8F0; border-radius: 10px; padding: 10px; text-align: center; background-color: white; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

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
    "Jhonnattan Andrei Melo Salas"
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

# --- FUNCIONES DE AUTENTICACIÓN ---
def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def verify_password(password, hashed):
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

# --- FUNCIONES DE SINCRONIZACIÓN CON FIREBASE ---
def cargar_datos_nube():
    if not db: return
    try:
        ing_docs = db.collection('ingresos').stream()
        ing_data = [doc.to_dict() for doc in ing_docs]
        if ing_data:
            st.session_state.ingresos_df = pd.DataFrame(ing_data)
        
        gas_docs = db.collection('gastos').stream()
        gas_data = [doc.to_dict() for doc in gas_docs]
        if gas_data:
            st.session_state.gastos_df = pd.DataFrame(gas_data)
    except Exception as e:
        st.sidebar.error("Error al sincronizar con la nube.")

def guardar_registro_nube(coleccion, datos):
    if db:
        try:
            db.collection(coleccion).document(datos['ID']).set(datos)
        except Exception:
            pass

def eliminar_registro_nube(coleccion, doc_id):
    if db:
        try:
            db.collection(coleccion).document(doc_id).delete()
        except Exception:
            pass

# --- FUNCIÓN GENERAR MINIATURA PDF ---
def generar_miniatura_pdf(file_bytes):
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=70)
        return buf.getvalue()
    except Exception as e:
        return None

# --- FUNCIÓN PARA GENERAR IMAGEN DE RECIBO DECORADA ---
def generar_imagen_recibo(rec_id, fecha, tot_ing, tot_gas, saldo, qr_img_pil):
    img_w, img_h = 650, 880
    base_img = Image.new("RGB", (img_w, img_h), color="#FFFFFF")
    draw = ImageDraw.Draw(base_img)
    
    try:
        font_title = ImageFont.truetype("arial.ttf", 22)
        font_bold = ImageFont.truetype("arialbd.ttf", 15)
        font_regular = ImageFont.truetype("arial.ttf", 14)
        font_small = ImageFont.truetype("arial.ttf", 11)
    except IOError:
        font_title = font_bold = font_regular = font_small = ImageFont.load_default()

    draw.rectangle([(0, 0), (img_w, 110)], fill="#1E3A8A")
    draw.text((30, 25), "COLEGIO FRANCISCO DE PAULA SANTANDER", fill="#FFFFFF", font=font_title)
    draw.text((30, 60), "Comprobante General de Balance Financiero", fill="#93C5FD", font=font_regular)
    
    draw.rectangle([(30, 130), (img_w - 30, img_h - 40)], outline="#E2E8F0", width=2, fill="#F8FAFC")
    
    draw.text((55, 160), "ID de Comprobante:", fill="#64748B", font=font_small)
    draw.text((200, 158), f"{rec_id}", fill="#1E293B", font=font_bold)
    
    draw.text((55, 190), "Fecha de Emisión:", fill="#64748B", font=font_small)
    draw.text((200, 188), f"{fecha}", fill="#1E293B", font=font_bold)

    draw.text((55, 220), "Institución:", fill="#64748B", font=font_small)
    draw.text((200, 218), "Colegio Francisco de Paula Santander", fill="#1E293B", font=font_bold)
    
    draw.line([(55, 255), (img_w - 55, 255)], fill="#CBD5E1", width=1)
    
    draw.text((55, 280), "RESUMEN DE MOVIMIENTOS", fill="#1E3A8A", font=font_bold)
    
    draw.text((55, 320), "(+) Total Ingresos:", fill="#334155", font=font_regular)
    draw.text((400, 320), f"${tot_ing:,.0f} COP", fill="#059669", font=font_bold)
    
    draw.text((55, 360), "(-) Total Gastos:", fill="#334155", font=font_regular)
    draw.text((400, 360), f"${tot_gas:,.0f} COP", fill="#DC2626", font=font_bold)
    
    draw.line([(55, 400), (img_w - 55, 400)], fill="#CBD5E1", width=1)
    
    draw.text((55, 420), "BALANCE NETO FINAL:", fill="#1E3A8A", font=font_bold)
    color_saldo = "#059669" if saldo >= 0 else "#DC2626"
    draw.text((370, 415), f"${saldo:,.0f} COP", fill=color_saldo, font=font_title)
    
    estado_txt = "ESTADO: APROBADO (SUPERÁVIT)" if saldo >= 0 else "ESTADO: ALERTA (DÉFICIT)"
    draw.text((55, 465), estado_txt, fill=color_saldo, font=font_small)

    qr_resized = qr_img_pil.resize((180, 180))
    base_img.paste(qr_resized, (int((img_w - 180) / 2), 510))
    
    draw.text((int(img_w / 2) - 130, 710), "Escanea este código QR para validar", fill="#64748B", font=font_small)
    draw.text((int(img_w / 2) - 120, 730), "la información general del balance", fill="#64748B", font=font_small)
    
    draw.text((int(img_w / 2) - 110, 800), "Sistema Automático de Gestión Financiera", fill="#94A3B8", font=font_small)

    buffer_img = BytesIO()
    base_img.save(buffer_img, format="PNG")
    buffer_img.seek(0)
    return buffer_img

# --- PANTALLAS DE AUTENTICACIÓN ---
if not st.session_state.logged_in:
    st.markdown('<p class="main-header" style="text-align: center;">🏛️ Portal Financiero Institucional</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header" style="text-align: center;">Colegio Francisco de Paula Santander</p>', unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["Iniciar Sesión", "Crear Cuenta", "Olvidé mi Contraseña"])
    
    with tab1:
        st.markdown("### Acceso con Credenciales")
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
                            st.success("¡Bienvenido/a!")
                            st.rerun()
                        else:
                            st.error("Contraseña incorrecta.")
                    else:
                        st.error("No existe una cuenta registrada con este correo.")

    with tab2:
        with st.form("register_form"):
            inst_name = st.text_input("Nombre de la Institución / Persona")
            email_reg = st.text_input("Correo Electrónico")
            pass_reg = st.text_input("Contraseña (Min. 6 caracteres, 1 mayúscula)", type="password")
            submit_reg = st.form_submit_button("Registrar Cuenta")
            
            if submit_reg and db:
                if len(pass_reg) < 6 or not any(c.isupper() for c in pass_reg):
                    st.error("La contraseña debe tener al menos 6 caracteres y 1 letra mayúscula.")
                elif not inst_name or not email_reg:
                    st.error("Todos los campos son obligatorios.")
                else:
                    email_clean = email_reg.lower().strip()
                    email_exists = db.collection('usuarios').document(email_clean).get().exists
                    
                    if email_exists:
                        st.error("Ya existe una cuenta con este correo electrónico.")
                    else:
                        nuevo_usuario = {
                            'institucion': inst_name,
                            'email': email_clean,
                            'password': hash_password(pass_reg),
                            'fecha_creacion': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        }
                        db.collection('usuarios').document(email_clean).set(nuevo_usuario)
                        st.success("¡Cuenta creada exitosamente! Ya puedes iniciar sesión.")

    with tab3:
        with st.form("forgot_form"):
            st.info("Ingresa tu correo y te enviaremos una contraseña temporal de recuperación a tu bandeja.")
            email_forgot = st.text_input("Correo Electrónico registrado")
            submit_forgot = st.form_submit_button("Enviar Contraseña Temporal")
            
            if submit_forgot and db:
                email_clean = email_forgot.lower().strip()
                user_ref = db.collection('usuarios').document(email_clean)
                user_doc = user_ref.get()
                if user_doc.exists:
                    temp_pass = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
                    user_ref.update({'password': hash_password(temp_pass)})
                    
                    try:
                        remitente = st.secrets["smtp"]["email"]
                        password_smtp = st.secrets["smtp"]["password"]
                        
                        msg = MIMEMultipart()
                        msg['From'] = remitente
                        msg['To'] = email_clean
                        msg['Subject'] = "Recuperación de Contraseña - Colegio Francisco de Paula Santander"
                        
                        cuerpo = f"Hola,\n\nHas solicitado recuperar tu contraseña en el Portal Financiero.\nTu nueva contraseña temporal es: {temp_pass}\n\nInicia sesión con ella y recuerda cambiarla."
                        msg.attach(MIMEText(cuerpo, 'plain'))
                        
                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.starttls()
                        server.login(remitente, password_smtp)
                        server.sendmail(remitente, email_clean, msg.as_string())
                        server.quit()
                        
                        st.success(f"✅ ¡Correo enviado exitosamente a {email_clean}! Revisa tu bandeja de entrada o spam.")
                    except Exception as e:
                        st.error(f"Error al enviar el correo. Asegúrate de configurar los secretos [smtp] en Streamlit. Detalle: {e}")
                else:
                    st.error("El correo no está registrado en nuestra base de datos.")
    
    st.stop()

# --- MENÚ LATERAL ---
st.sidebar.markdown(f"👋 **Hola, {st.session_state.user_data['institucion']}**")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.user_data = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("⚙️ **Configuración de Presupuesto**")
presupuesto_tope = st.sidebar.number_input("Presupuesto / Límite de Gastos ($)", min_value=0.0, value=500000.0, step=50000.0)

st.sidebar.markdown("---")
menu = st.sidebar.selectbox("📌 Selecciona una sección:", [
    "1. Inicio", 
    "2. Registro de Ingresos", 
    "3. Registro de Gastos", 
    "4. Balance Financiero", 
    "5. Dashboard y Gráficos", 
    "6. Anexo de Recibos & QR", 
    "7. Gestión de Archivos",
    "8. Reporte Final",
    "9. Indicadores y Dashboard",
])

# --- APARTADO DE IA EN EL BORDE ---
st.sidebar.markdown("---")
st.sidebar.markdown("🤖 **Asistente IA del Borde**")
if st.sidebar.button("💬 Abrir / Cerrar Asistente IA"):
    st.session_state.ia_abierta = not st.session_state.ia_abierta

if st.session_state.ia_abierta:
    with st.sidebar.container():
        st.markdown("### 🧠 Chat Asesor IA")
        api_key_input = st.text_input("Clave de API Gemini:", type="password", key="api_key_ia")
        pregunta_ia = st.text_input("¿Qué deseas consultar?")
        
        if st.button("Consultar IA"):
            if not api_key_input:
                st.error("⚠️ Introduce tu clave de API.")
            else:
                try:
                    from google import genai
                    client = genai.Client(api_key=api_key_input)
                    
                    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
                    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
                    saldo = tot_ing - tot_gas
                    
                    contexto = f"Datos del proyecto: Ingresos=${tot_ing}, Gastos=${tot_gas}, Saldo=${saldo}."
                    prompt_completo = f"{contexto}\nPregunta: {pregunta_ia}"
                    
                    response = client.models.generate_content(
                        model="gemini-3.6-flash",
                        contents=prompt_completo,
                    )
                    st.success("Respuesta:")
                    st.write(response.text)
                except Exception as e:
                    st.error(f"Error con la IA: {e}")

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
                f_ing = st.date_input("Fecha", value=datetime.now())
                con_ing = st.text_input("Concepto")
            with c2:
                resp_ing = st.selectbox("Responsable", INTEGRANTES_LISTA)
                val_ing = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
            obs_ing = st.text_area("Observaciones (Opcional)")
            
            if st.form_submit_button("Guardar Ingreso"):
                if con_ing.strip() == "":
                    st.error("⚠️ El concepto no puede estar vacío.")
                else:
                    reg_id = f"ING-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                    nuevo_reg = {
                        "ID": reg_id,
                        "Fecha": f_ing.strftime("%Y-%m-%d"),
                        "Concepto": con_ing,
                        "Valor": float(val_ing),
                        "Responsable": resp_ing,
                        "Observaciones": obs_ing
                    }
                    st.session_state.ingresos_df = pd.concat([st.session_state.ingresos_df, pd.DataFrame([nuevo_reg])], ignore_index=True)
                    guardar_registro_nube('ingresos', nuevo_reg)
                    st.success("¡Ingreso guardado en la nube!")
                    st.rerun()

    if not st.session_state.ingresos_df.empty:
        st.dataframe(st.session_state.ingresos_df.drop(columns=['ID']), use_container_width=True)
        st.metric("💵 TOTAL INGRESOS", f"${st.session_state.ingresos_df['Valor'].astype(float).sum():,.0f} COP")
        
        st.markdown("### 🗑️ Eliminar Ingreso")
        opciones = [f"{row['Concepto']} - ${row['Valor']:,.0f}" for _, row in st.session_state.ingresos_df.iterrows()]
        seleccion = st.selectbox("Selecciona para eliminar:", opciones)
        
        if st.button("❌ Eliminar Ingreso"):
            idx = opciones.index(seleccion)
            doc_id = st.session_state.ingresos_df.iloc[idx]['ID']
            eliminar_registro_nube('ingresos', doc_id)
            st.session_state.ingresos_df = st.session_state.ingresos_df.drop(idx).reset_index(drop=True)
            st.success("Ingreso eliminado.")
            st.rerun()

elif menu == "3. Registro de Gastos":
    st.markdown('<p class="main-header">📉 Registro de Gastos</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    current_total_gastos = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    if presupuesto_tope > 0 and current_total_gastos > presupuesto_tope:
        st.error(f"🚨 ¡ATENCIÓN! Has superado el límite de ${presupuesto_tope:,.0f} COP.")
    elif presupuesto_tope > 0:
        st.info(f"ℹ️ Disponible: ${(presupuesto_tope - current_total_gastos):,.0f} COP.")

    with st.expander("➕ Agregar Nuevo Gasto", expanded=True):
        with st.form("form_nuevo_gasto"):
            c1, c2 = st.columns(2)
            with c1:
                f_gas = st.date_input("Fecha Gasto", value=datetime.now())
                con_gas = st.text_input("Concepto")
                cat_gas = st.selectbox("Categoría", ["Logística", "Publicidad", "Alimentación", "Varios"])
            with c2:
                val_gas = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
                resp_gas = st.selectbox("Responsable", INTEGRANTES_LISTA)
            
            if st.form_submit_button("Guardar Gasto"):
                if con_gas.strip() == "":
                    st.error("⚠️ El concepto no puede estar vacío.")
                else:
                    reg_id = f"GAS-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                    nuevo_reg = {
                        "ID": reg_id,
                        "Fecha": f_gas.strftime("%Y-%m-%d"),
                        "Concepto": con_gas,
                        "Categoría": cat_gas,
                        "Valor": float(val_gas),
                        "Responsable": resp_gas
                    }
                    st.session_state.gastos_df = pd.concat([st.session_state.gastos_df, pd.DataFrame([nuevo_reg])], ignore_index=True)
                    guardar_registro_nube('gastos', nuevo_reg)
                    st.success("¡Gasto guardado en la nube!")
                    st.rerun()

    if not st.session_state.gastos_df.empty:
        st.dataframe(st.session_state.gastos_df.drop(columns=['ID']), use_container_width=True)
        st.metric("💸 TOTAL GASTOS", f"${current_total_gastos:,.0f} COP")

        st.markdown("### 🗑️ Eliminar Gasto")
        opciones = [f"{row['Concepto']} - ${row['Valor']:,.0f}" for _, row in st.session_state.gastos_df.iterrows()]
        seleccion = st.selectbox("Selecciona para eliminar:", opciones)
        
        if st.button("❌ Eliminar Gasto"):
            idx = opciones.index(seleccion)
            doc_id = st.session_state.gastos_df.iloc[idx]['ID']
            eliminar_registro_nube('gastos', doc_id)
            st.session_state.gastos_df = st.session_state.gastos_df.drop(idx).reset_index(drop=True)
            st.success("Gasto eliminado.")
            st.rerun()

elif menu == "4. Balance Financiero":
    st.markdown('<p class="main-header">⚖️ Balance Financiero General</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    saldo = tot_ing - tot_gas
    
    c1, c2, c3 = st.columns(3)
    c1.metric("💵 Ingresos", f"${tot_ing:,.0f} COP")
    c2.metric("💸 Gastos", f"${tot_gas:,.0f} COP")
    c3.metric("💰 Saldo Neto", f"${saldo:,.0f} COP", delta=f"${saldo:,.0f} COP")

elif menu == "5. Dashboard y Gráficos":
    st.markdown('<p class="main-header">📊 Dashboard Interactivo</p>', unsafe_allow_html=True)
    st.markdown("---")
    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Ingresos vs Gastos")
        df_comp = pd.DataFrame({"Tipo": ["Ingresos", "Gastos"], "Monto": [tot_ing, tot_gas]})
        fig_bar = px.bar(df_comp, x="Tipo", y="Monto", color="Tipo", text_auto=True, color_discrete_sequence=["#10B981", "#EF4444"])
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("#### Gastos por Categoría")
        if not st.session_state.gastos_df.empty:
            df_cat = st.session_state.gastos_df.groupby("Categoría")["Valor"].sum().reset_index()
            fig_pie = px.pie(df_cat, names="Categoría", values="Valor", hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3)
            st.plotly_chart(fig_pie, use_container_width=True)

elif menu == "6. Anexo de Recibos & QR":
    st.markdown('<p class="main-header">🧾 Generador de Comprobante General</p>', unsafe_allow_html=True)
    st.markdown("---")
    st.info("💡 Haz clic para generar el comprobante oficial decorado con los datos financieros actuales y su código QR.")

    if st.button("🚀 Generar Comprobante Oficial"):
        tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
        tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
        saldo = tot_ing - tot_gas
        rec_id = f"GEN-{datetime.now().strftime('%Y%m%d%H%M')}"
        fecha_actual = datetime.now().strftime('%Y-%m-%d')
        
        texto_recibo = f"COMPROBANTE {rec_id}\nInstitucion: Colegio Francisco de Paula Santander\nIngresos: ${tot_ing:,.0f}\nGastos: ${tot_gas:,.0f}\nSaldo: ${saldo:,.0f}"
        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(texto_recibo)
        qr.make(fit=True)
        qr_img_pil = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        
        # Llamar a la función robusta de generación de imagen decorada
        buffer_recibo = generar_imagen_recibo(rec_id, fecha_actual, tot_ing, tot_gas, saldo, qr_img_pil)
        st.session_state.rec_img_bytes = buffer_recibo.getvalue()
        st.success("✅ ¡Comprobante generado exitosamente!")
        
    if 'rec_img_bytes' in st.session_state:
        st.image(st.session_state.rec_img_bytes, width=450)
        st.download_button(
            "📥 Descargar Comprobante PNG", 
            data=st.session_state.rec_img_bytes, 
            file_name="Comprobante_Financiero.png", 
            mime="image/png"
        )

elif menu == "7. Gestión de Archivos":
    st.markdown('<p class="main-header">📁 Repositorio de Documentos</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Registra, administra y visualiza los comprobantes del proyecto.</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    import base64

    with st.form("form_subir_archivo"):
        archivo_subido = st.file_uploader("Sube tu archivo (PDF, PNG, JPG)", type=["png", "jpg", "jpeg", "pdf"])
        descripcion_archivo = st.text_input("Descripción o Nota del Documento")
        submit_archivo = st.form_submit_button("💾 Guardar y Registrar Archivo")
        
        if submit_archivo and db:
            if archivo_subido is not None:
                bytes_archivo = archivo_subido.getvalue()
                base64_archivo = base64.b64encode(bytes_archivo).decode('utf-8')
                
                nombre_id = f"ARCH-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                doc_data = {
                    "ID": nombre_id,
                    "nombre": archivo_subido.name,
                    "tipo": archivo_subido.type,
                    "archivo_b64": base64_archivo,
                    "descripcion": descripcion_archivo if descripcion_archivo else "Sin descripción",
                    "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "subido_por": st.session_state.user_data['institucion']
                }
                db.collection("archivos").document(nombre_id).set(doc_data)
                st.success("✅ ¡Archivo guardado exitosamente en la base de datos!")
                st.rerun()
            else:
                st.error("⚠️ Por favor selecciona un archivo antes de guardar.")

    st.markdown("### 📋 Archivos Registrados con Vista Previa")
    if db:
        try:
            archivos_ref = db.collection("archivos").stream()
            archivos_lista = [a.to_dict() for a in archivos_ref]
            
            if archivos_lista:
                for row in archivos_lista:
                    with st.expander(f"📄 {row['nombre']} — ({row['fecha']})"):
                        st.write(f"**Descripción:** {row['descripcion']}")
                        st.write(f"**Subido por:** {row['subido_por']}")
                        
                        if "archivo_b64" in row:
                            b64_bytes = base64.b64decode(row['archivo_b64'])
                            
                            # Mostrar mini recuadro de vista previa si es imagen
                            if row['tipo'].startswith('image/'):
                                st.image(b64_bytes, caption="Vista previa", width=250)
                            elif row['tipo'] == 'application/pdf':
                                st.info("📎 Archivo PDF (Usa el botón de abajo para descargarlo y abrirlo)")
                            
                            st.download_button(
                                label=f"📥 Descargar / Abrir {row['nombre']}",
                                data=b64_bytes,
                                file_name=row['nombre'],
                                mime=row['tipo'],
                                key=f"dl_{row['ID']}"
                            )
                
                st.markdown("---")
                st.markdown("### 🗑️ Eliminar Registro de Archivo")
                opciones_arch = [f"{row['nombre']} ({row['fecha']})" for row in archivos_lista]
                sel_arch = st.selectbox("Selecciona archivo a eliminar:", opciones_arch)
                
                if st.button("❌ Eliminar Registro"):
                    idx = opciones_arch.index(sel_arch)
                    doc_id_eliminar = archivos_lista[idx]['ID']
                    db.collection("archivos").document(doc_id_eliminar).delete()
                    st.success("Registro eliminado correctamente.")
                    st.rerun()
            else:
                st.info("Aún no hay archivos registrados en el sistema.")
        except Exception as e:
            st.error(f"Error al cargar los archivos: {e}")
    else:
        st.warning("Conecta Firebase para ver la lista de archivos.")
elif menu == "8. Reporte Final":
    st.markdown('<p class="main-header">📑 Reporte Financiero Profesional</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Genera y descarga un libro de Excel con diseño institucional avanzado.</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    if st.button("📊 Generar Excel Profesional"):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        HEADER_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        TITLE_FONT = Font(name="Arial", size=16, bold=True, color="1E3A8A")
        REGULAR_FONT = Font(name="Arial", size=10, color="333333")
        
        THIN_BORDER = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        # Obtener el nombre del usuario logueado dinámicamente
        nombre_usuario = st.session_state.user_data.get('institucion', 'Institución Financiera')

        def estilizar_hoja(ws, titulo_base, df):
            titulo_completo = f"{nombre_usuario.upper()} - {titulo_base}"
            ws.append([titulo_completo])
            ws.append([]) 
            ws.cell(row=1, column=1).font = TITLE_FONT
            
            if df.empty:
                ws.append(["No hay registros disponibles."])
                return

            headers = [col for col in df.columns if col != 'ID']
            ws.append(headers)
            
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=3, column=col_num)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = THIN_BORDER

            for row_idx, row in df.iterrows():
                fila_datos = [row[col] for col in headers]
                ws.append(fila_datos)
                
                current_row = ws.max_row
                is_zebra = (row_idx % 2 != 0)
                
                for col_num in range(1, len(headers) + 1):
                    cell = ws.cell(row=current_row, column=col_num)
                    cell.font = REGULAR_FONT
                    cell.border = THIN_BORDER
                    if is_zebra:
                        cell.fill = zebra_fill
                        
                    if headers[col_num - 1] == "Valor" and isinstance(cell.value, (int, float)):
                        cell.number_format = '"$"#,##0'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row > 2 and cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

        # 1. Hoja de Ingresos
        ws_ing = wb.create_sheet(title="Ingresos")
        estilizar_hoja(ws_ing, "REPORTE DE INGRESOS", st.session_state.ingresos_df)

        # 2. Hoja de Gastos
        ws_gas = wb.create_sheet(title="Gastos")
        estilizar_hoja(ws_gas, "REPORTE DE GASTOS", st.session_state.gastos_df)

        wb.save(EXCEL_FILE)
        st.success("✅ ¡Reporte Excel generado con éxito usando el nombre de tu cuenta!")

    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                "📥 Descargar Archivo Excel Profesional", 
                data=f, 
                file_name="Reporte_Financiero.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

elif menu == "9. Indicadores y Dashboard":
    st.markdown('<p class="main-header">📈 Indicadores Clave (KPIs)</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Análisis visual en tiempo real del comportamiento financiero.</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    saldo = tot_ing - tot_gas
    
    # Tarjetas Métricas Modernas (KPI Cards)
    col1, col2, col3 = st.columns(3)
    col1.metric("💵 Total Ingresos", f"${tot_ing:,.0f} COP")
    col2.metric("💸 Total Gastos", f"${tot_gas:,.0f} COP")
    col3.metric("💰 Balance Neto", f"${saldo:,.0f} COP", delta=f"${saldo:,.0f} COP")
    
    st.markdown("---")
    c_graf1, c_graf2 = st.columns(2)
    
    with c_graf1:
        st.markdown("#### 🥧 Gastos por Categoría")
        if not st.session_state.gastos_df.empty and "Categoría" in st.session_state.gastos_df.columns:
            df_cat = st.session_state.gastos_df.groupby("Categoría")["Valor"].sum().reset_index()
            fig_pie = px.pie(df_cat, names="Categoría", values="Valor", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No hay suficientes datos de gastos para mostrar la categoría.")

    with c_graf2:
        st.markdown("#### 📈 Evolución Temporal de Movimientos")
        df_all = []
        if not st.session_state.ingresos_df.empty:
            df_i = st.session_state.ingresos_df[["Fecha", "Valor"]].copy()
            df_i["Tipo"] = "Ingreso"
            df_all.append(df_i)
        if not st.session_state.gastos_df.empty:
            df_g = st.session_state.gastos_df[["Fecha", "Valor"]].copy()
            df_g["Tipo"] = "Gasto"
            df_all.append(df_g)
            
        if df_all:
            df_timeline = pd.concat(df_all, ignore_index=True)
            df_timeline["Fecha"] = pd.to_datetime(df_timeline["Fecha"])
            df_timeline = df_timeline.sort_values("Fecha")
            fig_line = px.line(df_timeline, x="Fecha", y="Valor", color="Tipo", markers=True, color_discrete_map={"Ingreso": "#10B981", "Gasto": "#EF4444"})
            st.plotly_chart(fig_line, use_container_width=True)
        else:
            st.info("Registra ingresos o gastos para ver la línea de tiempo.")
