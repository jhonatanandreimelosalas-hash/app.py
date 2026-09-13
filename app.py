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
import datetime as dt_module  # renombrado para no chocar con 'from datetime import datetime'

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión Financiera - Colegio Francisco de Paula Santander",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS (tema claro/oscuro dinámico) ---
if 'modo_oscuro' not in st.session_state:
    st.session_state.modo_oscuro = False


def render_estilos(modo_oscuro: bool) -> str:
    """Devuelve el bloque <style> según el modo activo. Paleta sobria institucional
    en ambos casos; solo cambian fondo, superficies y contraste de texto."""
    if modo_oscuro:
        fondo_app = "#0F172A"
        superficie = "#1E293B"
        borde = "#334155"
        texto_principal = "#E2E8F0"
        texto_secundario = "#94A3B8"
        acento = "#3B82F6"
        acento_hover = "#60A5FA"
    else:
        fondo_app = "#FFFFFF"
        superficie = "#F8FAFC"
        borde = "#E2E8F0"
        texto_principal = "#1E293B"
        texto_secundario = "#4B5563"
        acento = "#1E3A8A"
        acento_hover = "#2563EB"

    return f"""
    <style>
        .stApp {{ background-color: {fondo_app}; color: {texto_principal}; }}
        .main-header {{ font-size: 2.3rem; color: {acento}; font-weight: 800; margin-bottom: 0px; letter-spacing: -0.5px; }}
        .sub-header {{ font-size: 1.1rem; color: {texto_secundario}; margin-bottom: 20px; }}
        .stButton>button {{ width: 100%; border-radius: 8px; font-weight: 600; background-color: {acento}; color: white; transition: 0.3s; border: none; }}
        .stButton>button:hover {{ background-color: {acento_hover}; border-color: {acento_hover}; }}
        div.stMetric {{ background-color: {superficie}; padding: 15px 20px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.15); border: 1px solid {borde}; }}
        .file-card {{ border: 1px solid {borde}; border-radius: 10px; padding: 10px; text-align: center; background-color: {superficie}; margin-bottom: 15px; }}
        section[data-testid="stSidebar"] {{ background-color: {superficie}; }}
        .stDataFrame, .stTable {{ background-color: {superficie}; }}
        p, span, label, .stMarkdown {{ color: {texto_principal}; }}
    </style>
    """


st.markdown(render_estilos(st.session_state.modo_oscuro), unsafe_allow_html=True)

st.sidebar.toggle("🌙 Modo oscuro", key="modo_oscuro")

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


# =====================================================================================
# REQUISITO 1: AISLAMIENTO DE DATOS POR INSTITUCIÓN
# -------------------------------------------------------------------------------------
# Se usa el correo limpio (lower + strip) como identificador de institución, porque
# es el mismo valor que ya se usa como ID del documento en la colección 'usuarios'
# (ver 'user_ref = db.collection("usuarios").document(email_login.lower().strip())'
# más abajo, sin modificar). El nombre visible ('institucion') puede repetirse entre
# cuentas distintas; el email de cuenta no.
#
# Estructura elegida: subcolecciones anidadas bajo el documento de cada institución:
#   usuarios/{institucion_id}/ingresos/{doc_id}
#   usuarios/{institucion_id}/gastos/{doc_id}
#   usuarios/{institucion_id}/archivos/{doc_id}
#   usuarios/{institucion_id}/auditoria/{doc_id}
#
# Se prefirió esto sobre "colecciones globales + campo institucion + where()" porque
# un query sin el filtro (por error humano futuro) no puede devolver datos ajenos:
# la separación vive en la ruta del documento, no en una condición que alguien podría
# omitir accidentalmente al añadir una función nueva más adelante.
# =====================================================================================

def get_institucion_id():
    """ID estable de la institución actual: el correo limpio del usuario logueado."""
    if st.session_state.user_data:
        return st.session_state.user_data.get('email', '').lower().strip()
    return None


def cargar_datos_nube():
    if not db:
        return
    institucion_id = get_institucion_id()
    if not institucion_id:
        return
    try:
        base_ref = db.collection('usuarios').document(institucion_id)

        ing_docs = base_ref.collection('ingresos').stream()
        ing_data = [doc.to_dict() for doc in ing_docs]
        if ing_data:
            st.session_state.ingresos_df = pd.DataFrame(ing_data)
        else:
            st.session_state.ingresos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Valor", "Responsable", "Observaciones", "ID"])

        gas_docs = base_ref.collection('gastos').stream()
        gas_data = [doc.to_dict() for doc in gas_docs]
        if gas_data:
            st.session_state.gastos_df = pd.DataFrame(gas_data)
        else:
            st.session_state.gastos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Categoría", "Valor", "Responsable", "ID"])
    except Exception:
        st.sidebar.error("Error al sincronizar con la nube.")


def guardar_registro_nube(coleccion, datos):
    if db:
        institucion_id = get_institucion_id()
        if not institucion_id:
            return
        try:
            db.collection('usuarios').document(institucion_id).collection(coleccion).document(datos['ID']).set(datos)
        except Exception:
            pass


def eliminar_registro_nube(coleccion, doc_id):
    if db:
        institucion_id = get_institucion_id()
        if not institucion_id:
            return
        try:
            db.collection('usuarios').document(institucion_id).collection(coleccion).document(doc_id).delete()
        except Exception:
            pass


CATEGORIAS_GASTO = ["Logística", "Publicidad", "Alimentación", "Varios"]


def cargar_presupuestos_categoria():
    """Lee el presupuesto asignado por categoría desde el documento de la institución.
    Si no existe todavía, devuelve un diccionario en ceros."""
    if not db:
        return {cat: 0.0 for cat in CATEGORIAS_GASTO}
    institucion_id = get_institucion_id()
    if not institucion_id:
        return {cat: 0.0 for cat in CATEGORIAS_GASTO}
    try:
        doc = db.collection('usuarios').document(institucion_id).get()
        datos = doc.to_dict() or {}
        guardado = datos.get('presupuestos_categoria', {})
        return {cat: float(guardado.get(cat, 0.0)) for cat in CATEGORIAS_GASTO}
    except Exception:
        return {cat: 0.0 for cat in CATEGORIAS_GASTO}


def guardar_presupuestos_categoria(presupuestos: dict):
    if not db:
        return
    institucion_id = get_institucion_id()
    if not institucion_id:
        return
    try:
        db.collection('usuarios').document(institucion_id).set(
            {'presupuestos_categoria': presupuestos}, merge=True
        )
    except Exception:
        pass


def calcular_tabla_desviaciones(gastos_df, presupuestos_categoria):
    """Calcula gasto real vs. presupuesto asignado por categoría. Se usa tanto en el
    Tablero Ejecutivo (sección 5) como en el Reporte PDF Corporativo (sección 8), para
    que ambos muestren siempre el mismo número."""
    df_gastos_cat = gastos_df.copy() if gastos_df is not None and not gastos_df.empty else pd.DataFrame(columns=["Categoría", "Valor"])
    if "Valor" not in df_gastos_cat.columns:
        df_gastos_cat["Valor"] = 0.0
    df_gastos_cat["Valor"] = pd.to_numeric(df_gastos_cat["Valor"], errors='coerce').fillna(0)
    gasto_real_por_cat = df_gastos_cat.groupby("Categoría")["Valor"].sum() if "Categoría" in df_gastos_cat.columns and not df_gastos_cat.empty else pd.Series(dtype=float)

    df_desviacion = pd.DataFrame({
        "Categoría": CATEGORIAS_GASTO,
        "Gasto Real": [gasto_real_por_cat.get(c, 0.0) for c in CATEGORIAS_GASTO],
        "Presupuesto Asignado": [presupuestos_categoria.get(c, 0.0) for c in CATEGORIAS_GASTO],
    })
    df_desviacion["Desviación"] = df_desviacion["Gasto Real"] - df_desviacion["Presupuesto Asignado"]
    df_desviacion["Estado"] = df_desviacion["Desviación"].apply(lambda d: "Sobre presupuesto" if d > 0 else "Dentro de presupuesto")
    return df_desviacion


# =====================================================================================
# OCR DE FACTURAS Y DETECCIÓN DE ANOMALÍAS (reglas, no IA, para duplicados/desviaciones)
# -------------------------------------------------------------------------------------
# La extracción de datos de la factura sí usa IA (Gemini, multimodal — reutiliza la
# misma clave [gemini] api_key que ya tienes configurada para el Asistente IA).
# La detección de anomalías es intencionalmente por reglas simples y explicables
# (no un modelo de IA): así cada alerta tiene un motivo concreto que un humano puede
# verificar, en vez de una "caja negra" diciendo que algo se ve raro.
# =====================================================================================

def extraer_datos_factura_gemini(archivo_bytes, mime_type, api_key):
    """Envía la imagen/PDF de una factura a Gemini y devuelve un diccionario con los
    campos extraídos, o None si la extracción falla."""
    try:
        from google import genai
        from google.genai import types
        import json as json_lib

        client = genai.Client(api_key=api_key)
        prompt = (
            "Eres un asistente que extrae datos de facturas y comprobantes de gasto escolares. "
            "Analiza la imagen o documento adjunto y responde ÚNICAMENTE con un objeto JSON "
            "(sin texto adicional, sin explicaciones, sin backticks de markdown), con "
            "exactamente estas claves:\n"
            '{"concepto": "descripción breve de qué se compró o pagó", '
            '"valor": numero_sin_simbolos_ni_comas_ni_texto, '
            '"fecha": "YYYY-MM-DD si es visible en el documento, o null si no se ve", '
            '"categoria_sugerida": "una de estas exactas: Logística, Publicidad, Alimentación, Varios", '
            '"numero_factura": "el número o folio de la factura si es visible, o null"}'
        )
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=[
                types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type),
                prompt,
            ],
        )
        texto_respuesta = response.text.strip().replace("```json", "").replace("```", "").strip()
        return json_lib.loads(texto_respuesta)
    except Exception as e:
        st.error(f"No se pudieron extraer los datos de la factura automáticamente: {e}")
        return None


def detectar_anomalias_gasto(nuevo_valor, nueva_fecha, nuevo_concepto, nueva_categoria,
                               gastos_existentes_df, presupuesto_tope_general, presupuestos_categoria):
    """Aplica reglas simples sobre un gasto candidato (duplicidad, desviación de
    presupuesto general, desviación de presupuesto por categoría, monto atípico frente
    al historial). Devuelve una lista de motivos en texto plano — lista vacía significa
    que no se detectó ninguna anomalía."""
    import difflib

    motivos = []
    if gastos_existentes_df is None or gastos_existentes_df.empty:
        gastos_existentes_df = pd.DataFrame(columns=["Fecha", "Concepto", "Categoría", "Valor"])

    df_temp = gastos_existentes_df.copy()
    df_temp["Valor"] = pd.to_numeric(df_temp["Valor"], errors='coerce').fillna(0)

    # --- Duplicidad: mismo valor, fecha cercana (±5 días) y concepto parecido ---
    if not df_temp.empty:
        df_temp["FechaDT"] = pd.to_datetime(df_temp["Fecha"], errors='coerce')
        fecha_nueva_dt = pd.to_datetime(nueva_fecha, errors='coerce')
        for _, fila in df_temp.iterrows():
            mismo_valor = abs(fila["Valor"] - nuevo_valor) < 1
            fecha_cercana = (
                pd.notna(fila["FechaDT"]) and pd.notna(fecha_nueva_dt)
                and abs((fila["FechaDT"] - fecha_nueva_dt).days) <= 5
            )
            similitud = difflib.SequenceMatcher(None, str(fila["Concepto"]).lower(), str(nuevo_concepto).lower()).ratio()
            if mismo_valor and fecha_cercana and similitud > 0.55:
                motivos.append(f"Posible duplicado: se parece a '{fila['Concepto']}' del {fila['Fecha']}, por el mismo valor.")
                break

    # --- Desviación del presupuesto general ---
    if presupuesto_tope_general and presupuesto_tope_general > 0:
        total_actual = df_temp["Valor"].sum()
        if (total_actual + nuevo_valor) > presupuesto_tope_general:
            motivos.append("Supera el presupuesto general disponible.")

    # --- Desviación del presupuesto asignado a la categoría ---
    presupuesto_cat = (presupuestos_categoria or {}).get(nueva_categoria, 0.0)
    if presupuesto_cat and presupuesto_cat > 0:
        gasto_cat_actual = df_temp[df_temp.get("Categoría") == nueva_categoria]["Valor"].sum() if "Categoría" in df_temp.columns else 0.0
        if (gasto_cat_actual + nuevo_valor) > presupuesto_cat:
            motivos.append(f"Supera el presupuesto asignado a la categoría '{nueva_categoria}'.")

    # --- Monto atípico frente al historial de esa categoría (requiere al menos 3 datos) ---
    if "Categoría" in df_temp.columns:
        df_cat_hist = df_temp[df_temp["Categoría"] == nueva_categoria]
        if len(df_cat_hist) >= 3:
            promedio_cat = df_cat_hist["Valor"].mean()
            if promedio_cat > 0 and nuevo_valor > promedio_cat * 2.5:
                motivos.append(f"Monto inusualmente alto frente al promedio histórico de la categoría (${promedio_cat:,.0f}).")

    return motivos


# =====================================================================================
# REQUISITO 2: AUDITORÍA AUTOMÁTICA
# -------------------------------------------------------------------------------------
# Se registra en usuarios/{institucion_id}/auditoria cada vez que se llama a esta
# función. Se llama explícitamente tras cada acción clave (alta/baja de ingreso o
# gasto, subida de archivo) en los puntos donde antes solo se hacía el guardado.
#
# El antiguo gate por contraseña hardcodeada en texto plano
# ("El amor que vale 123") se elimina: además de ser una mala práctica de seguridad
# (cualquiera que lea el código fuente tiene acceso), ya no tiene sentido una vez que
# los datos están aislados por institución — el propio login ya determina qué
# auditoría puede ver cada usuario.
# =====================================================================================

def registrar_auditoria(accion, detalle=""):
    if not db:
        return
    institucion_id = get_institucion_id()
    if not institucion_id or not st.session_state.user_data:
        return
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


# =====================================================================================
# REQUISITO 3: LOGIN CON GOOGLE (OAuth 2.0 directo)
# -------------------------------------------------------------------------------------
# NOTA DE ARQUITECTURA (léela antes de desplegar):
# Firebase Admin SDK (el que ya usa esta app) NO puede iniciar un flujo de consentimiento
# interactivo de Google — eso es responsabilidad de un cliente OAuth. Existen dos caminos
# razonables; aquí se implementa el Enfoque A por no depender de componentes JS de
# terceros para algo tan sensible como el login:
#
#   Enfoque A (implementado): OAuth 2.0 "Authorization Code" directo con Google,
#   usando el paquete 'requests' para el intercambio de tokens (evita añadir
#   'google-auth-oauthlib' como dependencia dura si no la tienes ya instalada).
#   Streamlit no maneja rutas HTTP propias, así que el 'code' de vuelta se lee desde
#   st.query_params tras la redirección de Google a esta misma URL.
#
#   Enfoque B (alternativa, no implementada): Firebase Authentication con SDK JS
#   embebido vía components.html + un componente puente JS→Python. Da mejor gestión
#   de sesión/revocación pero depende de paquetes comunitarios (streamlit-firebase-auth
#   u otros) que cambian de mantenimiento con frecuencia — riesgo alto para producción.
#
# REQUISITOS PARA ACTIVAR ESTO EN PRODUCCIÓN:
#   1. Crear un proyecto OAuth 2.0 en Google Cloud Console (APIs & Services > Credentials).
#   2. Tipo de aplicación: "Web application".
#   3. Authorized redirect URI: la URL pública exacta de tu app Streamlit
#      (ej. https://tuapp.streamlit.app) — debe coincidir carácter por carácter.
#   4. Agregar a tus secrets de Streamlit:
#        [google_oauth]
#        client_id = "TU_CLIENT_ID.apps.googleusercontent.com"
#        client_secret = "TU_CLIENT_SECRET"
#        redirect_uri = "https://tuapp.streamlit.app"
#   Sin estos 3 valores en st.secrets["google_oauth"], el botón de Google se
#   desactiva automáticamente y solo queda visible el login tradicional (no rompe
#   la app si no lo configuras todavía).
# =====================================================================================

GOOGLE_OAUTH_DISPONIBLE = "google_oauth" in st.secrets

def construir_url_login_google():
    if not GOOGLE_OAUTH_DISPONIBLE:
        return None
    from urllib.parse import urlencode

    client_id = st.secrets["google_oauth"]["client_id"]
    redirect_uri = st.secrets["google_oauth"]["redirect_uri"]
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": "openid email profile",
        "access_type": "online",
        "prompt": "select_account",
    }
    return "https://accounts.google.com/o/oauth2/v2/auth?" + urlencode(params)


def procesar_callback_google():
    """Si Google acaba de redirigir con ?code=..., intercambia el code por el perfil
    del usuario, crea la cuenta si no existe, y abre sesión."""
    if not GOOGLE_OAUTH_DISPONIBLE:
        return
    query_params = st.query_params
    if "code" not in query_params:
        return

    import requests

    codigo = query_params["code"]
    client_id = st.secrets["google_oauth"]["client_id"]
    client_secret = st.secrets["google_oauth"]["client_secret"]
    redirect_uri = st.secrets["google_oauth"]["redirect_uri"]

    try:
        token_resp = requests.post(
            "https://oauth2.googleapis.com/token",
            data={
                "code": codigo,
                "client_id": client_id,
                "client_secret": client_secret,
                "redirect_uri": redirect_uri,
                "grant_type": "authorization_code",
            },
            timeout=10,
        )
        token_resp.raise_for_status()
        access_token = token_resp.json()["access_token"]

        perfil_resp = requests.get(
            "https://www.googleapis.com/oauth2/v2/userinfo",
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=10,
        )
        perfil_resp.raise_for_status()
        perfil = perfil_resp.json()

        email_google = perfil.get("email", "").lower().strip()
        nombre_google = perfil.get("name", email_google)

        if not email_google:
            st.error("Google no devolvió un correo válido.")
            st.query_params.clear()
            return

        if not db:
            st.error("No hay conexión con la base de datos.")
            return

        user_ref = db.collection("usuarios").document(email_google)
        user_doc = user_ref.get()

        if user_doc.exists:
            user_data = user_doc.to_dict()
        else:
            # Cuenta nueva vía Google: sin password local (login_provider marca el origen)
            user_data = {
                'institucion': nombre_google,
                'email': email_google,
                'password': hash_password(''.join(random.choices(string.ascii_letters + string.digits, k=24))),
                'login_provider': 'google',
                'fecha_creacion': dt_module.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            user_ref.set(user_data)

        st.session_state.logged_in = True
        st.session_state.user_data = user_data
        st.query_params.clear()
        cargar_datos_nube()
        st.rerun()

    except Exception as e:
        st.error(f"Error al validar el inicio de sesión con Google: {e}")
        st.query_params.clear()


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
    except Exception:
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


# =====================================================================================
# EXPORTACIÓN: REPORTE PDF CORPORATIVO
# -------------------------------------------------------------------------------------
# Formal y con la identidad visual del proyecto (mismo azul institucional #1E3A8A),
# pero sin firma digital ni certificación legal — eso requiere un certificado real de
# una autoridad certificadora, que está fuera del alcance de lo que este código puede
# proveer por sí solo.
# =====================================================================================

def _estilo_tabla_corporativo():
    from reportlab.platypus import TableStyle
    from reportlab.lib import colors
    return TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ])


def generar_reporte_pdf_corporativo(nombre_institucion, df_ingresos, df_gastos, df_desviacion):
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    buffer_pdf = BytesIO()
    doc = SimpleDocTemplate(
        buffer_pdf, pagesize=letter,
        topMargin=2.2 * cm, bottomMargin=2 * cm, leftMargin=1.8 * cm, rightMargin=1.8 * cm
    )
    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle('TituloCorp', parent=styles['Title'], textColor=colors.HexColor('#1E3A8A'), fontSize=20)
    estilo_subtitulo = ParagraphStyle('SubtituloCorp', parent=styles['Normal'], textColor=colors.HexColor('#4B5563'), fontSize=10)
    estilo_seccion = ParagraphStyle('SeccionCorp', parent=styles['Heading2'], textColor=colors.HexColor('#1E3A8A'), spaceBefore=14, spaceAfter=8)

    total_ing = pd.to_numeric(df_ingresos["Valor"], errors='coerce').sum() if not df_ingresos.empty else 0.0
    total_gas = pd.to_numeric(df_gastos["Valor"], errors='coerce').sum() if not df_gastos.empty else 0.0
    balance = total_ing - total_gas
    estado = "Superávit" if balance >= 0 else "Déficit"

    story = [
        Paragraph(nombre_institucion.upper(), estilo_titulo),
        Paragraph("Reporte Financiero Corporativo", estilo_subtitulo),
        Paragraph(f"Generado el {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", estilo_subtitulo),
        Spacer(1, 16),
        Paragraph("Balance General", estilo_seccion),
    ]

    tabla_balance = Table(
        [["Concepto", "Valor"],
         ["Total Ingresos", f"${total_ing:,.0f}"],
         ["Total Gastos", f"${total_gas:,.0f}"],
         [f"Balance Neto ({estado})", f"${balance:,.0f}"]],
        colWidths=[10 * cm, 6 * cm]
    )
    tabla_balance.setStyle(_estilo_tabla_corporativo())
    story.append(tabla_balance)

    if df_desviacion is not None and not df_desviacion.empty:
        story.append(Paragraph("Desviaciones Presupuestarias por Categoría", estilo_seccion))
        filas_desv = [["Categoría", "Gasto Real", "Presupuesto", "Desviación"]]
        for _, fila in df_desviacion.iterrows():
            filas_desv.append([
                fila["Categoría"], f"${fila['Gasto Real']:,.0f}",
                f"${fila['Presupuesto Asignado']:,.0f}", f"${fila['Desviación']:,.0f}",
            ])
        tabla_desv = Table(filas_desv, colWidths=[4 * cm, 4 * cm, 4 * cm, 4 * cm])
        tabla_desv.setStyle(_estilo_tabla_corporativo())
        story.append(tabla_desv)

    story.append(PageBreak())
    story.append(Paragraph("Detalle de Ingresos", estilo_seccion))
    if not df_ingresos.empty:
        columnas_ing = [c for c in df_ingresos.columns if c != "ID"]
        filas_ing = [columnas_ing] + df_ingresos[columnas_ing].astype(str).values.tolist()
        tabla_ing = Table(filas_ing, repeatRows=1)
        tabla_ing.setStyle(_estilo_tabla_corporativo())
        story.append(tabla_ing)
    else:
        story.append(Paragraph("No hay ingresos registrados.", styles['Normal']))

    story.append(Spacer(1, 16))
    story.append(Paragraph("Detalle de Gastos", estilo_seccion))
    if not df_gastos.empty:
        columnas_gas = [c for c in df_gastos.columns if c != "ID"]
        filas_gas = [columnas_gas] + df_gastos[columnas_gas].astype(str).values.tolist()
        tabla_gas = Table(filas_gas, repeatRows=1)
        tabla_gas.setStyle(_estilo_tabla_corporativo())
        story.append(tabla_gas)
    else:
        story.append(Paragraph("No hay gastos registrados.", styles['Normal']))

    def pie_de_pagina(canvas_obj, doc_obj):
        canvas_obj.saveState()
        canvas_obj.setFont('Helvetica', 8)
        canvas_obj.setFillColor(colors.HexColor('#94A3B8'))
        canvas_obj.drawString(1.8 * cm, 1.2 * cm, nombre_institucion)
        canvas_obj.drawRightString(letter[0] - 1.8 * cm, 1.2 * cm, f"Página {doc_obj.page}")
        canvas_obj.restoreState()

    doc.build(story, onFirstPage=pie_de_pagina, onLaterPages=pie_de_pagina)
    buffer_pdf.seek(0)
    return buffer_pdf


# --- PANTALLAS DE AUTENTICACIÓN ---
if not st.session_state.logged_in:
    # Procesa el retorno de Google (si Google acaba de redirigir con ?code=...)
    # ANTES de dibujar el formulario, para que si ya hay sesión válida no se
    # muestre el login de nuevo.
    procesar_callback_google()

    st.markdown('<p class="main-header" style="text-align: center;">🏛️ Portal Financiero Institucional</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header" style="text-align: center;">Colegio Francisco de Paula Santander</p>', unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["Iniciar Sesión", "Crear Cuenta", "Olvidé mi Contraseña"])

    with tab1:
        st.markdown("### Acceso con Credenciales")

        # --- Botón de Google (Requisito 3) ---
        if GOOGLE_OAUTH_DISPONIBLE:
            url_google = construir_url_login_google()
            # st.link_button es el widget nativo de Streamlit para navegación externa.
            # Un <a href> metido a mano vía st.markdown puede quedar interceptado por el
            # manejo de clics propio de Streamlit (clic izquierdo normal no navega, aunque
            # "abrir en pestaña nueva" sí funciona) — este widget evita ese problema.
            st.link_button("🔵 Iniciar sesión con Google", url_google, use_container_width=True)
            st.markdown("<p style='text-align:center; color:#94A3B8; font-size:0.85rem;'>— o con tu correo y contraseña —</p>", unsafe_allow_html=True)
        else:
            st.caption("ℹ️ El login con Google no está configurado aún (faltan secrets [google_oauth]). Usa correo y contraseña.")

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

import datetime

import datetime

# --- INICIALIZAR ESTADO DE OMITIR ALERTA ---
if "omitir_alerta_presupuesto" not in st.session_state:
    st.session_state.omitir_alerta_presupuesto = False

# --- MENÚ LATERAL ---
st.sidebar.markdown(f"👋 **Hola, {st.session_state.user_data['institucion']}**")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.user_data = None
    st.session_state.omitir_alerta_presupuesto = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("⚙️ **Configuración de Presupuesto**")

hoy = datetime.date.today()
fin_estimado = hoy + datetime.timedelta(days=30)

periodo_presupuesto = st.sidebar.date_input(
    "📅 Período de Ejecución",
    value=(hoy, fin_estimado)
)

presupuesto_tope = st.sidebar.number_input("Presupuesto / Límite de Gastos ($)", min_value=0.0, value=500000.0, step=50000.0)

if isinstance(periodo_presupuesto, tuple) and len(periodo_presupuesto) == 2:
    fecha_fin = periodo_presupuesto[1]

    if hoy > fecha_fin:
        if not st.session_state.omitir_alerta_presupuesto:
            with st.sidebar.container():
                st.warning("⚠️ **¡Tiempo finalizado!**\n\nEl período de tu presupuesto ha terminado. Por favor, asigna uno nuevo.")
                if st.button("Omitir por ahora"):
                    st.session_state.omitir_alerta_presupuesto = True
                    st.rerun()
    else:
        st.session_state.omitir_alerta_presupuesto = False

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
    "9. Auditoría del Sistema",
    "10. Facturas (OCR) y Anomalías",
])
# --- APARTADO DE IA EN EL BORDE ---
st.sidebar.markdown("---")
st.sidebar.markdown("🤖 **Asistente IA del Borde**")
if st.sidebar.button("💬 Abrir / Cerrar Asistente IA"):
    st.session_state.ia_abierta = not st.session_state.ia_abierta

if st.session_state.ia_abierta:
    with st.sidebar.container():
        st.markdown("### 🧠 Chat Asesor IA")

        # La clave ya no se pide en pantalla — se lee de tus secrets ([gemini] api_key = "...").
        gemini_api_key = st.secrets.get("gemini", {}).get("api_key") if "gemini" in st.secrets else None

        if not gemini_api_key:
            st.warning("⚠️ Falta configurar la clave de Gemini en los secrets del proyecto (sección [gemini]).")
        else:
            pregunta_ia = st.text_input("¿Qué deseas consultar?")

            if st.button("Consultar IA"):
                if not pregunta_ia.strip():
                    st.error("⚠️ Escribe una pregunta antes de consultar.")
                else:
                    try:
                        from google import genai
                        client = genai.Client(api_key=gemini_api_key)

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
import datetime

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
                f_ing = st.date_input("Fecha", value=datetime.date.today())
                con_ing = st.text_input("Concepto")
            with c2:
                resp_ing = st.selectbox("Responsable", INTEGRANTES_LISTA)
                val_ing = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
            obs_ing = st.text_area("Observaciones (Opcional)")

            if st.form_submit_button("Guardar Ingreso"):
                if con_ing.strip() == "":
                    st.error("⚠️ El concepto no puede estar vacío.")
                else:
                    reg_id = f"ING-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
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
                    registrar_auditoria("Registró Ingreso", f"{con_ing} - ${val_ing:,.0f} (Responsable: {resp_ing})")
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
            fila = st.session_state.ingresos_df.iloc[idx]
            doc_id = fila['ID']
            eliminar_registro_nube('ingresos', doc_id)
            registrar_auditoria("Eliminó Ingreso", f"{fila['Concepto']} - ${float(fila['Valor']):,.0f}")
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
                f_gas = st.date_input("Fecha Gasto", value=datetime.date.today())
                con_gas = st.text_input("Concepto")
                cat_gas = st.selectbox("Categoría", CATEGORIAS_GASTO)
            with c2:
                val_gas = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
                resp_gas = st.selectbox("Responsable", INTEGRANTES_LISTA)

            if st.form_submit_button("Guardar Gasto"):
                if con_gas.strip() == "":
                    st.error("⚠️ El concepto no puede estar vacío.")
                else:
                    reg_id = f"GAS-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
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
                    registrar_auditoria("Registró Gasto", f"{con_gas} - ${val_gas:,.0f} (Categoría: {cat_gas}, Responsable: {resp_gas})")
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
            fila = st.session_state.gastos_df.iloc[idx]
            doc_id = fila['ID']
            eliminar_registro_nube('gastos', doc_id)
            registrar_auditoria("Eliminó Gasto", f"{fila['Concepto']} - ${float(fila['Valor']):,.0f}")
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

    tot_ing = pd.to_numeric(st.session_state.ingresos_df["Valor"], errors='coerce').sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = pd.to_numeric(st.session_state.gastos_df["Valor"], errors='coerce').sum() if not st.session_state.gastos_df.empty else 0.0

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Ingresos vs Gastos")
        df_comp = pd.DataFrame({"Tipo": ["Ingresos", "Gastos"], "Monto": [tot_ing, tot_gas]})
        fig_bar = px.bar(df_comp, x="Tipo", y="Monto", color="Tipo", text_auto=True, color_discrete_sequence=["#10B981", "#EF4444"])
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("#### Gastos por Categoría")
        if not st.session_state.gastos_df.empty:
            df_gastos_temp = st.session_state.gastos_df.copy()
            df_gastos_temp["Valor"] = pd.to_numeric(df_gastos_temp["Valor"], errors='coerce').fillna(0)
            df_cat = df_gastos_temp.groupby("Categoría")["Valor"].sum().reset_index()
            fig_pie = px.pie(df_cat, names="Categoría", values="Valor", hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No hay gastos registrados aún.")

    st.markdown("#### 📈 Evolución Temporal de Movimientos")
    df_all = []
    if not st.session_state.ingresos_df.empty:
        df_i = st.session_state.ingresos_df[["Fecha", "Valor"]].copy()
        df_i["Valor"] = pd.to_numeric(df_i["Valor"], errors='coerce').fillna(0)
        df_i["Tipo"] = "Ingreso"
        df_all.append(df_i)

    if not st.session_state.gastos_df.empty:
        df_g = st.session_state.gastos_df[["Fecha", "Valor"]].copy()
        df_g["Valor"] = pd.to_numeric(df_g["Valor"], errors='coerce').fillna(0)
        df_g["Tipo"] = "Gasto"
        df_all.append(df_g)

    if df_all:
        df_timeline = pd.concat(df_all, ignore_index=True)
        df_timeline["Fecha"] = pd.to_datetime(df_timeline["Fecha"], errors='coerce')
        df_timeline = df_timeline.dropna(subset=["Fecha"]).sort_values("Fecha")

        if not df_timeline.empty:
            fig_line = px.line(df_timeline, x="Fecha", y="Valor", color="Tipo", markers=True, color_discrete_map={"Ingreso": "#10B981", "Gasto": "#EF4444"})
            st.plotly_chart(fig_line, use_container_width=True)
        else:
            st.info("Las fechas registradas no tienen un formato válido para la línea de tiempo.")
    else:
        st.info("Registra ingresos o gastos para ver la línea de tiempo.")

    # =====================================================================================
    # TABLERO EJECUTIVO: desviaciones presupuestarias por categoría, con drill-down.
    # Requiere Streamlit >= 1.35 para las selecciones de gráfico (on_select="rerun").
    # Si tu versión es anterior, el gráfico se ve igual pero el clic no filtra la tabla.
    # =====================================================================================
    st.markdown("---")
    st.markdown('<p class="main-header" style="font-size:1.6rem;">🎯 Tablero Ejecutivo — Desviaciones Presupuestarias</p>', unsafe_allow_html=True)

    with st.expander("⚙️ Configurar presupuesto asignado por categoría"):
        presupuestos_actuales = cargar_presupuestos_categoria()
        nuevos_presupuestos = {}
        cols_presupuesto = st.columns(len(CATEGORIAS_GASTO))
        for col, categoria in zip(cols_presupuesto, CATEGORIAS_GASTO):
            with col:
                nuevos_presupuestos[categoria] = st.number_input(
                    categoria, min_value=0.0, value=presupuestos_actuales.get(categoria, 0.0),
                    step=50000.0, key=f"presupuesto_{categoria}"
                )
        if st.button("💾 Guardar presupuestos por categoría"):
            guardar_presupuestos_categoria(nuevos_presupuestos)
            registrar_auditoria("Actualizó presupuestos por categoría", str(nuevos_presupuestos))
            st.success("Presupuestos actualizados.")
            st.rerun()

    presupuestos_categoria = cargar_presupuestos_categoria()

    "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),y:
        df_gastos_cat = st.session_state.gastos_df.copy()
        df_gastos_cat["Valor"] = pd.to_numeric(df_gastos_cat["Valor"], errors='coerce').fillna(0)
        df_desviacion = calcular_tabla_desviaciones(st.session_state.gastos_df, presupuestos_categoria)

        fig_desviacion = px.bar(
            df_desviacion, x="Categoría", y=["Gasto Real", "Presupuesto Asignado"],
            barmode="group", text_auto=True,
            color_discrete_map={"Gasto Real": "#EF4444", "Presupuesto Asignado": "#94A3B8"},
            title="Gasto real vs. presupuesto asignado por categoría — haz clic en una barra para ver el detalle"
        )

        try:
            evento = st.plotly_chart(
                fig_desviacion, use_container_width=True,
                on_select="rerun", selection_mode="points", key="chart_desviaciones"
            )
            categoria_seleccionada = None
            puntos = evento.get("selection", {}).get("points", []) if evento else []
            if puntos:
                categoria_seleccionada = puntos[0].get("x")
        except TypeError:
            # Streamlit < 1.35: no soporta on_select. Se muestra el gráfico sin drill-down.
            st.plotly_chart(fig_desviacion, use_container_width=True)
            categoria_seleccionada = None
            st.caption("ℹ️ Actualiza Streamlit a la versión 1.35 o superior para habilitar el drill-down por clic.")

        st.dataframe(
            df_desviacion.style.apply(
                lambda fila: ['background-color: #FEE2E2' if fila["Desviación"] > 0 else '' for _ in fila], axis=1
            ),
            use_container_width=True, hide_index=True
        )

        if categoria_seleccionada:
            st.markdown(f"#### 🔍 Detalle de transacciones — {categoria_seleccionada}")
            detalle_cat = df_gastos_cat[df_gastos_cat["Categoría"] == categoria_seleccionada].drop(columns=['ID'], errors='ignore')
            if not detalle_cat.empty:
                st.dataframe(detalle_cat, use_container_width=True, hide_index=True)
            else:
                st.info("No hay transacciones registradas en esta categoría todavía.")
    else:
        st.info("Registra gastos para ver el tablero de desviaciones presupuestarias.")

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

                nombre_id = f"ARCH-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
                institucion_id = get_institucion_id()
                doc_data = {
                    "ID": nombre_id,
                    "nombre": archivo_subido.name,
                    "tipo": archivo_subido.type,
                    "archivo_b64": base64_archivo,
                    "descripcion": descripcion_archivo if descripcion_archivo else "Sin descripción",
                    "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "subido_por": st.session_state.user_data['institucion']
                }
                db.collection("usuarios").document(institucion_id).collection("archivos").document(nombre_id).set(doc_data)
                registrar_auditoria("Subió Archivo", f"{archivo_subido.name} — {descripcion_archivo or 'Sin descripción'}")
                st.success("✅ ¡Archivo guardado exitosamente en la base de datos!")
                st.rerun()
            else:
                st.error("⚠️ Por favor selecciona un archivo antes de guardar.")

    st.markdown("### 📋 Archivos Registrados con Vista Previa")
    if db:
        try:
            institucion_id = get_institucion_id()
            archivos_ref = db.collection("usuarios").document(institucion_id).collection("archivos").stream()
            archivos_lista = [a.to_dict() for a in archivos_ref]

            if archivos_lista:
                for row in archivos_lista:
                    with st.expander(f"📄 {row['nombre']} — ({row['fecha']})"):
                        st.write(f"**Descripción:** {row['descripcion']}")
                        st.write(f"**Subido por:** {row['subido_por']}")

                        if "archivo_b64" in row:
                            b64_bytes = base64.b64decode(row['archivo_b64'])

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
                    archivo_a_eliminar = archivos_lista[idx]
                    doc_id_eliminar = archivo_a_eliminar['ID']
                    db.collection("usuarios").document(institucion_id).collection("archivos").document(doc_id_eliminar).delete()
                    registrar_auditoria("Eliminó Archivo", archivo_a_eliminar['nombre'])
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
        import pandas as pd
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

                    if headers[col_num - 1] in ("Valor", "Gasto Real", "Presupuesto Asignado", "Desviación") and isinstance(cell.value, (int, float)):
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

        df_ing = st.session_state.ingresos_df
        df_gas = st.session_state.gastos_df

        total_ingresos = float(df_ing['Valor'].sum()) if not df_ing.empty and 'Valor' in df_ing.columns else 0.0
        total_gastos = float(df_gas['Valor'].sum()) if not df_gas.empty and 'Valor' in df_gas.columns else 0.0
        balance_neto = total_ingresos - total_gastos

        estado = "Superávit (Ganancia)" if balance_neto >= 0 else "Déficit (Pérdida)"

        df_balance = pd.DataFrame({
            "Concepto": ["Total Ingresos Recaudados", "Total Gastos Ejecutados", f"Balance Neto: {estado}"],
            "Valor": [total_ingresos, total_gastos, balance_neto]
        })

        ws_bal = wb.create_sheet(title="Balance General", index=0)
        estilizar_hoja(ws_bal, "RESUMEN DE BALANCE GENERAL", df_balance)

        ws_ing = wb.create_sheet(title="Ingresos")
        estilizar_hoja(ws_ing, "REPORTE DE INGRESOS", df_ing)

        ws_gas = wb.create_sheet(title="Gastos")
        estilizar_hoja(ws_gas, "REPORTE DE GASTOS", df_gas)

        # --- Hoja adicional: Desviaciones por Categoría, con gráfico de barras embebido ---
        from openpyxl.chart import BarChart, Reference

        presupuestos_cat_excel = cargar_presupuestos_categoria()
        df_desv_excel = calcular_tabla_desviaciones(df_gas, presupuestos_cat_excel)
        df_desv_excel_tabla = df_desv_excel[["Categoría", "Gasto Real", "Presupuesto Asignado"]].copy()

        ws_desv = wb.create_sheet(title="Desviaciones por Categoría")
        estilizar_hoja(ws_desv, "DESVIACIONES PRESUPUESTARIAS POR CATEGORÍA", df_desv_excel_tabla)
        # estilizar_hoja escribe: fila 1 = título, fila 2 = vacía, fila 3 = encabezados,
        # filas 4..N = datos. Como esta tabla no tiene columna 'ID', las 3 columnas
        # (Categoría, Gasto Real, Presupuesto Asignado) ocupan A, B y C tal cual.

        grafico_desv = BarChart()
        grafico_desv.title = "Gasto Real vs. Presupuesto Asignado"
        grafico_desv.y_axis.title = "Valor ($)"
        grafico_desv.x_axis.title = "Categoría"
        fila_encabezado = 3
        fila_fin = 3 + len(df_desv_excel_tabla)
        datos_grafico = Reference(ws_desv, min_col=2, max_col=3, min_row=fila_encabezado, max_row=fila_fin)
        categorias_grafico = Reference(ws_desv, min_col=1, min_row=fila_encabezado + 1, max_row=fila_fin)
        grafico_desv.add_data(datos_grafico, titles_from_data=True)
        grafico_desv.set_categories(categorias_grafico)
        grafico_desv.width, grafico_desv.height = 18, 10
        ws_desv.add_chart(grafico_desv, f"A{fila_fin + 3}")

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

    st.markdown("---")
    st.markdown("### 📄 Reporte PDF Corporativo")
    st.caption("Documento formal con la identidad visual del proyecto — no incluye firma digital ni certificación legal.")

    if st.button("📄 Generar Reporte PDF Corporativo"):
        nombre_institucion_pdf = st.session_state.user_data.get('institucion', 'Institución Financiera')
        presupuestos_cat_pdf = cargar_presupuestos_categoria()
        df_desviacion_pdf = calcular_tabla_desviaciones(st.session_state.gastos_df, presupuestos_cat_pdf)
        buffer_pdf = generar_reporte_pdf_corporativo(
            nombre_institucion_pdf, st.session_state.ingresos_df, st.session_state.gastos_df, df_desviacion_pdf
        )
        st.session_state.pdf_reporte_bytes = buffer_pdf.getvalue()
        registrar_auditoria("Generó Reporte PDF Corporativo")
        st.success("✅ ¡Reporte PDF generado!")

    if 'pdf_reporte_bytes' in st.session_state:
        st.download_button(
            "📥 Descargar Reporte PDF",
            data=st.session_state.pdf_reporte_bytes,
            file_name="Reporte_Financiero_Corporativo.pdf",
            mime="application/pdf"
        )
elif menu == "9. Auditoría del Sistema":
    # --------------------------------------------------------------------------------
    # REQUISITO 2 (vista): ya no depende de una contraseña hardcodeada. El acceso está
    # controlado por el login (cada institución solo puede leer su propia subcolección
    # usuarios/{institucion_id}/auditoria), que es exactamente el aislamiento pedido en
    # el requisito 1. Cualquier usuario autenticado ve su propio historial en tiempo real.
    # --------------------------------------------------------------------------------
    st.markdown('<p class="main-header">🔐 Auditoría y Registro de Actividad (Logs)</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Historial de control y seguridad institucional.</p>', unsafe_allow_html=True)
    st.markdown("---")

    if db:
        try:
            institucion_id = get_institucion_id()
            logs_ref = (
                db.collection("usuarios").document(institucion_id)
                .collection("auditoria")
                .order_by("fecha_hora", direction=firestore.Query.DESCENDING)
                .stream()
            )
            logs_lista = [l.to_dict() for l in logs_ref]

            if logs_lista:
                df_logs = pd.DataFrame(logs_lista)
                columnas_orden = [c for c in ["fecha_hora", "usuario", "accion", "detalle"] if c in df_logs.columns]
                otras = [c for c in df_logs.columns if c not in columnas_orden]
                st.dataframe(df_logs[columnas_orden + otras], use_container_width=True, hide_index=True)
            else:
                st.info("🔒 Aún no hay movimientos registrados para tu institución. Las acciones (ingresos, gastos, archivos) se registrarán aquí automáticamente.")
        except Exception as e:
            st.warning(f"No se pudieron cargar los registros de auditoría: {e}")
    else:
        st.warning("Conecta Firebase para habilitar la auditoría en la nube.")

elif menu == "10. Facturas (OCR) y Anomalías":
    st.markdown('<p class="main-header">🧾 Facturas (OCR) y Detección de Anomalías</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Sube una factura — la IA extrae los datos y las reglas del sistema deciden si se registra directo o queda pendiente de revisión.</p>', unsafe_allow_html=True)
    st.markdown("---")

    gemini_api_key = st.secrets.get("gemini", {}).get("api_key") if "gemini" in st.secrets else None

    if not gemini_api_key:
        st.warning("⚠️ Falta configurar la clave de Gemini en los secrets del proyecto (sección [gemini]) — es la misma que usa el Asistente IA de la barra lateral.")
    else:
        archivo_factura = st.file_uploader("Sube la foto o PDF de la factura", type=["png", "jpg", "jpeg", "pdf"], key="upload_factura_ocr")

        if archivo_factura is not None:
            if st.button("🔎 Extraer datos con IA"):
                mime_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "pdf": "application/pdf"}
                extension = archivo_factura.name.split(".")[-1].lower()
                datos_extraidos = extraer_datos_factura_gemini(
                    archivo_factura.getvalue(), mime_map.get(extension, "image/jpeg"), gemini_api_key
                )
                if datos_extraidos:
                    st.session_state.factura_extraida = datos_extraidos
                    st.success("✅ Datos extraídos. Revísalos y corrígelos si algo no quedó bien antes de procesar.")

        if "factura_extraida" in st.session_state:
            st.markdown("#### ✏️ Confirma o corrige los datos antes de procesar")
            datos = st.session_state.factura_extraida
            c1, c2 = st.columns(2)
            with c1:
                concepto_confirmado = st.text_input("Concepto", value=str(datos.get("concepto", "")))
                try:
                    valor_default = float(datos.get("valor") or 0)
                except (TypeError, ValueError):
                    valor_default = 0.0
                valor_confirmado = st.number_input("Valor ($)", min_value=0.0, value=valor_default, step=1000.0)
            with c2:
                fecha_texto = datos.get("fecha")
                try:
                    fecha_default = datetime.datetime.strptime(fecha_texto, "%Y-%m-%d").date() if fecha_texto else datetime.date.today()
                except (ValueError, TypeError):
                    fecha_default = datetime.date.today()
                fecha_confirmada = st.date_input("Fecha", value=fecha_default)
                categoria_sugerida = datos.get("categoria_sugerida") if datos.get("categoria_sugerida") in CATEGORIAS_GASTO else "Varios"
                categoria_confirmada = st.selectbox("Categoría", CATEGORIAS_GASTO, index=CATEGORIAS_GASTO.index(categoria_sugerida))
            responsable_confirmado = st.selectbox("Responsable", INTEGRANTES_LISTA)

            if st.button("✅ Procesar factura"):
                presupuestos_cat_actuales = cargar_presupuestos_categoria()
                motivos = detectar_anomalias_gasto(
                    valor_confirmado, fecha_confirmada.strftime("%Y-%m-%d"), concepto_confirmado, categoria_confirmada,
                    st.session_state.gastos_df, presupuesto_tope, presupuestos_cat_actuales
                )

                if not motivos:
                    reg_id = f"GAS-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
                    nuevo_reg = {
                        "ID": reg_id, "Fecha": fecha_confirmada.strftime("%Y-%m-%d"),
                        "Concepto": concepto_confirmado, "Categoría": categoria_confirmada,
                        "Valor": float(valor_confirmado), "Responsable": responsable_confirmado,
                    }
                    st.session_state.gastos_df = pd.concat([st.session_state.gastos_df, pd.DataFrame([nuevo_reg])], ignore_index=True)
                    guardar_registro_nube('gastos', nuevo_reg)
                    registrar_auditoria("Registró Gasto (factura OCR, sin anomalías)", f"{concepto_confirmado} - ${valor_confirmado:,.0f}")
                    st.success("✅ Sin anomalías detectadas — el gasto se registró directamente.")
                else:
                    institucion_id = get_institucion_id()
                    pendiente_id = f"PEND-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
                    registro_pendiente = {
                        "ID": pendiente_id, "Fecha": fecha_confirmada.strftime("%Y-%m-%d"),
                        "Concepto": concepto_confirmado, "Categoría": categoria_confirmada,
                        "Valor": float(valor_confirmado), "Responsable": responsable_confirmado,
                        "Motivos": motivos, "Estado": "Pendiente",
                        "Numero_Factura": datos.get("numero_factura"),
                    }
                    if db and institucion_id:
                        db.collection("usuarios").document(institucion_id).collection("gastos_pendientes").document(pendiente_id).set(registro_pendiente)
                    registrar_auditoria("Factura bloqueada para revisión", f"{concepto_confirmado} - ${valor_confirmado:,.0f} — Motivos: {'; '.join(motivos)}")
                    st.warning("🚫 Se detectaron anomalías — el gasto quedó pendiente de revisión humana (no se contabilizó todavía).")
                    for motivo in motivos:
                        st.caption(f"• {motivo}")

                del st.session_state.factura_extraida
                st.rerun()

    st.markdown("---")
    st.markdown("### 🚦 Gastos Pendientes de Revisión")

    if db:
        institucion_id = get_institucion_id()
        try:
            pendientes_ref = (
                db.collection("usuarios").document(institucion_id).collection("gastos_pendientes")
                .where("Estado", "==", "Pendiente").stream()
            )
            pendientes_lista = [p.to_dict() for p in pendientes_ref]

            if not pendientes_lista:
                st.info("No hay gastos pendientes de revisión en este momento.")
            else:
                for pendiente in pendientes_lista:
                    with st.expander(f"🚫 {pendiente['Concepto']} — ${float(pendiente['Valor']):,.0f} ({pendiente['Fecha']})"):
                        st.write(f"**Categoría:** {pendiente['Categoría']}  |  **Responsable:** {pendiente['Responsable']}")
                        if pendiente.get("Numero_Factura"):
                            st.write(f"**N.° de factura:** {pendiente['Numero_Factura']}")
                        st.markdown("**Motivos de la alerta:**")
                        for motivo in pendiente.get("Motivos", []):
                            st.caption(f"• {motivo}")

                        col_aprobar, col_rechazar = st.columns(2)
                        with col_aprobar:
                            if st.button("✅ Aprobar y registrar", key=f"aprobar_{pendiente['ID']}"):
                                reg_id = f"GAS-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
                                nuevo_reg = {
                                    "ID": reg_id, "Fecha": pendiente["Fecha"], "Concepto": pendiente["Concepto"],
                                    "Categoría": pendiente["Categoría"], "Valor": float(pendiente["Valor"]),
                                    "Responsable": pendiente["Responsable"],
                                }
                                guardar_registro_nube('gastos', nuevo_reg)
                                db.collection("usuarios").document(institucion_id).collection("gastos_pendientes").document(pendiente["ID"]).update({"Estado": "Aprobado"})
                                registrar_auditoria("Aprobó Gasto Pendiente", f"{pendiente['Concepto']} - ${float(pendiente['Valor']):,.0f}")
                                st.success("Gasto aprobado y registrado.")
                                cargar_datos_nube()
                                st.rerun()
                        with col_rechazar:
                            if st.button("❌ Rechazar", key=f"rechazar_{pendiente['ID']}"):
                                db.collection("usuarios").document(institucion_id).collection("gastos_pendientes").document(pendiente["ID"]).update({"Estado": "Rechazado"})
                                registrar_auditoria("Rechazó Gasto Pendiente", f"{pendiente['Concepto']} - ${float(pendiente['Valor']):,.0f}")
                                st.info("Gasto rechazado — queda en el historial como referencia, sin contabilizarse.")
                                st.rerun()
        except Exception as e:
            st.warning(f"No se pudieron cargar los gastos pendientes: {e}")
    else:
        st.warning("Conecta Firebase para habilitar esta sección.")
