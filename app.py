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
import hashlib
import re
import tempfile
import uuid
from urllib.parse import quote

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión Financiera Empresarial",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS (tema claro/oscuro dinámico) ---
if 'modo_oscuro' not in st.session_state:
    st.session_state.modo_oscuro = False


def render_estilos(modo_oscuro: bool) -> str:
    """Devuelve el bloque <style> con selectores específicos de Streamlit
    para garantizar contraste en modo oscuro y claro.

    DIAGNÓSTICO DEL BUG REPORTADO ("se ve feo si el navegador está en modo claro"):
    Streamlit dibuja su propia barra superior (el header con el menú ☰ y el botón
    "Deploy") como un elemento aparte de `.stApp` — vive en [data-testid="stHeader"].
    Antes solo coloreábamos `.stApp`, así que esa barra se quedaba con el color que
    trae Streamlit por defecto (generalmente blanco) sin importar nuestro toggle,
    y al activar el modo oscuro quedaba una franja clara pegada encima de un cuerpo
    oscuro — de ahí lo "feo". La solución es pintar también esa barra y el resto de
    contenedores estructurales de Streamlit (toolbar, línea decorativa superior,
    contenedor de vista principal), no solo el cuerpo de la app.

    LÍMITE HONESTO: el `st.dataframe` nativo de Streamlit se dibuja con un
    componente de canvas (no HTML plano) que toma sus colores del tema interno de
    Streamlit, no de este CSS. Puede que las tablas se vean con fondo claro incluso
    en modo oscuro — eso no se puede forzar con CSS puro; requeriría fijar el tema
    real de Streamlit (archivo .streamlit/config.toml), que es estático y no se
    puede alternar en vivo con un switch como este.
    """
    if modo_oscuro:
        fondo_app = "#0F172A"
        superficie = "#1E293B"
        borde = "#334155"
        texto_principal = "#E2E8F0"
        texto_secundario = "#94A3B8"
        acento = "#3B82F6"
        acento_hover = "#60A5FA"
        input_bg = "#0F172A"
        color_scheme = "dark"
    else:
        fondo_app = "#FFFFFF"
        superficie = "#F8FAFC"
        borde = "#E2E8F0"
        texto_principal = "#1E293B"
        texto_secundario = "#4B5563"
        acento = "#1E3A8A"
        acento_hover = "#2563EB"
        input_bg = "#FFFFFF"
        color_scheme = "light"

    return f"""
    <style>
        :root {{ color-scheme: {color_scheme}; }}

        /* Cuerpo de la app */
        .stApp {{ background-color: {fondo_app}; color: {texto_principal}; }}

        /* Header/toolbar nativos de Streamlit — antes quedaban sin colorear y
           creaban la franja clara/oscura desalineada que reportaste */
        [data-testid="stHeader"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stToolbar"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stDecoration"] {{ background-image: none !important; background-color: {acento} !important; }}
        [data-testid="stAppViewContainer"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stBottomBlockContainer"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stMainBlockContainer"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stSidebarContent"] {{ background-color: {superficie} !important; }}
        [data-testid="stFileUploaderDropzone"],
        [data-testid="stExpander"] details,
        [data-testid="stForm"],
        [data-baseweb="popover"] > div,
        [data-baseweb="menu"],
        ul[role="listbox"] {{
            background-color: {superficie} !important;
            color: {texto_principal} !important;
            border-color: {borde} !important;
        }}
        [role="option"]:hover,
        [data-testid="stDateInput"] input,
        [data-testid="stNumberInput"] input,
        [data-testid="stTextInput"] input,
        [data-testid="stTextArea"] textarea {{
            background-color: {input_bg} !important;
            color: {texto_principal} !important;
            border-color: {borde} !important;
        }}
        [data-testid="stAlert"] {{
            background-color: {superficie} !important;
            color: {texto_principal} !important;
            border: 1px solid {borde} !important;
        }}
        [data-testid="stDataFrame"] {{
            border: 1px solid {borde} !important;
            border-radius: 8px;
        }}

        .main-header {{ font-size: 2.3rem; color: {acento}; font-weight: 800; margin-bottom: 0px; letter-spacing: -0.5px; }}
        .sub-header {{ font-size: 1.1rem; color: {texto_secundario}; margin-bottom: 20px; }}
        .stButton>button {{ width: 100%; border-radius: 8px; font-weight: 600; background-color: {acento}; color: white; transition: 0.3s; border: none; }}
        .stButton>button:hover {{ background-color: {acento_hover}; border-color: {acento_hover}; }}
        div.stMetric {{ background-color: {superficie}; padding: 15px 20px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.15); border: 1px solid {borde}; }}
        section[data-testid="stSidebar"] {{ background-color: {superficie}; }}

        /* Forzar visibilidad de textos generales, markdown y elementos de la barra lateral */
        p, span, label, .stMarkdown, div[data-testid="stSidebar"] {{ color: {texto_principal} !important; }}

        /* Corregir contenedores, expanders y áreas de chat/IA para que no oculten las letras */
        div[data-testid="stExpander"], div[data-testid="stVerticalBlock"] {{ color: {texto_principal}; background-color: transparent; }}
        div[data-testid="stExpander"] summary {{ background-color: {superficie}; }}

        /* Forzar colores en campos de entrada, áreas de texto y selectores */
        input, textarea, select {{ background-color: {input_bg} !important; color: {texto_principal} !important; border-color: {borde} !important; }}
        div[data-baseweb="select"] > div {{ background-color: {input_bg} !important; color: {texto_principal} !important; border-color: {borde} !important; }}

        /* Corregir el texto dentro de los inputs de Streamlit */
        input::-webkit-input-placeholder {{ color: {texto_secundario} !important; }}
    </style>
    """

st.markdown(render_estilos(st.session_state.modo_oscuro), unsafe_allow_html=True)

if st.sidebar.toggle("🌙 Modo oscuro", key="modo_oscuro_toggle", value=st.session_state.get('modo_oscuro', False)):
    if not st.session_state.modo_oscuro:
        st.session_state.modo_oscuro = True
        st.rerun()
else:
    if st.session_state.modo_oscuro:
        st.session_state.modo_oscuro = False
        st.rerun()

# --- INICIALIZACIÓN DE FIREBASE ---
firebase_secrets = st.secrets.get("firebase", {}) if "firebase" in st.secrets else {}
FIREBASE_STORAGE_BUCKET = (
    firebase_secrets.get("storage_bucket")
    or os.environ.get("FIREBASE_STORAGE_BUCKET")
    or "proyecto-app-ffdb5.firebasestorage.app"
)

if not firebase_admin._apps:
    try:
        if "firebase" in st.secrets:
            cred_dict = dict(st.secrets["firebase"])
            cred_dict.pop("storage_bucket", None)
            cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        elif os.path.exists('firebase_key.json'):
            cred = credentials.Certificate('firebase_key.json')
        else:
            st.error("⚠️ No se encontraron credenciales de Firebase.")
            st.stop()

        # Usa un valor explícito de configuración si existe. Si no, deriva el
        # bucket moderno desde el project_id del service account. Para proyectos
        # con bucket legacy o personalizado, define firebase.storage_bucket.
        if not FIREBASE_STORAGE_BUCKET:
            project_id = getattr(cred, "project_id", None)
            if project_id:
                FIREBASE_STORAGE_BUCKET = f"{project_id}.firebasestorage.app"

        firebase_options = {}
        if FIREBASE_STORAGE_BUCKET:
            firebase_options["storageBucket"] = FIREBASE_STORAGE_BUCKET
        firebase_admin.initialize_app(cred, firebase_options)
    except Exception as e:
        st.error(f"⚠️ Error al conectar con Firebase: {e}")

db = firestore.client() if firebase_admin._apps else None
bucket = (
    storage.bucket(FIREBASE_STORAGE_BUCKET)
    if firebase_admin._apps and FIREBASE_STORAGE_BUCKET else None
)

# --- DATOS GLOBALES ---
EXCEL_FILE = "Proyecto_Financiero_Actualizado.xlsx"

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
    """ID estable de la 'nube' compartida: el correo del Propietario de la empresa.

    Antes esto devolvía siempre el correo del usuario logueado. Con el sistema de
    roles, un empleado inicia sesión con SU PROPIO correo, pero debe leer/escribir
    en los datos de la EMPRESA a la que pertenece, no en una nube propia y vacía.
    Por eso ahora se prioriza 'empresa_id' (el correo del Propietario, guardado en
    el documento del empleado al invitarlo) y solo se cae de vuelta al propio correo
    si 'empresa_id' no existe — que es exactamente el caso de un Propietario (para
    quien empresa_id == su propio correo) y el de cualquier cuenta creada ANTES de
    este sistema de roles, que sigue funcionando exactamente igual que hasta ahora.
    """
    if st.session_state.user_data:
        empresa_id = st.session_state.user_data.get('empresa_id')
        if empresa_id:
            return empresa_id.lower().strip()
        return st.session_state.user_data.get('email', '').lower().strip()
    return None


# =====================================================================================
# ROLES Y PERMISOS (multi-empresa)
# -------------------------------------------------------------------------------------
# Cuatro rangos, pensados para una empresa real, no solo un equipo de 4 personas:
#   - Propietario: acceso total, único que gestiona el equipo.
#   - Gerente Financiero: todo excepto gestionar el equipo.
#   - Analista: registra movimientos, pero NO aprueba/rechaza pendientes.
#   - Auditor: solo lectura (auditoría, reportes).
#
# Compatibilidad hacia atrás: una cuenta creada antes de este sistema no tiene
# campo 'rol' en su documento. tiene_permiso() la trata como Propietario para no
# quitarle acceso a nadie que ya estaba usando la app.
# =====================================================================================

def get_perm_labels():
    return {
        "ver_finanzas": "Ver ingresos, gastos, balance, paneles y reportes",
        "registrar": "Registrar ingresos y gastos",
        "subir_comprobantes": "Subir y procesar comprobantes con IA",
        "aprobar_pendientes": "Aprobar o rechazar ingresos y gastos pendientes",
        "configurar_presupuesto": "Configurar presupuestos y períodos",
        "ver_auditoria": "Consultar la auditoría",
        "gestionar_archivos": "Administrar documentos de la empresa",
        "eliminar_registros": "Eliminar ingresos y gastos ya registrados",
        "generar_recibos": "Generar comprobantes y códigos QR",
        "gestionar_equipo": "Invitar empleados y administrar rangos",
    }

PERMISOS_DISPONIBLES = get_perm_labels()
TODOS_LOS_PERMISOS = set(PERMISOS_DISPONIBLES)
ROLES_DISPONIBLES = ["Propietario", "Gerente Financiero", "Analista", "Auditor"]

PERMISOS_POR_ROL = {
    "Propietario": set(TODOS_LOS_PERMISOS),
    "Gerente Financiero": {
        "ver_finanzas", "registrar", "subir_comprobantes", "aprobar_pendientes",
        "configurar_presupuesto", "ver_auditoria", "gestionar_archivos",
        "eliminar_registros", "generar_recibos",
    },
    "Analista": {
        "ver_finanzas", "registrar", "subir_comprobantes", "gestionar_archivos",
    },
    "Auditor": {"ver_finanzas", "ver_auditoria"},
}


def cargar_roles_personalizados():
    institucion_id = get_institucion_id()
    if not db or not institucion_id:
        return {}
    try:
        doc = db.collection("usuarios").document(institucion_id).get()
        datos = doc.to_dict() or {}
        roles = datos.get("roles_personalizados", {})
        return roles if isinstance(roles, dict) else {}
    except Exception:
        return {}


def obtener_roles_asignables():
    roles_custom = cargar_roles_personalizados()
    return [r for r in ROLES_DISPONIBLES if r != "Propietario"] + sorted(roles_custom.keys())


def permisos_del_usuario():
    user_data = st.session_state.get("user_data") or {}
    rol_actual = user_data.get("rol", "Propietario")
    if rol_actual == "Propietario":
        return set(TODOS_LOS_PERMISOS)
    roles_custom = cargar_roles_personalizados()
    if rol_actual in roles_custom:
        permisos = roles_custom[rol_actual].get("permisos", [])
        return set(permisos) & TODOS_LOS_PERMISOS
    return set(PERMISOS_POR_ROL.get(rol_actual, set()))


def tiene_permiso(accion: str) -> bool:
    return accion in permisos_del_usuario()


def validar_membresia_actual():
    user_data = st.session_state.get("user_data") or {}
    if not user_data:
        return False
    email = (user_data.get("email") or "").lower().strip()
    empresa_id = (user_data.get("empresa_id") or email).lower().strip()
    if not email or not empresa_id:
        return False
    if empresa_id == email:
        return True
    if not db:
        return False
    try:
        miembro = db.collection("usuarios").document(empresa_id).collection("empleados").document(email).get()
        if not miembro.exists:
            return False
        datos = miembro.to_dict() or {}
        if datos.get("estado") != "Activo":
            return False
        user_data["rol"] = datos.get("rol", "Analista")
        return True
    except Exception:
        return False


def obtener_lista_responsables():
    user_data = st.session_state.get("user_data") or {}
    institucion_id = get_institucion_id()
    fallback = [user_data.get("institucion") or user_data.get("email") or "Usuario actual"]
    if not db or not institucion_id:
        return fallback
    try:
        nombres = []
        doc_empresa = db.collection("usuarios").document(institucion_id).get()
        datos_empresa = doc_empresa.to_dict() or {}
        nombres.append(datos_empresa.get("responsable") or datos_empresa.get("email") or institucion_id)
        empleados_ref = (
            db.collection("usuarios").document(institucion_id).collection("empleados")
            .where("estado", "==", "Activo").stream()
        )
        for emp in empleados_ref:
            emp_data = emp.to_dict() or {}
            nombre = emp_data.get("nombre") or emp_data.get("email")
            if nombre and nombre not in nombres:
                nombres.append(nombre)
        return nombres or fallback
    except Exception:
        return fallback


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
    if not db:
        return False
    institucion_id = get_institucion_id()
    if not institucion_id:
        return False
    try:
        db.collection("usuarios").document(institucion_id).collection(coleccion).document(datos["ID"]).set(datos)
        return True
    except Exception:
        return False


def eliminar_registro_nube(coleccion, doc_id):
    if not db:
        return False
    institucion_id = get_institucion_id()
    if not institucion_id:
        return False
    try:
        db.collection("usuarios").document(institucion_id).collection(coleccion).document(doc_id).delete()
        return True
    except Exception:
        return False


def generar_id(prefijo):
    return f"{prefijo}-{uuid.uuid4().hex.upper()}"


def prefijo_storage_empresa():
    institucion_id = get_institucion_id()
    if not institucion_id:
        return None
    hash_empresa = hashlib.sha256(institucion_id.encode("utf-8")).hexdigest()[:32]
    return f"empresas/{hash_empresa}"


def subir_evidencia_comprobante(datos_bytes, nombre_archivo, mime_type):
    if not bucket:
        raise RuntimeError(
            "Cloud Storage no está configurado. Define firebase.storage_bucket con el nombre exacto del bucket que aparece en Firebase Console > Storage."
        )
    prefijo = prefijo_storage_empresa()
    if not prefijo:
        raise RuntimeError("No se pudo identificar la empresa para guardar el comprobante.")
    extension = os.path.splitext(nombre_archivo)[1].lower()
    if not re.fullmatch(r"\.[a-z0-9]{1,8}", extension):
        extension = ".bin"
    ruta = f"{prefijo}/comprobantes/{uuid.uuid4().hex}{extension}"
    blob = bucket.blob(ruta)
    blob.upload_from_string(datos_bytes, content_type=mime_type or "application/octet-stream")
    return ruta


def descargar_evidencia_comprobante(ruta):
    if not bucket or not ruta:
        return None
    prefijo = prefijo_storage_empresa()
    if not prefijo or not ruta.startswith(prefijo + "/"):
        return None
    return bucket.blob(ruta).download_as_bytes()


def guardar_pendiente_movimiento(tipo, registro, motivos=None, origen="Manual", evidencia_ruta=None, evidencia_nombre=None, evidencia_mime=None):
    if not db:
        return None
    institucion_id = get_institucion_id()
    if not institucion_id:
        return None
    tipo_limpio = "Ingreso" if str(tipo).lower().startswith("ing") else "Gasto"
    pendiente_id = generar_id("PEND-ING" if tipo_limpio == "Ingreso" else "PEND-GAS")
    pendiente = dict(registro)
    pendiente.update({
        "ID": pendiente_id,
        "Tipo": tipo_limpio,
        "Motivos": motivos or ["Enviado para revisión humana."],
        "Estado": "Pendiente",
        "Origen": origen,
        "creado_por": (st.session_state.get("user_data") or {}).get("email", ""),
        "creado_en": dt_module.datetime.now(dt_module.timezone.utc).isoformat(),
    })
    if evidencia_ruta:
        pendiente["EvidenciaStoragePath"] = evidencia_ruta
        pendiente["EvidenciaNombre"] = evidencia_nombre or "comprobante"
        pendiente["EvidenciaMime"] = evidencia_mime or "application/octet-stream"
    coleccion = "ingresos_pendientes" if tipo_limpio == "Ingreso" else "gastos_pendientes"
    try:
        db.collection("usuarios").document(institucion_id).collection(coleccion).document(pendiente_id).set(pendiente)
        return pendiente_id
    except Exception:
        if evidencia_ruta and bucket:
            try:
                bucket.blob(evidencia_ruta).delete()
            except Exception:
                pass
        return None


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


def cargar_configuracion_presupuesto():
    hoy = dt_module.date.today()
    default = {
        "presupuesto_tope": 500000.0,
        "fecha_inicio_presupuesto": hoy.isoformat(),
        "fecha_fin_presupuesto": (hoy + dt_module.timedelta(days=30)).isoformat(),
    }
    if not db:
        return default
    institucion_id = get_institucion_id()
    if not institucion_id:
        return default
    try:
        datos = db.collection("usuarios").document(institucion_id).get().to_dict() or {}
        valor = float(datos.get("presupuesto_tope", default["presupuesto_tope"]))
        inicio = datos.get("fecha_inicio_presupuesto", default["fecha_inicio_presupuesto"])
        fin = datos.get("fecha_fin_presupuesto", default["fecha_fin_presupuesto"])
        dt_module.date.fromisoformat(inicio)
        dt_module.date.fromisoformat(fin)
        return {
            "presupuesto_tope": max(0.0, valor),
            "fecha_inicio_presupuesto": inicio,
            "fecha_fin_presupuesto": fin,
        }
    except Exception:
        return default


def guardar_configuracion_presupuesto(presupuesto_tope, periodo):
    if not db or not get_institucion_id():
        return False
    if not isinstance(periodo, tuple) or len(periodo) != 2:
        return False
    inicio, fin = periodo
    datos = {
        "presupuesto_tope": float(presupuesto_tope),
        "fecha_inicio_presupuesto": inicio.isoformat(),
        "fecha_fin_presupuesto": fin.isoformat(),
    }
    try:
        db.collection("usuarios").document(get_institucion_id()).set(datos, merge=True)
        return True
    except Exception:
        return False


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
# OCR DE COMPROBANTES DE INGRESO + RELACIÓN INGRESOS-PRESUPUESTO
# -------------------------------------------------------------------------------------
# A diferencia de los gastos, un ingreso NO se compara contra el presupuesto como
# regla de bloqueo — el presupuesto es un tope de gasto, no de ingreso, así que un
# ingreso alto nunca debería "quedar pendiente" solo por ser alto frente al
# presupuesto. Lo que sí tiene sentido es duplicidad (mismo depósito registrado dos
# veces) y monto atípico frente al historial. La relación con el presupuesto se
# muestra aparte, como contexto informativo (calcular_relacion_presupuesto), en la
# sección 2 y junto al resultado del OCR — no como una condición que bloquea nada.
# =====================================================================================

def extraer_datos_ingreso_gemini(archivo_bytes, mime_type, api_key):
    """Envía la imagen/PDF de un comprobante de ingreso (transferencia, consignación,
    recibo de pago recibido) a Gemini y devuelve un diccionario con los campos
    extraídos, o None si la extracción falla."""
    try:
        from google import genai
        from google.genai import types
        import json as json_lib

        client = genai.Client(api_key=api_key)
        prompt = (
            "Eres un asistente que extrae datos de comprobantes de ingreso escolares "
            "(transferencias, consignaciones, recibos de pago recibido, donaciones). "
            "Analiza la imagen o documento adjunto y responde ÚNICAMENTE con un objeto JSON "
            "(sin texto adicional, sin explicaciones, sin backticks de markdown), con "
            "exactamente estas claves:\n"
            '{"concepto": "descripción breve de qué es el ingreso (ej. pago de rifa, donación, matrícula)", '
            '"valor": numero_sin_simbolos_ni_comas_ni_texto, '
            '"fecha": "YYYY-MM-DD si es visible en el documento, o null si no se ve", '
            '"observaciones": "cualquier dato adicional visible (referencia, quién paga), o null"}'
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
        st.error(f"No se pudieron extraer los datos del comprobante automáticamente: {e}")
        return None


def extraer_datos_comprobante_gemini(archivo_bytes, mime_type, nombre_archivo, api_key):
    """Clasifica un comprobante como ingreso o gasto y extrae sus datos editables."""
    client = None
    archivo_remoto = None
    ruta_temporal = None
    try:
        from google import genai
        from google.genai import types
        import json as json_lib

        extension = os.path.splitext(nombre_archivo)[1] or ".bin"
        if mime_type == "application/pdf" and len(archivo_bytes) > 50 * 1024 * 1024:
            raise ValueError("Gemini acepta PDF de hasta 50 MB por comprobante.")
        if len(archivo_bytes) > 200 * 1024 * 1024:
            raise ValueError("El límite de esta aplicación es 200 MB por archivo.")

        client = genai.Client(api_key=api_key)
        prompt = (
            "Analiza el comprobante adjunto y clasifícalo como un movimiento financiero de una empresa. "
            "Puede ser un recibo de compra o servicio (Gasto), o evidencia de dinero recibido (Ingreso), "
            "incluyendo comprobantes de Nequi, bancos, transferencias, consignaciones, supermercados, "
            "facturas, recibos y pagos. No inventes datos: si no son legibles usa null. "
            "Interpreta el valor en pesos colombianos y la fecha del documento. "
            "Para gastos elige exactamente una categoría: Logística, Publicidad, Alimentación, Varios. "
            "Devuelve solo JSON válido con exactamente estas claves: "
            '{"tipo":"Ingreso, Gasto o Indeterminado","concepto":"texto breve","valor":numero_o_null,'
            '"fecha":"YYYY-MM-DD o null","categoria_sugerida":"una categoría exacta o null",'
            '"observaciones":"referencia, pagador o comercio visible, o null",'
            '"numero_factura":"folio visible o null"}.'
        )
        if len(archivo_bytes) <= 20 * 1024 * 1024:
            parte_archivo = types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)
            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=[parte_archivo, prompt],
            )
        else:
            with tempfile.NamedTemporaryFile(suffix=extension, delete=False) as temporal:
                temporal.write(archivo_bytes)
                ruta_temporal = temporal.name
            archivo_remoto = client.files.upload(file=ruta_temporal)
            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=[archivo_remoto, prompt],
            )

        texto = (response.text or "").strip()
        fence = chr(96) * 3
        if texto.startswith(fence + "json"):
            texto = texto[len(fence + "json"):].strip()
        elif texto.startswith(fence):
            texto = texto[3:].strip()
        if texto.endswith(fence):
            texto = texto[:-3].strip()
        datos = json_lib.loads(texto)
        tipo = str(datos.get("tipo") or "Indeterminado").strip().lower()
        if tipo in {"ingreso", "abono", "depósito", "deposito"}:
            datos["tipo"] = "Ingreso"
        elif tipo in {"gasto", "compra", "egreso", "pago"}:
            datos["tipo"] = "Gasto"
        else:
            datos["tipo"] = "Indeterminado"
        categoria = datos.get("categoria_sugerida")
        datos["categoria_sugerida"] = categoria if categoria in CATEGORIAS_GASTO else "Varios"
        return datos
    except Exception as e:
        st.error(f"No se pudo analizar {nombre_archivo}: {e}")
        return None
    finally:
        if client and archivo_remoto:
            try:
                client.files.delete(name=archivo_remoto.name)
            except Exception:
                pass
        if ruta_temporal and os.path.exists(ruta_temporal):
            try:
                os.remove(ruta_temporal)
            except Exception:
                pass


def detectar_anomalias_ingreso(nuevo_valor, nueva_fecha, nuevo_concepto, ingresos_existentes_df):
    """Reglas para ingresos: solo duplicidad y monto atípico frente al historial.
    Intencionalmente NO incluye ninguna regla de presupuesto (ver nota arriba)."""
    import difflib

    motivos = []
    if ingresos_existentes_df is None or ingresos_existentes_df.empty:
        return motivos

    df_temp = ingresos_existentes_df.copy()
    df_temp["Valor"] = pd.to_numeric(df_temp["Valor"], errors='coerce').fillna(0)
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
            motivos.append(f"Posible ingreso duplicado: se parece a '{fila['Concepto']}' del {fila['Fecha']}, por el mismo valor.")
            break

    if len(df_temp) >= 3:
        promedio = df_temp["Valor"].mean()
        if promedio > 0 and nuevo_valor > promedio * 2.5:
            motivos.append(f"Monto inusualmente alto frente al promedio histórico de ingresos (${promedio:,.0f}).")

    return motivos


def calcular_relacion_presupuesto(total_ingresos, total_gastos, presupuesto_tope):
    """Muestra cobertura de ingresos y ejecución de gastos contra el presupuesto."""
    total_ingresos = float(total_ingresos or 0)
    total_gastos = float(total_gastos or 0)
    presupuesto_tope = float(presupuesto_tope or 0)
    saldo = total_ingresos - total_gastos
    if presupuesto_tope <= 0:
        return {
            "nivel": "sin_presupuesto",
            "mensaje": "Configura un presupuesto para mostrar cuánto cubren los ingresos y cuánto se ha ejecutado.",
            "saldo": saldo,
            "porcentaje_cubierto": 0.0,
        }
    porcentaje_gastado = (total_gastos / presupuesto_tope) * 100
    porcentaje_cubierto = (total_ingresos / presupuesto_tope) * 100
    faltante_ingresos = max(0.0, presupuesto_tope - total_ingresos)
    disponible_gasto = presupuesto_tope - total_gastos
    if porcentaje_gastado >= 100:
        nivel = "critico"
        mensaje = (
            f"Los gastos superan el límite de {presupuesto_tope:,.0f} COP. "
            f"Los ingresos registrados cubren el {porcentaje_cubierto:.0f}% del presupuesto."
        )
    elif porcentaje_gastado >= 80:
        nivel = "alerta"
        mensaje = (
            f"Se ha usado el {porcentaje_gastado:.0f}% del límite de gastos; "
            f"quedan {disponible_gasto:,.0f} COP. Los ingresos cubren el {porcentaje_cubierto:.0f}%."
        )
    else:
        mensaje = (
            f"Los ingresos registrados cubren el {porcentaje_cubierto:.0f}% del presupuesto de gastos; "
            f"quedan {faltante_ingresos:,.0f} COP para que los ingresos registrados igualen el presupuesto. "
            f"Se ha usado el {porcentaje_gastado:.0f}% del límite."
        )
        nivel = "ok"
    return {
        "nivel": nivel,
        "mensaje": mensaje,
        "saldo": saldo,
        "porcentaje_cubierto": porcentaje_cubierto,
        "porcentaje_gastado": porcentaje_gastado,
        "faltante_ingresos": faltante_ingresos,
        "disponible_gasto": disponible_gasto,
    }


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
            # Cuenta nueva vía Google: mismo chequeo de invitación que el registro tradicional.
            invitacion_doc = db.collection('empleados_index').document(email_google).get()
            if invitacion_doc.exists:
                invitacion = invitacion_doc.to_dict()
                user_data = {
                    'institucion': nombre_google,
                    'email': email_google,
                    'password': hash_password(''.join(random.choices(string.ascii_letters + string.digits, k=24))),
                    'rol': invitacion.get('rol', 'Analista'),
                    'empresa_id': invitacion.get('empresa_id'),
                    'login_provider': 'google',
                    'fecha_creacion': dt_module.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                }
                db.collection('usuarios').document(invitacion.get('empresa_id')).collection('empleados').document(email_google).update({'nombre': nombre_google, 'estado': 'Activo'})
            else:
                user_data = {
                    'institucion': nombre_google,
                    'email': email_google,
                    'password': hash_password(''.join(random.choices(string.ascii_letters + string.digits, k=24))),
                    'rol': 'Propietario',
                    'empresa_id': email_google,
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
def generar_imagen_recibo(rec_id, fecha, tot_ing, tot_gas, saldo, qr_img_pil, nombre_empresa):
    img_w, img_h = 650, 940
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
    draw.text((30, 25), "COMPROBANTE FINANCIERO", fill="#FFFFFF", font=font_title)
    draw.text((30, 60), "Resumen de movimientos de la empresa", fill="#93C5FD", font=font_regular)
    draw.rectangle([(30, 130), (img_w - 30, img_h - 35)], outline="#E2E8F0", width=2, fill="#F8FAFC")
    draw.text((55, 160), "ID de comprobante:", fill="#64748B", font=font_small)
    draw.text((200, 158), rec_id, fill="#1E293B", font=font_bold)
    draw.text((55, 190), "Fecha de emisión:", fill="#64748B", font=font_small)
    draw.text((200, 188), fecha, fill="#1E293B", font=font_bold)
    draw.text((55, 220), "Empresa:", fill="#64748B", font=font_small)

    nombre_visual = str(nombre_empresa or "Empresa")
    while nombre_visual and draw.textbbox((0, 0), nombre_visual, font=font_bold)[2] > img_w - 255:
        nombre_visual = nombre_visual[:-2] + "…"
    draw.text((200, 218), nombre_visual, fill="#1E293B", font=font_bold)

    draw.line([(55, 255), (img_w - 55, 255)], fill="#CBD5E1", width=1)
    draw.text((55, 280), "RESUMEN DE MOVIMIENTOS", fill="#1E3A8A", font=font_bold)
    draw.text((55, 320), "(+) Total ingresos:", fill="#334155", font=font_regular)
    draw.text((400, 320), f"{chr(36)}{tot_ing:,.0f} COP", fill="#059669", font=font_bold)
    draw.text((55, 360), "(-) Total gastos:", fill="#334155", font=font_regular)
    draw.text((400, 360), f"{chr(36)}{tot_gas:,.0f} COP", fill="#DC2626", font=font_bold)
    draw.line([(55, 400), (img_w - 55, 400)], fill="#CBD5E1", width=1)
    draw.text((55, 420), "BALANCE NETO:", fill="#1E3A8A", font=font_bold)
    color_saldo = "#059669" if saldo >= 0 else "#DC2626"
    draw.text((370, 415), f"{chr(36)}{saldo:,.0f} COP", fill=color_saldo, font=font_title)
    draw.text((55, 465), "ESTADO: SUPERÁVIT" if saldo >= 0 else "ESTADO: DÉFICIT", fill=color_saldo, font=font_small)

    qr_resized = qr_img_pil.resize((250, 250))
    base_img.paste(qr_resized, (200, 490))
    draw.text((int(img_w / 2) - 135, 760), "Escanea el QR para descargar este PNG", fill="#64748B", font=font_small)
    draw.text((int(img_w / 2) - 132, 782), "El enlace se puede revocar desde Firebase Storage", fill="#64748B", font=font_small)
    draw.text((int(img_w / 2) - 120, 875), "Sistema de Gestión Financiera", fill="#94A3B8", font=font_small)

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
    st.markdown('<p class="sub-header" style="text-align: center;">Administración financiera empresarial</p>', unsafe_allow_html=True)

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
            inst_name = st.text_input("Nombre de la Institución / Persona", help="Si tu correo fue invitado como empleado de una empresa, escribe aquí tu propio nombre.")
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
                        # ¿Este correo fue invitado como empleado de alguna empresa?
                        # Si sí, se vincula a esa empresa en vez de crear una nube nueva.
                        invitacion_doc = db.collection('empleados_index').document(email_clean).get()
                        if invitacion_doc.exists:
                            invitacion = invitacion_doc.to_dict()
                            nuevo_usuario = {
                                'institucion': inst_name,
                                'email': email_clean,
                                'password': hash_password(pass_reg),
                                'rol': invitacion.get('rol', 'Analista'),
                                'empresa_id': invitacion.get('empresa_id'),
                                'fecha_creacion': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            }
                            db.collection('usuarios').document(email_clean).set(nuevo_usuario)
                            db.collection('usuarios').document(invitacion.get('empresa_id')).collection('empleados').document(email_clean).update({'nombre': inst_name, 'estado': 'Activo'})
                            st.success(f"✅ ¡Cuenta creada! Quedaste vinculado a {invitacion.get('nombre_empresa', 'tu empresa')} como {invitacion.get('rol', 'Analista')}.")
                        else:
                            nuevo_usuario = {
                                'institucion': inst_name,
                                'email': email_clean,
                                'password': hash_password(pass_reg),
                                'rol': 'Propietario',
                                'empresa_id': email_clean,
                                'fecha_creacion': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
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
                        msg['Subject'] = "Recuperación de contraseña - Gestión Financiera Empresarial"

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

if st.session_state.get("logged_in") and not validar_membresia_actual():
    st.session_state.logged_in = False
    st.session_state.user_data = None
    st.session_state.pop("ingresos_df", None)
    st.session_state.pop("gastos_df", None)
    st.error("Tu acceso a esta empresa fue revocado o no está activo. Contacta al administrador.")
    st.stop()

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
st.sidebar.markdown("⚙️ **Presupuesto de la empresa**")
cfg_presupuesto = cargar_configuracion_presupuesto()
hoy = dt_module.date.today()
try:
    inicio_default = dt_module.date.fromisoformat(cfg_presupuesto["fecha_inicio_presupuesto"])
    fin_default = dt_module.date.fromisoformat(cfg_presupuesto["fecha_fin_presupuesto"])
except (KeyError, TypeError, ValueError):
    inicio_default, fin_default = hoy, hoy + dt_module.timedelta(days=30)

puede_configurar_presupuesto = tiene_permiso("configurar_presupuesto")
clave_tenant = hashlib.sha256((get_institucion_id() or "sin-empresa").encode("utf-8")).hexdigest()[:12]
with st.sidebar.form("form_configuracion_presupuesto"):
    periodo_presupuesto = st.date_input(
        "Período de ejecución",
        value=(inicio_default, fin_default),
        key=f"periodo_presupuesto_{clave_tenant}",
        disabled=not puede_configurar_presupuesto,
    )
    presupuesto_tope = st.number_input(
        "Límite de gastos (COP)",
        min_value=0.0,
        value=float(cfg_presupuesto["presupuesto_tope"]),
        step=50000.0,
        key=f"presupuesto_tope_{clave_tenant}",
        disabled=not puede_configurar_presupuesto,
    )
    guardar_presupuesto_click = st.form_submit_button(
        "Guardar presupuesto",
        disabled=not puede_configurar_presupuesto,
        use_container_width=True,
    )
if guardar_presupuesto_click:
    if guardar_configuracion_presupuesto(presupuesto_tope, periodo_presupuesto):
        registrar_auditoria("Actualizó presupuesto general", f"{presupuesto_tope:,.0f} COP")
        st.sidebar.success("Presupuesto guardado para esta empresa.")
    else:
        st.sidebar.error("No se pudo guardar el presupuesto.")
if not puede_configurar_presupuesto:
    st.sidebar.caption("Solo lectura: tu rango no permite modificarlo.")

if isinstance(periodo_presupuesto, tuple) and len(periodo_presupuesto) == 2:
    fecha_fin = periodo_presupuesto[1]
    if hoy > fecha_fin and not st.session_state.omitir_alerta_presupuesto:
        with st.sidebar.container():
            st.warning("El período del presupuesto terminó. Actualiza las fechas cuando corresponda.")
            if puede_configurar_presupuesto and st.button("Omitir por ahora"):
                st.session_state.omitir_alerta_presupuesto = True
                st.rerun()
    elif hoy <= fecha_fin:
        st.session_state.omitir_alerta_presupuesto = False

st.sidebar.markdown("---")
opciones_menu = [
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
]
opciones_menu.append("11. Gestión de Equipo")
permisos_menu = {
    "1. Inicio": None,
    "2. Registro de Ingresos": "ver_finanzas",
    "3. Registro de Gastos": "ver_finanzas",
    "4. Balance Financiero": "ver_finanzas",
    "5. Dashboard y Gráficos": "ver_finanzas",
    "6. Anexo de Recibos & QR": "generar_recibos",
    "7. Gestión de Archivos": "gestionar_archivos",
    "8. Reporte Final": "ver_finanzas",
    "9. Auditoría del Sistema": "ver_auditoria",
    "10. Facturas (OCR) y Anomalías": ("subir_comprobantes", "aprobar_pendientes"),
    "11. Gestión de Equipo": "gestionar_equipo",
}
opciones_menu = [
    opcion for opcion in opciones_menu
    if permisos_menu.get(opcion) is None
    or (any(tiene_permiso(p) for p in permisos_menu[opcion])
        if isinstance(permisos_menu[opcion], tuple)
        else tiene_permiso(permisos_menu[opcion]))
]
if not opciones_menu:
    opciones_menu = ["1. Inicio"]
menu = st.sidebar.selectbox("📌 Selecciona una sección:", opciones_menu)
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
        st.write("Control transparente y automatizado de ingresos, gastos, presupuestos, documentos y aprobaciones, con datos separados por empresa.")
    with col2:
        st.success("✅ **Estado del Sistema:** Operativo y Guardado en Nube.")

    st.markdown("---")
    st.markdown("### 👥 Equipo")
    equipo_actual = obtener_lista_responsables()
    equipo_data = [{"N.°": i + 1, "Nombre": nombre} for i, nombre in enumerate(equipo_actual)]
    st.dataframe(pd.DataFrame(equipo_data), use_container_width=True, hide_index=True)
    if tiene_permiso("gestionar_equipo"):
        st.caption("Gestiona rangos e invita nuevas personas desde '11. Gestión de Equipo' en el menú lateral.")

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
                resp_ing = st.selectbox("Responsable", obtener_lista_responsables())
                val_ing = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
            obs_ing = st.text_area("Observaciones (Opcional)")

            if st.form_submit_button("Guardar Ingreso", disabled=not tiene_permiso("registrar")):
                if not tiene_permiso("registrar"):
                    st.error("Tu rango no permite registrar ingresos.")
                elif con_ing.strip() == "":
                    st.error("El concepto no puede estar vacío.")
                else:
                    nuevo_reg = {
                        "ID": generar_id("ING"),
                        "Fecha": f_ing.strftime("%Y-%m-%d"),
                        "Concepto": con_ing.strip(),
                        "Valor": float(val_ing),
                        "Responsable": resp_ing,
                        "Observaciones": obs_ing,
                    }
                    if tiene_permiso("aprobar_pendientes"):
                        if guardar_registro_nube("ingresos", nuevo_reg):
                            st.session_state.ingresos_df = pd.concat(
                                [st.session_state.ingresos_df, pd.DataFrame([nuevo_reg])],
                                ignore_index=True,
                            )
                            registrar_auditoria("Registró Ingreso", f"{con_ing} - {val_ing:,.0f} COP")
                            st.success("Ingreso registrado.")
                            st.rerun()
                        else:
                            st.error("No se pudo guardar el ingreso en la nube.")
                    else:
                        pendiente_id = guardar_pendiente_movimiento(
                            "Ingreso", nuevo_reg,
                            ["Registro manual enviado a revisión por el rango del usuario."],
                            origen="Registro manual",
                        )
                        if pendiente_id:
                            registrar_auditoria("Envió Ingreso a revisión", f"{con_ing} - {val_ing:,.0f} COP")
                            st.success("Ingreso enviado a revisión; se contará en el presupuesto al ser aprobado.")
                            st.rerun()
                        else:
                            st.error("No se pudo crear el ingreso pendiente.")

    # Solo las transacciones aprobadas participan en esta comparación.
    tot_ing_actual = pd.to_numeric(
        st.session_state.ingresos_df.get("Valor", pd.Series(dtype=float)), errors="coerce"
    ).fillna(0).sum()
    tot_gas_actual = pd.to_numeric(
        st.session_state.gastos_df.get("Valor", pd.Series(dtype=float)), errors="coerce"
    ).fillna(0).sum()
    relacion = calcular_relacion_presupuesto(tot_ing_actual, tot_gas_actual, presupuesto_tope)
    st.markdown("### Relación de ingresos y gastos con el presupuesto")
    if relacion["nivel"] == "critico":
        st.error(relacion["mensaje"])
    elif relacion["nivel"] == "alerta":
        st.warning(relacion["mensaje"])
    elif relacion["nivel"] == "ok":
        st.info(relacion["mensaje"])
    else:
        st.caption(relacion["mensaje"])
    col_ing_rel, col_gas_rel, col_cobertura_rel = st.columns(3)
    col_ing_rel.metric("Ingresos aprobados", f"{tot_ing_actual:,.0f} COP")
    col_gas_rel.metric("Gastos aprobados", f"{tot_gas_actual:,.0f} COP")
    col_cobertura_rel.metric("Cobertura del presupuesto", f"{relacion.get('porcentaje_cubierto', 0):.0f}%")

    if not st.session_state.ingresos_df.empty:
        st.dataframe(st.session_state.ingresos_df.drop(columns=['ID']), use_container_width=True)
        st.metric("💵 TOTAL INGRESOS", f"${st.session_state.ingresos_df['Valor'].astype(float).sum():,.0f} COP")

        if tiene_permiso("eliminar_registros"):
            st.markdown("### Eliminar ingreso")
            opciones = [f"{row['Concepto']} - {chr(36)}{row['Valor']:,.0f}" for _, row in st.session_state.ingresos_df.iterrows()]
            seleccion = st.selectbox("Selecciona para eliminar:", opciones, key="sel_ingreso_eliminar")
            if st.button("Eliminar ingreso", key="btn_eliminar_ingreso"):
                idx = opciones.index(seleccion)
                fila = st.session_state.ingresos_df.iloc[idx]
                if eliminar_registro_nube("ingresos", fila["ID"]):
                    registrar_auditoria("Eliminó Ingreso", f"{fila['Concepto']} - {float(fila['Valor']):,.0f} COP")
                    st.session_state.ingresos_df = st.session_state.ingresos_df.drop(idx).reset_index(drop=True)
                    st.success("Ingreso eliminado.")
                    st.rerun()
                else:
                    st.error("No se pudo eliminar el ingreso.")
        else:
            st.caption("Tu rango no permite eliminar ingresos.")
    
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
                resp_gas = st.selectbox("Responsable", obtener_lista_responsables())

            if st.form_submit_button("Guardar Gasto", disabled=not tiene_permiso("registrar")):
                if not tiene_permiso("registrar"):
                    st.error("Tu rango no permite registrar gastos.")
                elif con_gas.strip() == "":
                    st.error("El concepto no puede estar vacío.")
                else:
                    nuevo_reg = {
                        "ID": generar_id("GAS"),
                        "Fecha": f_gas.strftime("%Y-%m-%d"),
                        "Concepto": con_gas.strip(),
                        "Categoría": cat_gas,
                        "Valor": float(val_gas),
                        "Responsable": resp_gas,
                    }
                    if tiene_permiso("aprobar_pendientes"):
                        if guardar_registro_nube("gastos", nuevo_reg):
                            st.session_state.gastos_df = pd.concat(
                                [st.session_state.gastos_df, pd.DataFrame([nuevo_reg])],
                                ignore_index=True,
                            )
                            registrar_auditoria("Registró Gasto", f"{con_gas} - {val_gas:,.0f} COP")
                            st.success("Gasto registrado.")
                            st.rerun()
                        else:
                            st.error("No se pudo guardar el gasto en la nube.")
                    else:
                        pendiente_id = guardar_pendiente_movimiento(
                            "Gasto", nuevo_reg,
                            ["Registro manual enviado a revisión por el rango del usuario."],
                            origen="Registro manual",
                        )
                        if pendiente_id:
                            registrar_auditoria("Envió Gasto a revisión", f"{con_gas} - {val_gas:,.0f} COP")
                            st.success("Gasto enviado a revisión; aún no afecta el presupuesto.")
                            st.rerun()
                        else:
                            st.error("No se pudo crear el gasto pendiente.")
    
    if not st.session_state.gastos_df.empty:
        st.dataframe(st.session_state.gastos_df.drop(columns=['ID']), use_container_width=True)
        st.metric("💸 TOTAL GASTOS", f"${current_total_gastos:,.0f} COP")

        if tiene_permiso("eliminar_registros"):
            st.markdown("### Eliminar gasto")
            opciones = [f"{row['Concepto']} - {chr(36)}{row['Valor']:,.0f}" for _, row in st.session_state.gastos_df.iterrows()]
            seleccion = st.selectbox("Selecciona para eliminar:", opciones, key="sel_gasto_eliminar")
            if st.button("Eliminar gasto", key="btn_eliminar_gasto"):
                idx = opciones.index(seleccion)
                fila = st.session_state.gastos_df.iloc[idx]
                if eliminar_registro_nube("gastos", fila["ID"]):
                    registrar_auditoria("Eliminó Gasto", f"{fila['Concepto']} - {float(fila['Valor']):,.0f} COP")
                    st.session_state.gastos_df = st.session_state.gastos_df.drop(idx).reset_index(drop=True)
                    st.success("Gasto eliminado.")
                    st.rerun()
                else:
                    st.error("No se pudo eliminar el gasto.")
        else:
            st.caption("Tu rango no permite eliminar gastos.")
    
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
        if not tiene_permiso("configurar_presupuesto"):
            st.caption("🔒 Solo lectura — tu rango no permite modificar el presupuesto.")
            for categoria in CATEGORIAS_GASTO:
                st.write(f"**{categoria}:** ${presupuestos_actuales.get(categoria, 0.0):,.0f}")
        else:
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

    if not st.session_state.gastos_df.empty:
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
    st.markdown('<p class="main-header">Comprobante financiero y descarga QR</p>', unsafe_allow_html=True)
    st.markdown("---")
    st.info("El QR abre la descarga PNG del comprobante. El enlace usa un token revocable de Firebase Storage.")
    if st.button("Generar y guardar comprobante PNG", disabled=not tiene_permiso("generar_recibos")):
        try:
            if not bucket:
                raise RuntimeError("Firebase Storage no está configurado.")
            tot_ing = pd.to_numeric(st.session_state.ingresos_df.get("Valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
            tot_gas = pd.to_numeric(st.session_state.gastos_df.get("Valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
            saldo = tot_ing - tot_gas
            rec_id = generar_id("GEN")
            fecha_actual = dt_module.date.today().isoformat()
            empresa_doc = db.collection("usuarios").document(get_institucion_id()).get().to_dict() or {}
            nombre_empresa = empresa_doc.get("institucion") or empresa_doc.get("nombre_empresa") or "Empresa"
            prefijo = prefijo_storage_empresa()
            if not prefijo:
                raise RuntimeError("No se pudo identificar la nube de la empresa.")
            ruta_storage = f"{prefijo}/recibos/{rec_id}.png"
            token_descarga = uuid.uuid4().hex
            blob_recibo = bucket.blob(ruta_storage)
            blob_recibo.metadata = {"firebaseStorageDownloadTokens": token_descarga}
            blob_recibo.content_type = "image/png"
            blob_recibo.content_disposition = f'attachment; filename="{rec_id}.png"'
            url_descarga = (
                "https://firebasestorage.googleapis.com/v0/b/"
                + quote(bucket.name, safe="")
                + "/o/" + quote(ruta_storage, safe="")
                + "?alt=media&token=" + token_descarga
            )
            qr = qrcode.QRCode(box_size=8, border=3, error_correction=qrcode.constants.ERROR_CORRECT_M)
            qr.add_data(url_descarga)
            qr.make(fit=True)
            qr_img_pil = qr.make_image(fill_color="black", back_color="white").convert("RGB")
            buffer_recibo = generar_imagen_recibo(
                rec_id, fecha_actual, tot_ing, tot_gas, saldo, qr_img_pil, nombre_empresa
            )
            bytes_recibo = buffer_recibo.getvalue()
            blob_recibo.upload_from_string(bytes_recibo, content_type="image/png")
            st.session_state.rec_img_bytes = bytes_recibo
            st.session_state.rec_img_nombre = f"{rec_id}.png"
            registrar_auditoria("Generó comprobante descargable", rec_id)
            st.success("Comprobante guardado. Al escanear el QR se descarga el PNG.")
        except Exception as e:
            st.error(f"No se pudo guardar el comprobante con QR: {e}")

    if "rec_img_bytes" in st.session_state:
        st.image(st.session_state.rec_img_bytes, width=450)
        st.download_button(
            "Descargar comprobante PNG",
            data=st.session_state.rec_img_bytes,
            file_name=st.session_state.get("rec_img_nombre", "Comprobante_Financiero.png"),
            mime="image/png",
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

                nombre_id = generar_id("ARCH")
                institucion_id = get_institucion_id()
                doc_data = {
                    "ID": nombre_id,
                    "nombre": archivo_subido.name,
                    "tipo": archivo_subido.type,
                    "archivo_b64": base64_archivo,
                    "descripcion": descripcion_archivo if descripcion_archivo else "Sin descripción",
                    "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
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
    st.markdown('<p class="main-header">Comprobantes, OCR y revisión</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Carga uno o varios recibos de ingreso o gasto. La IA propone el tipo, la categoría y los datos; todo queda pendiente hasta que un aprobador lo revise.</p>', unsafe_allow_html=True)
    st.markdown("---")

    gemini_api_key = st.secrets.get("gemini", {}).get("api_key") if "gemini" in st.secrets else None
    puede_subir = tiene_permiso("subir_comprobantes")
    puede_revisar = tiene_permiso("aprobar_pendientes")

    archivos = []
    if not gemini_api_key and puede_subir:
        st.warning("Falta configurar [gemini].api_key en los secrets del proyecto.")
    elif puede_subir:
        archivos = st.file_uploader(
            "Sube recibos, facturas o comprobantes (PNG, JPG, PDF). Puedes seleccionar varios.",
            type=["png", "jpg", "jpeg", "pdf"],
            accept_multiple_files=True,
            key="upload_comprobantes_lote",
            help="Hasta 200 MB por archivo; los PDF se limitan a 50 MB por Gemini.",
        )
        if archivos and st.button(f"Analizar {len(archivos)} archivo(s) con IA", key="analizar_lote_ocr"):
            resultados = []
            mime_por_ext = {
                ".png": "image/png",
                ".jpg": "image/jpeg",
                ".jpeg": "image/jpeg",
                ".pdf": "application/pdf",
            }
            for archivo in archivos:
                datos_bytes = archivo.getvalue()
                hash_archivo = hashlib.sha256(datos_bytes).hexdigest()
                mime = mime_por_ext.get(os.path.splitext(archivo.name)[1].lower(), archivo.type or "application/octet-stream")
                extraidos = extraer_datos_comprobante_gemini(
                    datos_bytes, mime, archivo.name, gemini_api_key
                )
                if extraidos:
                    resultados.append({
                        "hash": hash_archivo,
                        "nombre": archivo.name,
                        "mime": mime,
                        "datos": extraidos,
                    })
            st.session_state.ocr_lote = resultados
            if resultados:
                st.success(f"Se analizaron {len(resultados)} archivo(s). Confirma cada uno para enviarlo a revisión.")
            else:
                st.warning("No se pudieron extraer datos de los archivos seleccionados.")

    if "ocr_lote" in st.session_state:
        archivos_actuales = {
            hashlib.sha256(a.getvalue()).hexdigest(): a for a in (archivos or [])
        } if puede_subir else {}
        if not archivos_actuales:
            st.info("Vuelve a seleccionar los archivos del lote para confirmar sus datos.")
        else:
            for item in list(st.session_state.ocr_lote):
                item_hash = item["hash"]
                archivo_original = archivos_actuales.get(item_hash)
                if not archivo_original:
                    continue
                datos = item["datos"]
                with st.expander(f"{item['nombre']} — confirmar datos", expanded=True):
                    tipo_sugerido = datos.get("tipo", "Indeterminado")
                    opciones_tipo = ["Seleccionar tipo", "Ingreso", "Gasto"]
                    tipo_default = tipo_sugerido if tipo_sugerido in ("Ingreso", "Gasto") else "Seleccionar tipo"
                    with st.form(f"confirmar_ocr_{item_hash[:12]}"):
                        tipo_confirmado = st.selectbox(
                            "Tipo de movimiento", opciones_tipo,
                            index=opciones_tipo.index(tipo_default),
                            key=f"tipo_ocr_{item_hash[:12]}",
                        )
                        c1, c2 = st.columns(2)
                        with c1:
                            concepto_confirmado = st.text_input(
                                "Concepto", value=str(datos.get("concepto") or ""),
                                key=f"concepto_ocr_{item_hash[:12]}",
                            )
                            try:
                                valor_default = float(datos.get("valor") or 0)
                            except (TypeError, ValueError):
                                valor_default = 0.0
                            valor_confirmado = st.number_input(
                                "Valor (COP)", min_value=0.0, value=max(0.0, valor_default),
                                step=1000.0, key=f"valor_ocr_{item_hash[:12]}",
                            )
                        with c2:
                            try:
                                fecha_default = dt_module.date.fromisoformat(str(datos.get("fecha")))
                            except (TypeError, ValueError):
                                fecha_default = dt_module.date.today()
                            fecha_confirmada = st.date_input(
                                "Fecha del comprobante", value=fecha_default,
                                key=f"fecha_ocr_{item_hash[:12]}",
                            )
                            responsable = st.selectbox(
                                "Responsable", obtener_lista_responsables(),
                                key=f"responsable_ocr_{item_hash[:12]}",
                            )
                        categoria = None
                        if tipo_confirmado == "Gasto":
                            sugerida = datos.get("categoria_sugerida", "Varios")
                            categoria = st.selectbox(
                                "Categoría sugerida por IA", CATEGORIAS_GASTO,
                                index=CATEGORIAS_GASTO.index(sugerida) if sugerida in CATEGORIAS_GASTO else 0,
                                key=f"categoria_ocr_{item_hash[:12]}",
                            )
                        observaciones = st.text_area(
                            "Observaciones / referencia",
                            value=str(datos.get("observaciones") or ""),
                            key=f"observaciones_ocr_{item_hash[:12]}",
                        )
                        enviar = st.form_submit_button("Enviar a revisión")

                    if enviar:
                        if tipo_confirmado not in ("Ingreso", "Gasto"):
                            st.error("Confirma si el movimiento es un ingreso o un gasto.")
                        elif not concepto_confirmado.strip() or valor_confirmado <= 0:
                            st.error("Completa un concepto y un valor mayor que cero.")
                        elif not tiene_permiso("registrar"):
                            st.error("Tu rango no permite enviar movimientos a revisión.")
                        else:
                            try:
                                ruta_evidencia = subir_evidencia_comprobante(
                                    archivo_original.getvalue(), item["nombre"], item["mime"]
                                )
                                movimiento = {
                                    "Fecha": fecha_confirmada.isoformat(),
                                    "Concepto": concepto_confirmado.strip(),
                                    "Valor": float(valor_confirmado),
                                    "Responsable": responsable,
                                    "Observaciones": observaciones,
                                    "Categoría": categoria if tipo_confirmado == "Gasto" else "",
                                    "Numero_Factura": datos.get("numero_factura"),
                                    "Hash_Evidencia": item_hash,
                                }
                                if tipo_confirmado == "Gasto":
                                    motivos = detectar_anomalias_gasto(
                                        movimiento["Valor"], movimiento["Fecha"], movimiento["Concepto"],
                                        movimiento["Categoría"], st.session_state.gastos_df,
                                        presupuesto_tope, cargar_presupuestos_categoria(),
                                    )
                                else:
                                    motivos = detectar_anomalias_ingreso(
                                        movimiento["Valor"], movimiento["Fecha"], movimiento["Concepto"],
                                        st.session_state.ingresos_df,
                                    )
                                motivos = ["Requiere aprobación humana antes de contabilizarse."] + motivos
                                pendiente_id = guardar_pendiente_movimiento(
                                    tipo_confirmado, movimiento, motivos, origen="OCR con IA",
                                    evidencia_ruta=ruta_evidencia,
                                    evidencia_nombre=item["nombre"], evidencia_mime=item["mime"],
                                )
                                if not pendiente_id:
                                    raise RuntimeError("No se pudo crear el pendiente en Firestore.")
                                registrar_auditoria(
                                    f"Envió {tipo_confirmado} OCR a revisión",
                                    f"{concepto_confirmado.strip()} - {valor_confirmado:,.0f} COP",
                                )
                                st.session_state.ocr_lote = [
                                    x for x in st.session_state.ocr_lote if x["hash"] != item_hash
                                ]
                                st.success(f"{tipo_confirmado} enviado a revisión. No se suma al balance hasta ser aprobado.")
                                st.rerun()
                            except Exception as e:
                                detalle_error = str(e)
                                if "does not exist" in detalle_error.lower() or "notfound" in detalle_error.lower():
                                    bucket_nombre = getattr(bucket, "name", FIREBASE_STORAGE_BUCKET or "sin configurar")
                                    st.error(
                                        f"El bucket de Firebase Storage '{bucket_nombre}' no existe. "
                                        "En Firebase Console > Storage > Files copia el nombre exacto y configúralo "
                                        "en Streamlit secrets como [firebase] storage_bucket = \"nombre-real-del-bucket\". "
                                        "Si aún no tienes bucket, créalo primero en Storage."
                                    )
                                else:
                                    st.error(f"No se pudo enviar el comprobante a revisión: {detalle_error}")

    st.markdown("---")
    st.markdown("### Ingresos y gastos pendientes de revisión")
    if not db:
        st.warning("Conecta Firebase para consultar las revisiones.")
    else:
        institucion_id = get_institucion_id()
        try:
            base_ref = db.collection("usuarios").document(institucion_id)
            gastos_docs = base_ref.collection("gastos_pendientes").where("Estado", "==", "Pendiente").limit(100).stream()
            ingresos_docs = base_ref.collection("ingresos_pendientes").where("Estado", "==", "Pendiente").limit(100).stream()
            pendientes = (
                [{**(p.to_dict() or {}), "Tipo": "Gasto"} for p in gastos_docs]
                + [{**(p.to_dict() or {}), "Tipo": "Ingreso"} for p in ingresos_docs]
            )
            if not pendientes:
                st.info("No hay ingresos ni gastos pendientes de revisión.")
            else:
                if len(pendientes) >= 200:
                    st.caption("Se muestran hasta 100 pendientes por tipo. Al aprobar o rechazar, se cargarán los siguientes.")
                for pendiente in pendientes:
                    tipo = pendiente["Tipo"]
                    prefijo_tipo = "🔴 Gasto" if tipo == "Gasto" else "🟢 Ingreso"
                    concepto = pendiente.get("Concepto", "Sin concepto")
                    valor = float(pendiente.get("Valor", 0) or 0)
                    fecha = pendiente.get("Fecha", "")
                    with st.expander(f"{prefijo_tipo} — {concepto} — {chr(36)}{valor:,.0f} ({fecha})"):
                        if tipo == "Gasto":
                            st.write(f"**Categoría:** {pendiente.get('Categoría', 'Varios')} | **Responsable:** {pendiente.get('Responsable', '')}")
                        else:
                            st.write(f"**Responsable:** {pendiente.get('Responsable', '')}")
                            rel_pendiente = calcular_relacion_presupuesto(
                                pd.to_numeric(st.session_state.ingresos_df.get("Valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum() + valor,
                                pd.to_numeric(st.session_state.gastos_df.get("Valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
                                presupuesto_tope,
                            )
                            st.caption("Si se aprueba, " + rel_pendiente["mensaje"])
                        if pendiente.get("Observaciones"):
                            st.write(f"**Observaciones:** {pendiente['Observaciones']}")
                        st.caption(f"Origen: {pendiente.get('Origen', 'Registro')} | Enviado por: {pendiente.get('creado_por', '—')}")
                        if pendiente.get("EvidenciaStoragePath") and puede_revisar:
                            try:
                                evidencia = descargar_evidencia_comprobante(pendiente["EvidenciaStoragePath"])
                                if evidencia:
                                    st.download_button(
                                        "Descargar comprobante original",
                                        data=evidencia,
                                        file_name=pendiente.get("EvidenciaNombre", "comprobante"),
                                        mime=pendiente.get("EvidenciaMime", "application/octet-stream"),
                                        key=f"evidencia_{pendiente['ID']}",
                                    )
                            except Exception as e:
                                st.warning(f"No se pudo abrir el comprobante original: {e}")
                        for motivo in pendiente.get("Motivos", []):
                            st.caption("• " + str(motivo))

                        if not puede_revisar:
                            st.caption("Tu rango puede consultar la solicitud, pero no aprobarla ni rechazarla.")
                        else:
                            col_aprobar, col_rechazar = st.columns(2)
                            coleccion_pend = "gastos_pendientes" if tipo == "Gasto" else "ingresos_pendientes"
                            coleccion_final = "gastos" if tipo == "Gasto" else "ingresos"
                            with col_aprobar:
                                if st.button("Aprobar y contabilizar", key=f"aprobar_{pendiente['ID']}"):
                                    nuevo_reg = {
                                        k: pendiente[k] for k in (
                                            "Fecha", "Concepto", "Categoría", "Valor", "Responsable",
                                            "Observaciones", "Numero_Factura", "Origen",
                                            "EvidenciaStoragePath", "EvidenciaNombre", "Hash_Evidencia",
                                        ) if k in pendiente
                                    }
                                    nuevo_reg["ID"] = generar_id("GAS" if tipo == "Gasto" else "ING")
                                    try:
                                        ref_pend = base_ref.collection(coleccion_pend).document(pendiente["ID"])
                                        ref_final = base_ref.collection(coleccion_final).document(nuevo_reg["ID"])
                                        batch = db.batch()
                                        batch.set(ref_final, nuevo_reg)
                                        batch.update(ref_pend, {
                                            "Estado": "Aprobado",
                                            "revisado_por": (st.session_state.user_data or {}).get("email", ""),
                                            "revisado_en": dt_module.datetime.now(dt_module.timezone.utc).isoformat(),
                                        })
                                        batch.commit()
                                        registrar_auditoria(f"Aprobó {tipo} pendiente", f"{concepto} - {valor:,.0f} COP")
                                        cargar_datos_nube()
                                        st.success(f"{tipo} aprobado y contabilizado.")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"No se pudo aprobar el pendiente: {e}")
                            with col_rechazar:
                                if st.button("Rechazar", key=f"rechazar_{pendiente['ID']}"):
                                    try:
                                        base_ref.collection(coleccion_pend).document(pendiente["ID"]).update({
                                            "Estado": "Rechazado",
                                            "revisado_por": (st.session_state.user_data or {}).get("email", ""),
                                            "revisado_en": dt_module.datetime.now(dt_module.timezone.utc).isoformat(),
                                        })
                                        registrar_auditoria(f"Rechazó {tipo} pendiente", f"{concepto} - {valor:,.0f} COP")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"No se pudo rechazar el pendiente: {e}")
        except Exception as e:
            st.warning(f"No se pudieron cargar los pendientes: {e}")

elif menu == "11. Gestión de Equipo":
    st.markdown('<p class="main-header">Equipo y rangos de acceso</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Invita correos a la nube de esta empresa y define cargos con permisos claros.</p>', unsafe_allow_html=True)
    st.markdown("---")

    if not tiene_permiso("gestionar_equipo"):
        st.error("Tu rango no permite administrar el equipo.")
    elif not db:
        st.warning("Conecta Firebase para gestionar el equipo.")
    else:
        institucion_id = get_institucion_id()
        roles_custom = cargar_roles_personalizados()
        permisos_legibles = {
            clave: f"{clave}: {texto}" for clave, texto in PERMISOS_DISPONIBLES.items()
        }
        with st.expander("Crear rango personalizado", expanded=True):
            with st.form("form_crear_rango"):
                nombre_rango = st.text_input(
                    "Nombre del rango",
                    help="Usa de 2 a 40 letras, números, espacios o guiones.",
                )
                permisos_nuevos = st.multiselect(
                    "Funciones del rango",
                    options=list(PERMISOS_DISPONIBLES.keys()),
                    format_func=lambda p: permisos_legibles[p],
                )
                guardar_rango = st.form_submit_button("Guardar rango")
            if guardar_rango:
                nombre_rango = nombre_rango.strip()
                if not re.fullmatch(r"[A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ _-]{2,40}", nombre_rango):
                    st.error("Escribe un nombre de rango válido (2 a 40 caracteres).")
                elif nombre_rango in ROLES_DISPONIBLES:
                    st.error("Ese nombre está reservado para un rango integrado.")
                else:
                    roles_custom[nombre_rango] = {"permisos": sorted(set(permisos_nuevos))}
                    db.collection("usuarios").document(institucion_id).set(
                        {"roles_personalizados": roles_custom}, merge=True
                    )
                    registrar_auditoria("Guardó rango personalizado", nombre_rango)
                    st.success("Rango guardado.")
                    st.rerun()

        if roles_custom:
            st.markdown("### Rangos personalizados")
            for nombre, config in roles_custom.items():
                with st.expander(nombre):
                    seleccionados = [
                        p for p in config.get("permisos", [])
                        if p in PERMISOS_DISPONIBLES
                    ]
                    key_rango = hashlib.sha256(nombre.encode()).hexdigest()[:12]
                    permisos_editados = st.multiselect(
                        "Funciones", options=list(PERMISOS_DISPONIBLES.keys()),
                        default=seleccionados,
                        format_func=lambda p: permisos_legibles[p],
                        key=f"permisos_rango_{key_rango}",
                    )
                    col_guardar_rol, col_borrar_rol = st.columns(2)
                    with col_guardar_rol:
                        if st.button("Actualizar permisos", key=f"actualizar_rango_{key_rango}"):
                            roles_custom[nombre] = {"permisos": sorted(set(permisos_editados))}
                            db.collection("usuarios").document(institucion_id).set(
                                {"roles_personalizados": roles_custom}, merge=True
                            )
                            registrar_auditoria("Actualizó permisos de rango", nombre)
                            st.success("Permisos actualizados.")
                            st.rerun()
                    with col_borrar_rol:
                        if st.button("Eliminar rango", key=f"eliminar_rango_{key_rango}"):
                            miembros = db.collection("usuarios").document(institucion_id).collection("empleados").where("rol", "==", nombre).limit(1).stream()
                            if next(iter(miembros), None):
                                st.error("No se puede eliminar: hay empleados asignados a este rango.")
                            else:
                                roles_custom.pop(nombre, None)
                                db.collection("usuarios").document(institucion_id).set(
                                    {"roles_personalizados": roles_custom}, merge=True
                                )
                                registrar_auditoria("Eliminó rango personalizado", nombre)
                                st.rerun()

        roles_asignables = obtener_roles_asignables()
        with st.expander("Invitar empleado", expanded=True):
            with st.form("form_invitar_empleado"):
                correo_nuevo = st.text_input("Correo de la persona")
                nombre_nuevo = st.text_input("Nombre (opcional)")
                rol_nuevo = st.selectbox("Rango", roles_asignables)
                invitar = st.form_submit_button("Invitar")
            if invitar:
                correo_limpio = correo_nuevo.lower().strip()
                if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", correo_limpio):
                    st.error("Escribe un correo válido.")
                elif correo_limpio == institucion_id:
                    st.error("Ese correo ya es el propietario de esta empresa.")
                else:
                    nombre_empresa = (st.session_state.user_data or {}).get("institucion", "")
                    miembro_ref = db.collection("usuarios").document(institucion_id).collection("empleados").document(correo_limpio)
                    miembro_ref.set({
                        "email": correo_limpio,
                        "nombre": nombre_nuevo.strip() or correo_limpio,
                        "rol": rol_nuevo,
                        "estado": "Activo",
                        "fecha_invitacion": dt_module.datetime.now(dt_module.timezone.utc).isoformat(),
                    })
                    db.collection("empleados_index").document(correo_limpio).set({
                        "empresa_id": institucion_id,
                        "rol": rol_nuevo,
                        "nombre_empresa": nombre_empresa,
                    })
                    registrar_auditoria("Invitó empleado", f"{correo_limpio} — {rol_nuevo}")
                    st.success(f"{correo_limpio} fue invitado como {rol_nuevo}.")
                    st.rerun()

        st.markdown("### Equipo de la empresa")
        try:
            empleados_docs = list(
                db.collection("usuarios").document(institucion_id).collection("empleados").stream()
            )
            if not empleados_docs:
                st.info("Aún no hay empleados invitados.")
            for doc_empleado in empleados_docs:
                empleado = doc_empleado.to_dict() or {}
                correo = empleado.get("email", doc_empleado.id)
                estado = empleado.get("estado", "Activo")
                key_correo = hashlib.sha256(correo.encode()).hexdigest()[:12]
                with st.expander(f"{empleado.get('nombre', correo)} — {empleado.get('rol', 'Analista')} ({estado})"):
                    st.write(f"**Correo:** {correo}")
                    if estado == "Revocado":
                        st.caption("Acceso revocado.")
                    else:
                        rol_actual_emp = empleado.get("rol", "Analista")
                        opciones_roles = obtener_roles_asignables()
                        indice = opciones_roles.index(rol_actual_emp) if rol_actual_emp in opciones_roles else 0
                        nuevo_rol = st.selectbox(
                            "Rango", opciones_roles, index=indice,
                            key=f"rol_empleado_{key_correo}",
                        )
                        col_actualizar, col_revocar = st.columns(2)
                        with col_actualizar:
                            if st.button("Actualizar rango", key=f"actualizar_empleado_{key_correo}"):
                                miembro_ref = db.collection("usuarios").document(institucion_id).collection("empleados").document(correo)
                                miembro_ref.update({"rol": nuevo_rol})
                                db.collection("empleados_index").document(correo).set({
                                    "empresa_id": institucion_id,
                                    "rol": nuevo_rol,
                                    "nombre_empresa": (st.session_state.user_data or {}).get("institucion", ""),
                                }, merge=True)
                                user_ref = db.collection("usuarios").document(correo)
                                if user_ref.get().exists:
                                    user_ref.update({"rol": nuevo_rol})
                                registrar_auditoria("Actualizó rango de empleado", f"{correo} → {nuevo_rol}")
                                st.success("Rango actualizado.")
                                st.rerun()
                        with col_revocar:
                            if st.button("Revocar acceso", key=f"revocar_empleado_{key_correo}"):
                                db.collection("usuarios").document(institucion_id).collection("empleados").document(correo).update({
                                    "estado": "Revocado",
                                    "revocado_en": dt_module.datetime.now(dt_module.timezone.utc).isoformat(),
                                })
                                db.collection("empleados_index").document(correo).delete()
                                registrar_auditoria("Revocó acceso de empleado", correo)
                                st.warning("Acceso revocado. Su sesión se cerrará en la próxima interacción.")
                                st.rerun()
        except Exception as e:
            st.warning(f"No se pudo cargar el equipo: {e}")
