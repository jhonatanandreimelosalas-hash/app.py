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
import datetime as dt_module  # módulo único para fechas y horas
import uuid
import base64
from pathlib import Path

# --- CONFIGURACIÓN DE PÁGINA ---
LOGO_PATH = Path(__file__).resolve().parent / "assets" / "logo_financiero.png"
st.set_page_config(
    page_title="Portal de Gestión Financiera",
    page_icon=str(LOGO_PATH) if LOGO_PATH.is_file() else "💰",
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

    Las tablas y las gráficas usan el mismo selector para conservar el contraste
    aunque el tema interno del navegador o de Streamlit sea distinto.
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

        /* Superficies principales: el color de la app no depende del tema del navegador */
        html, body, .stApp {{ background: {fondo_app} !important; color: {texto_principal} !important; color-scheme: {color_scheme}; }}

        /* Header/toolbar nativos de Streamlit — antes quedaban sin colorear y
           creaban la franja clara/oscura desalineada que reportaste */
        [data-testid="stHeader"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stToolbar"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stDecoration"] {{ background-image: none !important; background-color: {acento} !important; }}
        [data-testid="stAppViewContainer"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stMain"], [data-testid="stMainBlockContainer"] {{ background-color: {fondo_app} !important; color: {texto_principal} !important; }}
        [data-testid="stBottomBlockContainer"] {{ background-color: {fondo_app} !important; }}
        [data-testid="stSidebar"] > div:first-child {{ background-color: {superficie} !important; }}

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
        [data-baseweb="popover"], [data-baseweb="menu"], [role="listbox"] {{ background-color: {superficie} !important; color: {texto_principal} !important; }}
        [role="option"] {{ color: {texto_principal} !important; }}

        /* Forzar colores en campos de entrada, áreas de texto y selectores */
        input, textarea, select {{ background-color: {input_bg} !important; color: {texto_principal} !important; border-color: {borde} !important; }}
        div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div {{ background-color: {input_bg} !important; color: {texto_principal} !important; border-color: {borde} !important; }}

        /* Corregir el texto dentro de los inputs de Streamlit */
        input::-webkit-input-placeholder {{ color: {texto_secundario} !important; }}
    </style>
    """


def aplicar_tema_grafico(figura):
    """Alinea las gráficas de Plotly con el selector claro/oscuro de la app."""
    oscuro = st.session_state.get("modo_oscuro", False)
    figura.update_layout(
        template="plotly_dark" if oscuro else "plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#E2E8F0" if oscuro else "#1E293B",
    )
    return figura


def mostrar_tabla_adaptativa(df, hide_index=True, columna_alerta=None):
    """Renderiza una tabla HTML segura que sigue el tema aunque Streamlit use otro."""
    if df is None or df.empty:
        st.info("No hay datos para mostrar.")
        return
    oscuro = st.session_state.get("modo_oscuro", False)
    fondo = "#0F172A" if oscuro else "#FFFFFF"
    superficie = "#1E293B" if oscuro else "#F8FAFC"
    texto = "#E2E8F0" if oscuro else "#1E293B"
    borde = "#334155" if oscuro else "#E2E8F0"
    tabla = df.to_html(index=not hide_index, escape=True, border=0, classes="tabla-adaptativa")
    if columna_alerta and columna_alerta in df.columns:
        filas = []
        inicio = tabla.find("<tbody>")
        fin = tabla.find("</tbody>", inicio)
        if inicio >= 0 and fin >= 0:
            cuerpo = tabla[inicio:fin]
            posicion = 0
            for _, fila in df.iterrows():
                fila_fin = cuerpo.find("<tr", posicion)
                if fila_fin < 0:
                    break
                cierre = cuerpo.find(">", fila_fin)
                if cierre < 0:
                    break
                try:
                    alerta = float(fila[columna_alerta]) > 0
                except (TypeError, ValueError):
                    alerta = False
                if alerta:
                    cuerpo = cuerpo[:fila_fin] + '<tr class="fila-alerta"' + cuerpo[cierre:]
                    cierre = cuerpo.find(">", fila_fin)
                posicion = cierre + 1
            tabla = tabla[:inicio] + cuerpo + tabla[fin:]
    color_alerta = "#7F1D1D" if oscuro else "#FEE2E2"
    st.markdown(
        f'''<div style="overflow-x:auto;border:1px solid {borde};border-radius:10px;background:{fondo};">
        <style>
        .tabla-adaptativa {{ width:100%; border-collapse:collapse; color:{texto}; font-size:0.92rem; }}
        .tabla-adaptativa th {{ position:sticky;top:0;background:{superficie};text-align:left;font-weight:700; }}
        .tabla-adaptativa th,.tabla-adaptativa td {{ padding:0.55rem 0.75rem;border-bottom:1px solid {borde}; }}
        .tabla-adaptativa tbody tr:nth-child(even) {{ background:{superficie}; }}
        .tabla-adaptativa tbody tr.fila-alerta {{ background:{color_alerta}; }}
        </style>{tabla}</div>''',
        unsafe_allow_html=True,
    )

st.markdown(render_estilos(st.session_state.modo_oscuro), unsafe_allow_html=True)

if st.sidebar.toggle("Modo oscuro", key="modo_oscuro_toggle", value=st.session_state.get('modo_oscuro', False)):
    if not st.session_state.modo_oscuro:
        st.session_state.modo_oscuro = True
        st.rerun()
else:
    if st.session_state.modo_oscuro:
        st.session_state.modo_oscuro = False
        st.rerun()

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


def generar_id_registro(prefijo):
    """Crea IDs únicos incluso cuando se procesa un lote de movimientos."""
    sello = dt_module.datetime.now().strftime("%Y%m%d%H%M%S%f")
    return f"{prefijo}-{sello}-{uuid.uuid4().hex[:8]}"


def limpiar_datos_privados_sesion():
    """Borra datos y adjuntos de la empresa anterior al salir o cambiar de cuenta."""
    st.session_state.ingresos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Valor", "Responsable", "Observaciones", "ID"])
    st.session_state.gastos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Categoría", "Valor", "Responsable", "ID"])
    for llave in ("ocr_comprobantes", "factura_extraida", "ingreso_extraido", "rec_img_bytes", "excel_reporte_bytes", "pdf_reporte_bytes", "roles_empresa_cache"):
        st.session_state.pop(llave, None)
    for llave in list(st.session_state.keys()):
        if str(llave).startswith(("upload_comprobantes_lote_", "repositorio_upload_")):
            st.session_state.pop(llave, None)


def validar_y_actualizar_acceso_empresa():
    """Al recargar la app, actualiza el rango del empleado y bloquea accesos revocados."""
    usuario = st.session_state.get("user_data") or {}
    empresa_id = usuario.get("empresa_id")
    correo = (usuario.get("email") or "").lower().strip()
    if not db or not empresa_id or empresa_id.lower().strip() == correo:
        return True
    try:
        miembro_ref = db.collection("usuarios").document(empresa_id).collection("empleados").document(correo)
        miembro_doc = miembro_ref.get()
        miembro = miembro_doc.to_dict() or {}
        if not miembro_doc.exists or miembro.get("estado") != "Activo":
            st.session_state.logged_in = False
            st.session_state.user_data = None
            st.session_state.roles_empresa_cache = None
            st.error("El acceso a esta empresa fue revocado o la invitación ya no está activa. Contacta al propietario.")
            return False
        usuario["rol"] = miembro.get("rol", "Analista")
        if miembro.get("nombre"):
            usuario["institucion"] = miembro["nombre"]
        st.session_state.user_data = usuario
        return True
    except Exception:
        st.error("No se pudo validar tu pertenencia a la empresa. Intenta de nuevo cuando haya conexión.")
        return False


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

ROLES_BASE = {
    "Propietario": {"ver_finanzas", "gestionar_equipo", "gestionar_roles", "gestionar_archivos", "aprobar_pendientes", "configurar_presupuesto", "ver_auditoria", "registrar", "eliminar_registros"},
    "Gerente Financiero": {"ver_finanzas", "gestionar_equipo", "gestionar_archivos", "aprobar_pendientes", "configurar_presupuesto", "ver_auditoria", "registrar", "eliminar_registros"},
    "Analista": {"ver_finanzas", "gestionar_archivos", "registrar"},
    "Auditor": {"ver_finanzas", "ver_auditoria"},
}

PERMISOS_DISPONIBLES = {
    "ver_finanzas": "Consultar registros e informes financieros",
    "registrar": "Registrar ingresos o gastos",
    "eliminar_registros": "Eliminar ingresos o gastos registrados",
    "aprobar_pendientes": "Aprobar o rechazar movimientos pendientes",
    "configurar_presupuesto": "Configurar presupuestos y períodos",
    "ver_auditoria": "Consultar la auditoría de la empresa",
    "gestionar_equipo": "Invitar y administrar integrantes",
    "gestionar_archivos": "Guardar y eliminar archivos de la empresa",
}


def obtener_configuracion_roles(forzar=False):
    """Devuelve permisos configurados por la empresa y conserva los roles base."""
    empresa_id = get_institucion_id()
    clave_cache = st.session_state.get("roles_empresa_cache")
    if not forzar and clave_cache and clave_cache.get("empresa_id") == empresa_id:
        return clave_cache.get("roles", {})

    roles = {nombre: sorted(permisos) for nombre, permisos in ROLES_BASE.items()}
    if db and empresa_id:
        try:
            datos = db.collection("usuarios").document(empresa_id).get().to_dict() or {}
            personalizados = datos.get("roles_config", {})
            if isinstance(personalizados, dict):
                for nombre, permisos in personalizados.items():
                    if nombre and nombre != "Propietario" and isinstance(permisos, list):
                        roles[str(nombre)] = [p for p in permisos if p in PERMISOS_DISPONIBLES]
        except Exception:
            pass
    st.session_state.roles_empresa_cache = {"empresa_id": empresa_id, "roles": roles}
    return roles


def obtener_nombres_roles():
    roles = obtener_configuracion_roles()
    return ["Propietario"] + sorted(nombre for nombre in roles if nombre != "Propietario")


def tiene_permiso(accion: str) -> bool:
    if not st.session_state.user_data:
        return False
    rol_actual = st.session_state.user_data.get('rol', 'Propietario')
    if rol_actual == "Propietario":
        return True
    permisos = obtener_configuracion_roles().get(rol_actual)
    if permisos is None:
        return accion in ROLES_BASE.get("Analista", set())
    return accion in permisos


def obtener_nombre_usuario_actual():
    usuario = st.session_state.get("user_data") or {}
    return usuario.get("institucion") or usuario.get("email") or "Usuario"


def obtener_nombre_empresa():
    empresa_id = get_institucion_id()
    if db and empresa_id:
        try:
            datos = db.collection("usuarios").document(empresa_id).get().to_dict() or {}
            return datos.get("institucion") or empresa_id
        except Exception:
            pass
    return empresa_id or "Empresa"


def obtener_lista_responsables():
    """Muestra una muestra acotada del equipo para no leer miles de documentos."""
    institucion_id = get_institucion_id()
    if not db or not institucion_id:
        return [obtener_nombre_usuario_actual()]
    try:
        nombres = []
        doc_empresa = db.collection('usuarios').document(institucion_id).get()
        datos_empresa = doc_empresa.to_dict() or {}
        nombre_dueno = datos_empresa.get('institucion')
        if nombre_dueno:
            nombres.append(nombre_dueno)

        empleados_ref = (
            db.collection('usuarios').document(institucion_id).collection('empleados')
            .where('estado', '==', 'Activo').limit(100).stream()
        )
        for emp in empleados_ref:
            emp_data = emp.to_dict()
            nombres.append(emp_data.get('nombre') or emp_data.get('email', ''))

        if nombres:
            return nombres
        return [obtener_nombre_usuario_actual()]
    except Exception:
        return [obtener_nombre_usuario_actual()]


def cargar_datos_nube():
    columnas_ing = ["Fecha", "Concepto", "Valor", "Responsable", "Observaciones", "ID"]
    columnas_gas = ["Fecha", "Concepto", "Categoría", "Valor", "Responsable", "ID"]
    st.session_state.ingresos_df = pd.DataFrame(columns=columnas_ing)
    st.session_state.gastos_df = pd.DataFrame(columns=columnas_gas)
    if not db:
        return
    institucion_id = get_institucion_id()
    if not institucion_id:
        return
    try:
        base_ref = db.collection('usuarios').document(institucion_id)

        ing_docs = base_ref.collection('ingresos').stream()
        ing_data = [doc.to_dict() for doc in ing_docs]
        gas_docs = base_ref.collection('gastos').stream()
        gas_data = [doc.to_dict() for doc in gas_docs]

        if ing_data:
            st.session_state.ingresos_df = pd.DataFrame(ing_data)
        if gas_data:
            st.session_state.gastos_df = pd.DataFrame(gas_data)
    except Exception:
        st.session_state.ingresos_df = pd.DataFrame(columns=columnas_ing)
        st.session_state.gastos_df = pd.DataFrame(columns=columnas_gas)
        st.sidebar.error("Error al sincronizar con la nube.")


def guardar_registro_nube(coleccion, datos):
    if not db:
        return False
    institucion_id = get_institucion_id()
    if not institucion_id:
        return False
    try:
        db.collection('usuarios').document(institucion_id).collection(coleccion).document(datos['ID']).set(datos)
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
        db.collection('usuarios').document(institucion_id).collection(coleccion).document(doc_id).delete()
        return True
    except Exception:
        return False


CATEGORIAS_GASTO = ["Logística", "Publicidad", "Alimentación", "Varios"]


def cargar_presupuesto_general():
    hoy = dt_module.date.today()
    valores = {
        "monto": 0.0,
        "fecha_inicio": hoy.isoformat(),
        "fecha_fin": (hoy + dt_module.timedelta(days=30)).isoformat(),
        "configurado": False,
    }
    empresa_id = get_institucion_id()
    if not db or not empresa_id:
        return valores
    try:
        datos = db.collection("usuarios").document(empresa_id).get().to_dict() or {}
        guardado = datos.get("presupuesto_general")
        if isinstance(guardado, dict):
            valores.update(guardado)
            valores["configurado"] = True
    except Exception:
        pass
    return valores


def guardar_presupuesto_general(monto, fecha_inicio, fecha_fin):
    empresa_id = get_institucion_id()
    if not db or not empresa_id:
        return False
    try:
        db.collection("usuarios").document(empresa_id).set({
            "presupuesto_general": {
                "monto": float(monto),
                "fecha_inicio": fecha_inicio.isoformat(),
                "fecha_fin": fecha_fin.isoformat(),
            }
        }, merge=True)
        return True
    except Exception:
        return False


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
        return False
    institucion_id = get_institucion_id()
    if not institucion_id:
        return False
    try:
        db.collection('usuarios').document(institucion_id).set(
            {'presupuestos_categoria': presupuestos}, merge=True
        )
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


def extraer_datos_comprobante_gemini(archivo_bytes, mime_type, api_key):
    """Clasifica un comprobante y extrae sus campos en una sola llamada de IA."""
    try:
        from google import genai
        from google.genai import types
        import json as json_lib

        cliente = genai.Client(api_key=api_key)
        instrucciones = (
            "Analiza este comprobante financiero en español. Decide si documenta un dinero "
            "recibido por la empresa (ingreso) o un pago realizado por la empresa (gasto). "
            "Devuelve solamente JSON válido, sin Markdown, con las claves: tipo (ingreso o gasto), "
            "concepto, valor (número COP sin separadores, null si no es legible), fecha (YYYY-MM-DD o null), "
            "observaciones, categoria_sugerida (una de Logística, Publicidad, Alimentación, Varios), "
            "numero_factura. Si no puedes clasificarlo con seguridad, usa tipo ingreso solo cuando el "
            "documento muestre claramente dinero recibido; en los demás casos usa gasto."
        )
        respuesta = cliente.models.generate_content(
            model="gemini-3.6-flash",
            contents=[types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type), instrucciones],
        )
        texto = (respuesta.text or "").strip()
        if texto.startswith("```"):
            texto = texto.split("\n", 1)[-1].rsplit("```", 1)[0].strip()
        extraidos = json_lib.loads(texto)
        if not isinstance(extraidos, dict):
            return None
        extraidos["tipo"] = "ingreso" if str(extraidos.get("tipo", "")).strip().lower() == "ingreso" else "gasto"
        return extraidos
    except Exception as e:
        st.error(f"No se pudo analizar el comprobante: {e}")
        return None


def guardar_comprobante_nube(nombre, mime_type, archivo_bytes, registro_id):
    """Guarda el original bajo la carpeta de la empresa en Firebase Storage."""
    empresa_id = get_institucion_id()
    if not bucket or not empresa_id:
        return None
    try:
        nombre_limpio = os.path.basename(nombre).replace("\\", "_").replace("/", "_")
        ruta = f"empresas/{empresa_id}/comprobantes/{registro_id}_{nombre_limpio}"
        blob = bucket.blob(ruta)
        blob.upload_from_string(archivo_bytes, content_type=mime_type)
        return ruta
    except Exception:
        return None


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
    """Resume ingresos, gastos, saldo y uso del límite, sin bloquear ingresos."""
    total_ingresos = float(total_ingresos or 0)
    total_gastos = float(total_gastos or 0)
    presupuesto_tope = float(presupuesto_tope or 0)
    saldo = total_ingresos - total_gastos
    if not presupuesto_tope or presupuesto_tope <= 0:
        return {"nivel": "sin_presupuesto", "mensaje": "Aún no se ha configurado un presupuesto general para esta empresa.", "saldo": saldo, "porcentaje_ingresos": 0.0, "porcentaje_gastado": 0.0}

    porcentaje_gastado = (total_gastos / presupuesto_tope) * 100
    porcentaje_ingresos = (total_ingresos / presupuesto_tope) * 100
    porcentaje_cubierto_gastos = (total_ingresos / total_gastos) * 100 if total_gastos > 0 else 0.0
    texto_cobertura = (
        f"Los ingresos cubren el {porcentaje_cubierto_gastos:.0f}% de los gastos registrados."
        if total_gastos > 0 else "Aún no hay gastos registrados para comparar con los ingresos."
    )

    if porcentaje_gastado >= 100:
        nivel = "critico"
        mensaje = f"Los gastos superan el límite en ${total_gastos - presupuesto_tope:,.0f}. Los ingresos equivalen al {porcentaje_ingresos:.0f}% del presupuesto. {texto_cobertura}"
    elif porcentaje_gastado >= 80:
        nivel = "alerta"
        mensaje = f"Se ha usado el {porcentaje_gastado:.0f}% del límite (${presupuesto_tope - total_gastos:,.0f} disponibles). Los ingresos equivalen al {porcentaje_ingresos:.0f}% del presupuesto. {texto_cobertura}"
    else:
        nivel = "ok"
        mensaje = f"Se ha usado el {porcentaje_gastado:.0f}% del límite. Los ingresos equivalen al {porcentaje_ingresos:.0f}% del presupuesto. {texto_cobertura}"

    return {"nivel": nivel, "mensaje": mensaje, "saldo": saldo, "porcentaje_ingresos": porcentaje_ingresos, "porcentaje_gastado": porcentaje_gastado}


def filtrar_movimientos_periodo(df, periodo):
    """Filtra un registro por las fechas inclusivas del presupuesto activo."""
    if df is None or df.empty or not isinstance(periodo, tuple) or len(periodo) != 2 or "Fecha" not in df.columns:
        return df.copy() if df is not None else pd.DataFrame()
    copia = df.copy()
    fechas = pd.to_datetime(copia["Fecha"], errors="coerce").dt.date
    return copia.loc[fechas.ge(periodo[0]) & fechas.le(periodo[1])].copy()


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
            "institucion": obtener_nombre_empresa(),
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
        limpiar_datos_privados_sesion()
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
    nombre_empresa = str(nombre_empresa or "Empresa")[:42]
    draw.text((30, 25), nombre_empresa, fill="#FFFFFF", font=font_title)
    draw.text((30, 60), "Comprobante General de Balance Financiero", fill="#93C5FD", font=font_regular)

    draw.rectangle([(30, 130), (img_w - 30, img_h - 40)], outline="#E2E8F0", width=2, fill="#F8FAFC")

    draw.text((55, 160), "ID de Comprobante:", fill="#64748B", font=font_small)
    draw.text((200, 158), f"{rec_id}", fill="#1E293B", font=font_bold)

    draw.text((55, 190), "Fecha de Emisión:", fill="#64748B", font=font_small)
    draw.text((200, 188), f"{fecha}", fill="#1E293B", font=font_bold)

    draw.text((55, 220), "Institución:", fill="#64748B", font=font_small)
    draw.text((200, 218), nombre_empresa, fill="#1E293B", font=font_bold)

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
        Paragraph(f"Generado el {dt_module.datetime.now().strftime('%Y-%m-%d %H:%M')}", estilo_subtitulo),
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
        columnas_ing = [c for c in df_ingresos.columns if c not in ("ID", "Ruta_Archivo")]
        filas_ing = [columnas_ing] + df_ingresos[columnas_ing].astype(str).values.tolist()
        tabla_ing = Table(filas_ing, repeatRows=1)
        tabla_ing.setStyle(_estilo_tabla_corporativo())
        story.append(tabla_ing)
    else:
        story.append(Paragraph("No hay ingresos registrados.", styles['Normal']))

    story.append(Spacer(1, 16))
    story.append(Paragraph("Detalle de Gastos", estilo_seccion))
    if not df_gastos.empty:
        columnas_gas = [c for c in df_gastos.columns if c not in ("ID", "Ruta_Archivo")]
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

    if LOGO_PATH.is_file():
        logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
        st.markdown(
            f'<div style="display:flex; justify-content:center; margin:0.5rem 0 1rem;">'
            f'<img src="data:image/png;base64,{logo_b64}" alt="Logo financiero" '
            f'style="width:150px; max-width:35vw; height:auto;">'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown('<p class="main-header" style="text-align: center;">Portal Financiero Institucional</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header" style="text-align: center;">Finanzas y control para organizaciones</p>', unsafe_allow_html=True)

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
                            limpiar_datos_privados_sesion()
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
                                'fecha_creacion': dt_module.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
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
                                'fecha_creacion': dt_module.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
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
                        msg['Subject'] = "Recuperación de contraseña - Portal Financiero"

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

if not validar_y_actualizar_acceso_empresa():
    st.stop()

# --- INICIALIZAR ESTADO DE OMITIR ALERTA ---
if "omitir_alerta_presupuesto" not in st.session_state:
    st.session_state.omitir_alerta_presupuesto = False

# --- MENÚ LATERAL ---
st.sidebar.markdown(f"👋 **Hola, {st.session_state.user_data['institucion']}**")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.user_data = None
    st.session_state.omitir_alerta_presupuesto = False
    limpiar_datos_privados_sesion()
    st.session_state.equipo_mostrar_hasta = 100
    st.rerun()

st.sidebar.markdown("---")
empresa_email_base = get_institucion_id() or "Sin empresa vinculada"
st.sidebar.caption(f"**Empresa:** {obtener_nombre_empresa()}\n\nCorreo base: {empresa_email_base}")
st.sidebar.markdown("⚙️ **Presupuesto de la empresa**")

config_presupuesto = cargar_presupuesto_general()
hoy = dt_module.date.today()
try:
    fecha_inicio_guardada = dt_module.date.fromisoformat(config_presupuesto.get("fecha_inicio", hoy.isoformat()))
except (TypeError, ValueError):
    fecha_inicio_guardada = hoy
try:
    fecha_fin_guardada = dt_module.date.fromisoformat(config_presupuesto.get("fecha_fin", (hoy + dt_module.timedelta(days=30)).isoformat()))
except (TypeError, ValueError):
    fecha_fin_guardada = hoy + dt_module.timedelta(days=30)

puede_configurar_presupuesto = tiene_permiso("configurar_presupuesto")
with st.sidebar.form("form_presupuesto_general"):
    periodo_presupuesto = st.date_input(
        "📅 Período de ejecución",
        value=(fecha_inicio_guardada, fecha_fin_guardada),
        disabled=not puede_configurar_presupuesto,
        key=f"periodo_presupuesto_{empresa_email_base}",
    )
    presupuesto_tope = st.number_input(
        "Límite de gastos ($)", min_value=0.0,
        value=float(config_presupuesto.get("monto", 0.0)), step=50000.0,
        disabled=not puede_configurar_presupuesto,
        key=f"monto_presupuesto_{empresa_email_base}",
    )
    guardar_presupuesto = st.form_submit_button("Guardar presupuesto", disabled=not puede_configurar_presupuesto)

if guardar_presupuesto:
    if not isinstance(periodo_presupuesto, tuple) or len(periodo_presupuesto) != 2 or periodo_presupuesto[1] < periodo_presupuesto[0]:
        st.sidebar.error("Selecciona un período válido.")
    elif guardar_presupuesto_general(presupuesto_tope, periodo_presupuesto[0], periodo_presupuesto[1]):
        registrar_auditoria("Actualizó presupuesto general", f"${presupuesto_tope:,.0f} | {periodo_presupuesto[0]} a {periodo_presupuesto[1]}")
        st.sidebar.success("Presupuesto guardado para toda la empresa.")
        st.rerun()
    else:
        st.sidebar.error("No se pudo guardar el presupuesto en la nube.")
if not puede_configurar_presupuesto:
    st.sidebar.caption("Solo lectura según tu rango.")

if config_presupuesto.get("configurado") and isinstance(periodo_presupuesto, tuple) and len(periodo_presupuesto) == 2:
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
opciones_menu = []
if tiene_permiso("ver_finanzas"):
    opciones_menu.extend([
        "1. Inicio", "2. Registro de Ingresos", "3. Registro de Gastos",
        "4. Balance Financiero", "5. Dashboard y Gráficos",
        "6. Anexo de Recibos & QR", "7. Gestión de Archivos", "8. Reporte Final",
    ])
if tiene_permiso("ver_auditoria"):
    opciones_menu.append("9. Auditoría del Sistema")
if tiene_permiso("ver_finanzas"):
    opciones_menu.append("10. Facturas (OCR) y Anomalías")
if tiene_permiso("gestionar_equipo"):
    opciones_menu.append("11. Gestión de Equipo")
if not opciones_menu:
    opciones_menu.append("Acceso restringido")
menu = st.sidebar.selectbox("Sección", opciones_menu)
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
if menu == "1. Inicio":
    st.markdown('<p class="main-header">Portal de Control Financiero</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Plataforma centralizada para la administración y supervisión de recursos</p>', unsafe_allow_html=True)
    st.markdown("---")

    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("### 🎯 Objetivo del Sistema")
        st.write("Control transparente de ingresos y gastos, revisión de comprobantes, auditoría y seguimiento del presupuesto compartido por tu empresa.")
    with col2:
        st.success("✅ **Estado del Sistema:** Operativo y Guardado en Nube.")

    st.markdown("---")
    st.markdown("### 👥 Equipo")
    equipo_actual = obtener_lista_responsables()
    equipo_data = [{"N.°": i + 1, "Nombre": nombre} for i, nombre in enumerate(equipo_actual)]
    mostrar_tabla_adaptativa(pd.DataFrame(equipo_data), hide_index=True)
    if len(equipo_actual) >= 100:
        st.caption("Se muestra una vista resumida del equipo. La gestión completa está disponible en la sección de integrantes.")
    if tiene_permiso("gestionar_equipo"):
        st.caption("Gestiona rangos e invita nuevas personas desde '11. Gestión de Equipo' en el menú lateral.")

elif menu == "2. Registro de Ingresos":
    st.markdown('<p class="main-header">Registro de Ingresos</p>', unsafe_allow_html=True)
    st.markdown("---")

    if tiene_permiso("registrar"):
        with st.expander("➕ Agregar Nuevo Ingreso", expanded=True):
            with st.form(f"form_nuevo_ingreso_{get_institucion_id()}"):
                c1, c2 = st.columns(2)
                with c1:
                    f_ing = st.date_input("Fecha", value=dt_module.date.today())
                    con_ing = st.text_input("Concepto")
                with c2:
                    resp_ing = st.text_input("Responsable", value=obtener_nombre_usuario_actual())
                    val_ing = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
                obs_ing = st.text_area("Observaciones (Opcional)")

                if st.form_submit_button("Guardar Ingreso"):
                    if con_ing.strip() == "" or val_ing <= 0:
                        st.error("Ingresa un concepto y un valor mayor que cero.")
                    else:
                        reg_id = generar_id_registro("ING")
                        nuevo_reg = {
                            "ID": reg_id, "Fecha": f_ing.strftime("%Y-%m-%d"),
                            "Concepto": con_ing.strip(), "Valor": float(val_ing),
                            "Responsable": resp_ing, "Observaciones": obs_ing,
                        }
                        if guardar_registro_nube('ingresos', nuevo_reg):
                            registrar_auditoria("Registró Ingreso", f"{con_ing} - ${val_ing:,.0f} (Responsable: {resp_ing})")
                            cargar_datos_nube()
                            st.success("Ingreso guardado en la nube de la empresa.")
                            st.rerun()
                        else:
                            st.error("No se pudo guardar el ingreso en la nube. No se modificó el balance.")
    else:
        st.info("Tu rango tiene acceso de consulta. Para registrar ingresos, solicita el permiso al propietario.")

    df_ingresos = st.session_state.ingresos_df
    df_ingresos_periodo = filtrar_movimientos_periodo(df_ingresos, periodo_presupuesto)
    df_gastos_periodo = filtrar_movimientos_periodo(st.session_state.gastos_df, periodo_presupuesto)
    tot_ing_actual = pd.to_numeric(df_ingresos_periodo.get("Valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    tot_gas_actual = pd.to_numeric(df_gastos_periodo.get("Valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    relacion = calcular_relacion_presupuesto(tot_ing_actual, tot_gas_actual, presupuesto_tope)

    st.markdown("### 📊 Ingresos relacionados con el presupuesto")
    m1, m2, m3 = st.columns(3)
    m1.metric("Ingresos del período", f"${tot_ing_actual:,.0f} COP")
    m2.metric("Límite de gastos", f"${presupuesto_tope:,.0f} COP")
    m3.metric("Saldo neto", f"${relacion['saldo']:,.0f} COP")
    if isinstance(periodo_presupuesto, tuple) and len(periodo_presupuesto) == 2:
        st.caption(f"Indicadores del período: {periodo_presupuesto[0]} a {periodo_presupuesto[1]}. La tabla conserva el historial completo.")
    if relacion["nivel"] == "critico":
        st.error(relacion["mensaje"])
    elif relacion["nivel"] == "alerta":
        st.warning(relacion["mensaje"])
    elif relacion["nivel"] == "ok":
        st.success(relacion["mensaje"])
    else:
        st.info(relacion["mensaje"])
    if presupuesto_tope > 0:
        st.progress(min(max(relacion["porcentaje_ingresos"] / 100, 0.0), 1.0), text=f"Ingresos equivalentes al {relacion['porcentaje_ingresos']:.0f}% del presupuesto")

    if not df_ingresos.empty:
        mostrar_tabla_adaptativa(df_ingresos.drop(columns=['ID'], errors='ignore'), hide_index=True)
        if tiene_permiso("eliminar_registros"):
            st.markdown("### 🗑️ Eliminar Ingreso")
            opciones = [f"{row['Concepto']} - ${row['Valor']:,.0f} ({row['ID']})" for _, row in df_ingresos.iterrows()]
            seleccion = st.selectbox("Selecciona para eliminar:", opciones, key=f"seleccion_eliminar_ingreso_{get_institucion_id()}")
            if st.button("❌ Eliminar Ingreso", key=f"boton_eliminar_ingreso_{get_institucion_id()}"):
                idx = opciones.index(seleccion)
                fila = df_ingresos.iloc[idx]
                if eliminar_registro_nube('ingresos', fila['ID']):
                    registrar_auditoria("Eliminó Ingreso", f"{fila['Concepto']} - ${float(fila['Valor']):,.0f}")
                    cargar_datos_nube()
                    st.success("Ingreso eliminado.")
                    st.rerun()
                else:
                    st.error("No se pudo eliminar el ingreso de la nube.")

elif menu == "3. Registro de Gastos":
    st.markdown('<p class="main-header">Registro de Gastos</p>', unsafe_allow_html=True)
    st.markdown("---")

    gastos_del_periodo = filtrar_movimientos_periodo(st.session_state.gastos_df, periodo_presupuesto)
    current_total_gastos = pd.to_numeric(gastos_del_periodo.get("Valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    if presupuesto_tope > 0 and current_total_gastos > presupuesto_tope:
        st.error(f"🚨 Los gastos del período superan el límite en ${current_total_gastos - presupuesto_tope:,.0f} COP.")
    elif presupuesto_tope > 0:
        st.info(f"ℹ️ Disponible en el período: ${(presupuesto_tope - current_total_gastos):,.0f} COP.")

    if tiene_permiso("registrar"):
        with st.expander("➕ Agregar Nuevo Gasto", expanded=True):
            with st.form(f"form_nuevo_gasto_{get_institucion_id()}"):
                c1, c2 = st.columns(2)
                with c1:
                    f_gas = st.date_input("Fecha Gasto", value=dt_module.date.today())
                    con_gas = st.text_input("Concepto")
                    cat_gas = st.selectbox("Categoría", CATEGORIAS_GASTO)
                with c2:
                    val_gas = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
                    resp_gas = st.text_input("Responsable", value=obtener_nombre_usuario_actual())

                if st.form_submit_button("Guardar Gasto"):
                    if con_gas.strip() == "" or val_gas <= 0:
                        st.error("Ingresa un concepto y un valor mayor que cero.")
                    else:
                        reg_id = generar_id_registro("GAS")
                        nuevo_reg = {
                            "ID": reg_id, "Fecha": f_gas.strftime("%Y-%m-%d"),
                            "Concepto": con_gas.strip(), "Categoría": cat_gas,
                            "Valor": float(val_gas), "Responsable": resp_gas,
                        }
                        if guardar_registro_nube('gastos', nuevo_reg):
                            registrar_auditoria("Registró Gasto", f"{con_gas} - ${val_gas:,.0f} (Categoría: {cat_gas}, Responsable: {resp_gas})")
                            cargar_datos_nube()
                            st.success("Gasto guardado en la nube de la empresa.")
                            st.rerun()
                        else:
                            st.error("No se pudo guardar el gasto en la nube. No se modificó el balance.")
    else:
        st.info("Tu rango tiene acceso de consulta. Para registrar gastos, solicita el permiso al propietario.")

    if not st.session_state.gastos_df.empty:
        mostrar_tabla_adaptativa(st.session_state.gastos_df.drop(columns=['ID'], errors='ignore'), hide_index=True)
        st.metric("💸 TOTAL GASTOS", f"${current_total_gastos:,.0f} COP")

        if tiene_permiso("eliminar_registros"):
            st.markdown("### 🗑️ Eliminar Gasto")
            opciones = [f"{row['Concepto']} - ${row['Valor']:,.0f} ({row['ID']})" for _, row in st.session_state.gastos_df.iterrows()]
            seleccion = st.selectbox("Selecciona para eliminar:", opciones, key=f"seleccion_eliminar_gasto_{get_institucion_id()}")

            if st.button("❌ Eliminar Gasto", key=f"boton_eliminar_gasto_{get_institucion_id()}"):
                idx = opciones.index(seleccion)
                fila = st.session_state.gastos_df.iloc[idx]
                if eliminar_registro_nube('gastos', fila['ID']):
                    registrar_auditoria("Eliminó Gasto", f"{fila['Concepto']} - ${float(fila['Valor']):,.0f}")
                    cargar_datos_nube()
                    st.success("Gasto eliminado.")
                    st.rerun()
                else:
                    st.error("No se pudo eliminar el gasto de la nube.")
elif menu == "4. Balance Financiero":
    st.markdown('<p class="main-header">Balance Financiero General</p>', unsafe_allow_html=True)
    st.markdown("---")

    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    saldo = tot_ing - tot_gas

    c1, c2, c3 = st.columns(3)
    c1.metric("💵 Ingresos", f"${tot_ing:,.0f} COP")
    c2.metric("💸 Gastos", f"${tot_gas:,.0f} COP")
    c3.metric("💰 Saldo Neto", f"${saldo:,.0f} COP", delta=f"${saldo:,.0f} COP")

elif menu == "5. Dashboard y Gráficos":
    st.markdown('<p class="main-header">Tablero Ejecutivo</p>', unsafe_allow_html=True)
    st.markdown("---")

    tot_ing = pd.to_numeric(st.session_state.ingresos_df["Valor"], errors='coerce').sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = pd.to_numeric(st.session_state.gastos_df["Valor"], errors='coerce').sum() if not st.session_state.gastos_df.empty else 0.0

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Ingresos vs Gastos")
        df_comp = pd.DataFrame({"Tipo": ["Ingresos", "Gastos"], "Monto": [tot_ing, tot_gas]})
        fig_bar = aplicar_tema_grafico(px.bar(df_comp, x="Tipo", y="Monto", color="Tipo", text_auto=True, color_discrete_sequence=["#10B981", "#EF4444"]))
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("#### Gastos por Categoría")
        if not st.session_state.gastos_df.empty:
            df_gastos_temp = st.session_state.gastos_df.copy()
            df_gastos_temp["Valor"] = pd.to_numeric(df_gastos_temp["Valor"], errors='coerce').fillna(0)
            df_cat = df_gastos_temp.groupby("Categoría")["Valor"].sum().reset_index()
            fig_pie = aplicar_tema_grafico(px.pie(df_cat, names="Categoría", values="Valor", hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3))
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
            fig_line = aplicar_tema_grafico(px.line(df_timeline, x="Fecha", y="Valor", color="Tipo", markers=True, color_discrete_map={"Ingreso": "#10B981", "Gasto": "#EF4444"}))
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
    st.markdown('<p class="main-header" style="font-size:1.6rem;">Desviaciones Presupuestarias</p>', unsafe_allow_html=True)

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
                        step=50000.0, key=f"presupuesto_{get_institucion_id()}_{categoria}"
                    )
            if st.button("💾 Guardar presupuestos por categoría"):
                if guardar_presupuestos_categoria(nuevos_presupuestos):
                    registrar_auditoria("Actualizó presupuestos por categoría", str(nuevos_presupuestos))
                    st.success("Presupuestos actualizados.")
                    st.rerun()
                else:
                    st.error("No se pudieron guardar los presupuestos.")

    presupuestos_categoria = cargar_presupuestos_categoria()

    if not st.session_state.gastos_df.empty:
        df_gastos_cat = st.session_state.gastos_df.copy()
        df_gastos_cat["Valor"] = pd.to_numeric(df_gastos_cat["Valor"], errors='coerce').fillna(0)
        df_desviacion = calcular_tabla_desviaciones(st.session_state.gastos_df, presupuestos_categoria)

        fig_desviacion = aplicar_tema_grafico(px.bar(
            df_desviacion, x="Categoría", y=["Gasto Real", "Presupuesto Asignado"],
            barmode="group", text_auto=True,
            color_discrete_map={"Gasto Real": "#EF4444", "Presupuesto Asignado": "#94A3B8"},
            title="Gasto real vs. presupuesto asignado por categoría — haz clic en una barra para ver el detalle"
        ))

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

        mostrar_tabla_adaptativa(df_desviacion, hide_index=True, columna_alerta="Desviación")

        if categoria_seleccionada:
            st.markdown(f"#### 🔍 Detalle de transacciones — {categoria_seleccionada}")
            detalle_cat = df_gastos_cat[df_gastos_cat["Categoría"] == categoria_seleccionada].drop(columns=['ID'], errors='ignore')
            if not detalle_cat.empty:
                mostrar_tabla_adaptativa(detalle_cat, hide_index=True)
            else:
                st.info("No hay transacciones registradas en esta categoría todavía.")
    else:
        st.info("Registra gastos para ver el tablero de desviaciones presupuestarias.")

elif menu == "6. Anexo de Recibos & QR":
    st.markdown('<p class="main-header">Generador de Comprobantes</p>', unsafe_allow_html=True)
    st.markdown("---")
    st.info("💡 Haz clic para generar el comprobante oficial decorado con los datos financieros actuales y su código QR.")

    if st.button("🚀 Generar Comprobante Oficial"):
        tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
        tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
        saldo = tot_ing - tot_gas
        rec_id = generar_id_registro("GEN")
        nombre_empresa_recibo = obtener_nombre_empresa()
        texto_recibo = f"COMPROBANTE {rec_id}\nEmpresa: {nombre_empresa_recibo}\nIngresos: ${tot_ing:,.0f}\nGastos: ${tot_gas:,.0f}\nSaldo: ${saldo:,.0f}"
        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(texto_recibo)
        qr.make(fit=True)
        qr_img_pil = qr.make_image(fill_color="black", back_color="white").convert("RGB")

        rec_id = f"GEN-{dt_module.datetime.now().strftime('%Y%m%d%H%M')}"
        fecha_actual = dt_module.datetime.now().strftime("%Y-%m-%d")
        buffer_recibo = generar_imagen_recibo(rec_id, fecha_actual, tot_ing, tot_gas, saldo, qr_img_pil, nombre_empresa_recibo)
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
    st.markdown('<p class="main-header">Repositorio de Documentos</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Registra, administra y visualiza los comprobantes del proyecto.</p>', unsafe_allow_html=True)
    st.markdown("---")

    import base64

    with st.form(f"form_subir_archivo_{get_institucion_id()}"):
        archivo_subido = st.file_uploader("Sube tu archivo (PDF, PNG, JPG)", type=["png", "jpg", "jpeg", "pdf"], disabled=not tiene_permiso("gestionar_archivos"), key=f"repositorio_upload_{get_institucion_id()}")
        descripcion_archivo = st.text_input("Descripción o Nota del Documento", disabled=not tiene_permiso("gestionar_archivos"))
        submit_archivo = st.form_submit_button("💾 Guardar y Registrar Archivo", disabled=not tiene_permiso("gestionar_archivos"))

        if submit_archivo and db and tiene_permiso("gestionar_archivos"):
            if archivo_subido is not None:
                bytes_archivo = archivo_subido.getvalue()
                base64_archivo = base64.b64encode(bytes_archivo).decode('utf-8')

                nombre_id = generar_id_registro("ARCH")
                institucion_id = get_institucion_id()
                doc_data = {
                    "ID": nombre_id,
                    "nombre": archivo_subido.name,
                    "tipo": archivo_subido.type,
                    "archivo_b64": base64_archivo,
                    "descripcion": descripcion_archivo if descripcion_archivo else "Sin descripción",
                    "fecha": dt_module.datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "subido_por": st.session_state.user_data['institucion']
                }
                db.collection("usuarios").document(institucion_id).collection("archivos").document(nombre_id).set(doc_data)
                registrar_auditoria("Subió Archivo", f"{archivo_subido.name} — {descripcion_archivo or 'Sin descripción'}")
                st.success("✅ ¡Archivo guardado exitosamente en la base de datos!")
                st.rerun()
            else:
                st.error("⚠️ Por favor selecciona un archivo antes de guardar.")

    if not tiene_permiso("gestionar_archivos"):
        st.caption("Tu rango puede consultar el repositorio, pero no guardar o eliminar archivos.")

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

                if tiene_permiso("gestionar_archivos"):
                    st.markdown("---")
                    st.markdown("### 🗑️ Eliminar Registro de Archivo")
                    opciones_arch = [f"{row['nombre']} ({row['fecha']})" for row in archivos_lista]
                    sel_arch = st.selectbox("Selecciona archivo a eliminar:", opciones_arch, key=f"seleccion_eliminar_archivo_{get_institucion_id()}")

                    if st.button("❌ Eliminar Registro", key=f"boton_eliminar_archivo_{get_institucion_id()}"):
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
    st.markdown('<p class="main-header">Reporte Financiero</p>', unsafe_allow_html=True)
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

        nombre_usuario = obtener_nombre_empresa()

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

        df_ing = st.session_state.ingresos_df.drop(columns=["Ruta_Archivo"], errors="ignore")
        df_gas = st.session_state.gastos_df.drop(columns=["Ruta_Archivo"], errors="ignore")

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

        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        st.session_state.excel_reporte_bytes = excel_buffer.getvalue()
        st.success("✅ ¡Reporte Excel generado con éxito usando el nombre de tu cuenta!")

    if "excel_reporte_bytes" in st.session_state:
        st.download_button(
            "Descargar archivo Excel",
            data=st.session_state.excel_reporte_bytes,
            file_name="Reporte_Financiero.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")
    st.markdown("### 📄 Reporte PDF Corporativo")
    st.caption("Documento formal con la identidad visual del proyecto — no incluye firma digital ni certificación legal.")

    if st.button("📄 Generar Reporte PDF Corporativo"):
        nombre_institucion_pdf = obtener_nombre_empresa()
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
    st.markdown('<p class="main-header">Auditoría y Registro de Actividad</p>', unsafe_allow_html=True)
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
                mostrar_tabla_adaptativa(df_logs[columnas_orden + otras], hide_index=True)
            else:
                st.info("🔒 Aún no hay movimientos registrados para tu institución. Las acciones (ingresos, gastos, archivos) se registrarán aquí automáticamente.")
        except Exception as e:
            st.warning(f"No se pudieron cargar los registros de auditoría: {e}")
    else:
        st.warning("Conecta Firebase para habilitar la auditoría en la nube.")

elif menu == "10. Facturas (OCR) y Anomalías":
    st.markdown('<p class="main-header">Comprobantes y revisión</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Carga varios comprobantes juntos. La IA identifica ingresos o gastos; revisa los datos antes de contabilizarlos.</p>', unsafe_allow_html=True)
    st.markdown("---")

    gemini_api_key = st.secrets.get("gemini", {}).get("api_key") if "gemini" in st.secrets else None
    archivos_comprobante = []
    if not gemini_api_key:
        st.warning("Falta configurar la clave de Gemini en los secretos de la aplicación ([gemini].api_key).")
    else:
        archivos_comprobante = st.file_uploader(
            "Selecciona comprobantes de ingresos y gastos",
            type=["png", "jpg", "jpeg", "pdf"],
            accept_multiple_files=True,
            disabled=not tiene_permiso("registrar"),
            key=f"upload_comprobantes_lote_{get_institucion_id()}",
            help="Puedes seleccionar varios archivos en una sola carga.",
        ) or []
        if archivos_comprobante and st.button("Analizar todos los comprobantes", type="primary", disabled=not tiene_permiso("registrar"), key="analizar_lote_ocr"):
            lote_id = uuid.uuid4().hex[:12]
            elementos = []
            mime_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "pdf": "application/pdf"}
            barra = st.progress(0, text="Analizando comprobantes…")
            for indice, archivo in enumerate(archivos_comprobante):
                extension = os.path.splitext(archivo.name)[1].lower().lstrip(".")
                tipo_mime = mime_map.get(extension, archivo.type or "application/octet-stream")
                contenido = archivo.getvalue()
                datos_extraidos = extraer_datos_comprobante_gemini(contenido, tipo_mime, gemini_api_key)
                if datos_extraidos:
                    elementos.append({
                        "clave": f"{lote_id}_{indice}",
                        "nombre": archivo.name,
                        "mime": tipo_mime,
                        "contenido": contenido,
                        "datos": datos_extraidos,
                    })
                barra.progress((indice + 1) / len(archivos_comprobante), text=f"Analizados {indice + 1} de {len(archivos_comprobante)}")
            st.session_state.ocr_comprobantes = elementos
            if elementos:
                st.success(f"Se analizaron {len(elementos)} comprobante(s). Confirma cada movimiento abajo.")
            else:
                st.warning("No se pudieron extraer datos. Revisa los archivos e inténtalo de nuevo.")

    elementos_ocr = st.session_state.get("ocr_comprobantes", [])
    if elementos_ocr:
        st.markdown("### Revisa los movimientos detectados")
        st.caption("Cada ingreso extraído por IA queda pendiente hasta que un usuario autorizado lo apruebe. Los gastos solo se contabilizan directamente si no tienen alertas.")
        datos_para_procesar = []
        with st.form("form_confirmar_lote_ocr"):
            for indice, item in enumerate(elementos_ocr):
                datos = item.get("datos", {})
                clave = item["clave"]
                with st.expander(f"{item['nombre']} — {str(datos.get('tipo', 'gasto')).title()}", expanded=indice == 0):
                    c1, c2 = st.columns(2)
                    with c1:
                        incluir = st.checkbox("Incluir este movimiento", value=True, key=f"ocr_{clave}_incluir")
                        tipo_movimiento = st.selectbox(
                            "Tipo de movimiento", ["Ingreso", "Gasto"],
                            index=0 if datos.get("tipo") == "ingreso" else 1,
                            key=f"ocr_{clave}_tipo",
                        )
                        concepto = st.text_input("Concepto", value=str(datos.get("concepto") or ""), key=f"ocr_{clave}_concepto")
                        try:
                            valor_inicial = float(datos.get("valor") or 0)
                            if valor_inicial != valor_inicial or valor_inicial < 0:
                                valor_inicial = 0.0
                        except (TypeError, ValueError):
                            valor_inicial = 0.0
                        valor = st.number_input("Valor en COP", min_value=0.0, value=valor_inicial, step=1000.0, key=f"ocr_{clave}_valor")
                    with c2:
                        fecha_texto = datos.get("fecha")
                        try:
                            fecha_inicial = dt_module.date.fromisoformat(str(fecha_texto)) if fecha_texto else dt_module.date.today()
                        except (TypeError, ValueError):
                            fecha_inicial = dt_module.date.today()
                        fecha = st.date_input("Fecha", value=fecha_inicial, key=f"ocr_{clave}_fecha")
                        responsable = st.text_input("Responsable", value=obtener_nombre_usuario_actual(), key=f"ocr_{clave}_responsable")
                        categoria_sugerida = datos.get("categoria_sugerida")
                        indice_categoria = CATEGORIAS_GASTO.index(categoria_sugerida) if categoria_sugerida in CATEGORIAS_GASTO else CATEGORIAS_GASTO.index("Varios")
                        categoria = st.selectbox("Categoría (para gastos)", CATEGORIAS_GASTO, index=indice_categoria, key=f"ocr_{clave}_categoria")
                    observaciones = st.text_area("Observaciones", value=str(datos.get("observaciones") or ""), key=f"ocr_{clave}_observaciones")
                    if item["mime"].startswith("image/"):
                        st.image(item["contenido"], caption="Comprobante original", width=260)
                    else:
                        st.caption(f"PDF adjunto: {item['nombre']}")
                    datos_para_procesar.append({
                        "clave": clave, "incluir": incluir, "tipo": tipo_movimiento,
                        "concepto": concepto.strip(), "valor": float(valor), "fecha": fecha,
                        "responsable": responsable, "categoria": categoria,
                        "observaciones": observaciones, "nombre": item["nombre"],
                        "mime": item["mime"], "contenido": item["contenido"],
                        "numero_factura": datos.get("numero_factura"),
                    })
            procesar_lote = st.form_submit_button("Procesar movimientos seleccionados", type="primary")

        if procesar_lote:
            if not tiene_permiso("registrar"):
                st.error("Tu rango no tiene permiso para registrar comprobantes.")
            else:
                exitosos = 0
                pendientes_ingreso = 0
                errores = []
                claves_procesadas = set()
                presupuesto_categoria_actual = cargar_presupuestos_categoria()
                ingresos_para_analisis = st.session_state.ingresos_df.copy()
                gastos_para_analisis = st.session_state.gastos_df.copy()
                for candidato in datos_para_procesar:
                    if not candidato["incluir"]:
                        continue
                    if not candidato["concepto"] or candidato["valor"] <= 0:
                        errores.append(f"{candidato['nombre']}: el concepto y un valor mayor que cero son obligatorios.")
                        continue
                    tipo = candidato["tipo"]
                    prefijo = "ING" if tipo == "Ingreso" else "GAS"
                    registro_id = f"{prefijo}-{candidato['clave']}"
                    ruta_archivo = guardar_comprobante_nube(candidato["nombre"], candidato["mime"], candidato["contenido"], registro_id)
                    if bucket and not ruta_archivo:
                        st.warning(f"No se pudo conservar en almacenamiento el original de {candidato['nombre']}; se guardarán los datos extraídos.")

                    if tipo == "Ingreso":
                        motivos = detectar_anomalias_ingreso(
                            candidato["valor"], candidato["fecha"].isoformat(), candidato["concepto"], ingresos_para_analisis
                        )
                        ingresos_para_analisis = pd.concat([ingresos_para_analisis, pd.DataFrame([{
                            "Fecha": candidato["fecha"].isoformat(), "Concepto": candidato["concepto"], "Valor": candidato["valor"]
                        }])], ignore_index=True)
                        motivos.insert(0, "Ingreso extraído por IA: requiere revisión humana antes de entrar al balance.")
                        pendiente_id = f"PEND-ING-{candidato['clave']}"
                        pendiente = {
                            "ID": pendiente_id, "Fecha": candidato["fecha"].isoformat(),
                            "Concepto": candidato["concepto"], "Valor": candidato["valor"],
                            "Responsable": candidato["responsable"], "Observaciones": candidato["observaciones"],
                            "Motivos": motivos, "Estado": "Pendiente", "Ruta_Archivo": ruta_archivo,
                            "Nombre_Archivo": candidato["nombre"],
                        }
                        try:
                            empresa_id = get_institucion_id()
                            db.collection("usuarios").document(empresa_id).collection("ingresos_pendientes").document(pendiente_id).set(pendiente)
                            registrar_auditoria("Ingreso OCR enviado a revisión", f"{candidato['concepto']} - ${candidato['valor']:,.0f}")
                            pendientes_ingreso += 1
                            exitosos += 1
                            claves_procesadas.add(candidato["clave"])
                        except Exception as e:
                            errores.append(f"{candidato['nombre']}: no se pudo guardar el ingreso pendiente ({e}).")
                    else:
                        motivos = detectar_anomalias_gasto(
                            candidato["valor"], candidato["fecha"].isoformat(), candidato["concepto"], candidato["categoria"],
                            gastos_para_analisis, presupuesto_tope, presupuesto_categoria_actual,
                        )
                        gastos_para_analisis = pd.concat([gastos_para_analisis, pd.DataFrame([{
                            "Fecha": candidato["fecha"].isoformat(), "Concepto": candidato["concepto"],
                            "Categoría": candidato["categoria"], "Valor": candidato["valor"],
                        }])], ignore_index=True)
                        registro = {
                            "ID": registro_id, "Fecha": candidato["fecha"].isoformat(),
                            "Concepto": candidato["concepto"], "Categoría": candidato["categoria"],
                            "Valor": candidato["valor"], "Responsable": candidato["responsable"],
                            "Observaciones": candidato["observaciones"], "Numero_Factura": candidato["numero_factura"],
                            "Ruta_Archivo": ruta_archivo, "Nombre_Archivo": candidato["nombre"],
                        }
                        try:
                            if motivos:
                                pendiente_id = f"PEND-GAS-{candidato['clave']}"
                                registro.update({"ID": pendiente_id, "Motivos": motivos, "Estado": "Pendiente"})
                                empresa_id = get_institucion_id()
                                db.collection("usuarios").document(empresa_id).collection("gastos_pendientes").document(pendiente_id).set(registro)
                                registrar_auditoria("Gasto OCR enviado a revisión", f"{candidato['concepto']} - ${candidato['valor']:,.0f} — {'; '.join(motivos)}")
                            elif not guardar_registro_nube("gastos", registro):
                                raise RuntimeError("no se pudo guardar en la nube")
                            else:
                                registrar_auditoria("Registró Gasto desde comprobante OCR", f"{candidato['concepto']} - ${candidato['valor']:,.0f}")
                            exitosos += 1
                            claves_procesadas.add(candidato["clave"])
                        except Exception as e:
                            errores.append(f"{candidato['nombre']}: no se pudo guardar el gasto ({e}).")

                if claves_procesadas:
                    st.session_state.ocr_comprobantes = [x for x in elementos_ocr if x["clave"] not in claves_procesadas]
                    cargar_datos_nube()
                    st.success(f"Procesados {exitosos} movimiento(s); {pendientes_ingreso} ingreso(s) esperan aprobación.")
                for error in errores:
                    st.error(error)
                if claves_procesadas and not errores:
                    st.rerun()

    st.markdown("---")
    st.markdown("### Ingresos y gastos pendientes de revisión")
    if not tiene_permiso("ver_finanzas"):
        st.info("Tu rango no tiene permiso para consultar los movimientos financieros.")
    elif db:
        empresa_id = get_institucion_id()
        try:
            gastos_pend_ref = db.collection("usuarios").document(empresa_id).collection("gastos_pendientes").where("Estado", "==", "Pendiente").stream()
            ingresos_pend_ref = db.collection("usuarios").document(empresa_id).collection("ingresos_pendientes").where("Estado", "==", "Pendiente").stream()
            pendientes_lista = ([{**p.to_dict(), "Tipo": "Gasto"} for p in gastos_pend_ref]
                                + [{**p.to_dict(), "Tipo": "Ingreso"} for p in ingresos_pend_ref])
            if not pendientes_lista:
                st.info("No hay ingresos ni gastos pendientes de revisión en esta empresa.")
            else:
                for pendiente in pendientes_lista:
                    es_gasto = pendiente["Tipo"] == "Gasto"
                    tipo_vista = "Gasto" if es_gasto else "Ingreso"
                    concepto_pend = pendiente.get("Concepto", "Sin concepto")
                    valor_pend = float(pendiente.get("Valor", 0) or 0)
                    fecha_pend = pendiente.get("Fecha", "")
                    with st.expander(f"{tipo_vista} · {concepto_pend} · ${valor_pend:,.0f} · {fecha_pend}"):
                        if es_gasto:
                            st.write(f"**Categoría:** {pendiente.get('Categoría', '—')}  |  **Responsable:** {pendiente.get('Responsable', '—')}")
                            if pendiente.get("Numero_Factura"):
                                st.write(f"**N.° de factura:** {pendiente['Numero_Factura']}")
                        else:
                            st.write(f"**Responsable:** {pendiente.get('Responsable', '—')}")
                            if pendiente.get("Observaciones"):
                                st.write(f"**Observaciones:** {pendiente['Observaciones']}")
                        if pendiente.get("Motivos"):
                            st.markdown("**Motivos de revisión**")
                            for motivo in pendiente["Motivos"]:
                                st.caption(f"• {motivo}")
                        ruta_archivo = pendiente.get("Ruta_Archivo")
                        prefijo_archivos_empresa = f"empresas/{empresa_id}/comprobantes/"
                        if ruta_archivo and ruta_archivo.startswith(prefijo_archivos_empresa) and bucket:
                            try:
                                archivo_original = bucket.blob(ruta_archivo).download_as_bytes()
                                if pendiente.get("Nombre_Archivo", "").lower().endswith((".png", ".jpg", ".jpeg")):
                                    st.image(archivo_original, caption="Comprobante original", width=300)
                                else:
                                    st.download_button("Descargar comprobante original", archivo_original, file_name=pendiente.get("Nombre_Archivo", "comprobante.pdf"), key=f"descarga_pend_{pendiente['ID']}")
                            except Exception:
                                st.caption("El archivo original no está disponible en almacenamiento.")

                        if not tiene_permiso("aprobar_pendientes"):
                            st.caption("Tu rango permite consultar, pero no aprobar. Solicita la revisión a un gerente autorizado.")
                        else:
                            col_aprobar, col_rechazar = st.columns(2)
                            coleccion_pend = "gastos_pendientes" if es_gasto else "ingresos_pendientes"
                            coleccion_final = "gastos" if es_gasto else "ingresos"
                            reg_id = f"APROB-{pendiente['ID']}"
                            with col_aprobar:
                                if st.button("Aprobar y registrar", key=f"aprobar_{pendiente['ID']}"):
                                    nuevo_reg = {
                                        "ID": reg_id, "Fecha": pendiente.get("Fecha", ""),
                                        "Concepto": pendiente.get("Concepto", ""), "Valor": valor_pend,
                                        "Responsable": pendiente.get("Responsable", ""),
                                        "Observaciones": pendiente.get("Observaciones", ""),
                                        "Ruta_Archivo": pendiente.get("Ruta_Archivo"),
                                        "Nombre_Archivo": pendiente.get("Nombre_Archivo"),
                                    }
                                    if es_gasto:
                                        nuevo_reg["Categoría"] = pendiente.get("Categoría", "Varios")
                                        nuevo_reg["Numero_Factura"] = pendiente.get("Numero_Factura")
                                    try:
                                        lote = db.batch()
                                        lote.set(db.collection("usuarios").document(empresa_id).collection(coleccion_final).document(reg_id), nuevo_reg)
                                        lote.update(db.collection("usuarios").document(empresa_id).collection(coleccion_pend).document(pendiente["ID"]), {
                                            "Estado": "Aprobado", "Aprobado_por": st.session_state.user_data.get("email"),
                                            "Fecha_revision": dt_module.datetime.now().isoformat(),
                                        })
                                        lote.commit()
                                        registrar_auditoria(f"Aprobó {tipo_vista} pendiente", f"{concepto_pend} - ${valor_pend:,.0f}")
                                        cargar_datos_nube()
                                        st.success(f"{tipo_vista} aprobado y registrado en el balance.")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"No se pudo aprobar el movimiento; sigue pendiente: {e}")
                            with col_rechazar:
                                if st.button("Rechazar", key=f"rechazar_{pendiente['ID']}"):
                                    try:
                                        db.collection("usuarios").document(empresa_id).collection(coleccion_pend).document(pendiente["ID"]).update({
                                            "Estado": "Rechazado", "Rechazado_por": st.session_state.user_data.get("email"),
                                            "Fecha_revision": dt_module.datetime.now().isoformat(),
                                        })
                                        registrar_auditoria(f"Rechazó {tipo_vista} pendiente", f"{concepto_pend} - ${valor_pend:,.0f}")
                                        st.info(f"{tipo_vista} rechazado; queda en el historial sin contabilizarse.")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"No se pudo rechazar el movimiento: {e}")
        except Exception as e:
            st.warning(f"No se pudieron cargar los pendientes: {e}")
    else:
        st.warning("Conecta Firebase para habilitar esta sección.")

elif menu == "11. Gestión de Equipo":
    st.markdown('<p class="main-header">Equipo y rangos de acceso</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Administra las personas vinculadas al correo base de esta empresa y define qué puede hacer cada rango.</p>', unsafe_allow_html=True)
    st.markdown("---")

    if not tiene_permiso("gestionar_equipo"):
        st.error("No tienes permiso para administrar el equipo de la empresa.")
    elif not db:
        st.warning("Conecta Firebase para habilitar la gestión del equipo.")
    else:
        empresa_id = get_institucion_id()
        nombre_empresa = obtener_nombre_empresa()
        st.info(f"**Empresa:** {nombre_empresa}  ·  **Correo base de la nube:** {empresa_id}")

        if tiene_permiso("gestionar_roles"):
            with st.expander("⚙️ Crear y configurar rangos", expanded=True):
                roles_actuales = obtener_configuracion_roles(forzar=True)
                roles_personalizados = [r for r in sorted(roles_actuales) if r not in ROLES_BASE and r != "Propietario"]
                opcion_edicion = st.selectbox("Rango que deseas editar", ["Crear un rango nuevo"] + roles_personalizados, key=f"selector_rango_edicion_{empresa_id}")
                valor_nombre = "" if opcion_edicion == "Crear un rango nuevo" else opcion_edicion
                permisos_iniciales = [] if not valor_nombre else roles_actuales.get(valor_nombre, [])
                with st.form(f"form_configurar_rango_{empresa_id}"):
                    nombre_rango = st.text_input("Nombre del rango", value=valor_nombre, max_chars=60)
                    permisos_rango = st.multiselect(
                        "Funciones de este rango",
                        options=list(PERMISOS_DISPONIBLES),
                        default=[p for p in permisos_iniciales if p in PERMISOS_DISPONIBLES],
                        format_func=lambda permiso: PERMISOS_DISPONIBLES[permiso],
                    )
                    guardar_rango = st.form_submit_button("Guardar rango")
                if guardar_rango:
                    nombre_limpio = nombre_rango.strip()
                    nombre_anterior = None if opcion_edicion == "Crear un rango nuevo" else opcion_edicion
                    ocupados = set(ROLES_BASE) | {"Propietario"}
                    if not nombre_limpio:
                        st.error("Escribe un nombre para el rango.")
                    elif nombre_limpio in ocupados:
                        st.error("Ese nombre está reservado para un rango base.")
                    elif nombre_limpio in roles_personalizados and nombre_limpio != nombre_anterior:
                        st.error("Ya existe un rango con ese nombre.")
                    else:
                        doc_empresa = db.collection("usuarios").document(empresa_id)
                        datos_empresa = doc_empresa.get().to_dict() or {}
                        config_guardada = datos_empresa.get("roles_config", {})
                        if not isinstance(config_guardada, dict):
                            config_guardada = {}
                        if nombre_anterior:
                            config_guardada.pop(nombre_anterior, None)
                        config_guardada[nombre_limpio] = list(permisos_rango)
                        try:
                            doc_empresa.set({"roles_config": config_guardada}, merge=True)
                            if nombre_anterior and nombre_anterior != nombre_limpio:
                                miembros = doc_empresa.collection("empleados").where("rol", "==", nombre_anterior).stream()
                                for miembro_doc in miembros:
                                    miembro = miembro_doc.to_dict() or {}
                                    correo_miembro = miembro.get("email", miembro_doc.id)
                                    miembro_doc.reference.update({"rol": nombre_limpio})
                                    cuenta_miembro_ref = db.collection("usuarios").document(correo_miembro)
                                    if cuenta_miembro_ref.get().exists:
                                        cuenta_miembro_ref.set({"rol": nombre_limpio}, merge=True)
                                    db.collection("empleados_index").document(correo_miembro).set({"rol": nombre_limpio}, merge=True)
                            st.session_state.roles_empresa_cache = None
                            registrar_auditoria("Configuró rango de empresa", nombre_limpio)
                            st.success(f"Rango '{nombre_limpio}' guardado.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"No se pudo guardar el rango: {e}")

                if roles_personalizados:
                    st.markdown("#### Eliminar un rango")
                    rango_eliminar = st.selectbox("Rango personalizado", roles_personalizados, key=f"selector_rango_eliminar_{empresa_id}")
                    if st.button("Eliminar rango", key=f"eliminar_rango_empresa_{empresa_id}"):
                        miembros_en_rango = list(db.collection("usuarios").document(empresa_id).collection("empleados").where("rol", "==", rango_eliminar).limit(1).stream())
                        if miembros_en_rango:
                            st.error("No se puede eliminar mientras haya integrantes asignados. Cambia sus rangos primero.")
                        else:
                            datos_empresa = db.collection("usuarios").document(empresa_id).get().to_dict() or {}
                            config_guardada = datos_empresa.get("roles_config", {})
                            config_guardada.pop(rango_eliminar, None)
                            db.collection("usuarios").document(empresa_id).set({"roles_config": config_guardada}, merge=True)
                            st.session_state.roles_empresa_cache = None
                            registrar_auditoria("Eliminó rango de empresa", rango_eliminar)
                            st.success("Rango eliminado.")
                            st.rerun()

        roles_asignables = [r for r in obtener_nombres_roles() if r != "Propietario"]
        if not roles_asignables:
            st.warning("Crea un rango antes de invitar integrantes.")
        else:
            with st.expander("➕ Invitar integrante por correo", expanded=True):
                with st.form(f"form_invitar_empleado_{empresa_id}"):
                    correo_nuevo = st.text_input("Correo del integrante")
                    nombre_nuevo = st.text_input("Nombre (opcional)")
                    rol_nuevo = st.selectbox("Rango", roles_asignables)
                    enviar_invitacion = st.form_submit_button("Vincular correo a la empresa")
                if enviar_invitacion:
                    correo_limpio = correo_nuevo.lower().strip()
                    if not correo_limpio or "@" not in correo_limpio or "." not in correo_limpio.split("@")[-1]:
                        st.error("Ingresa un correo electrónico válido.")
                    elif correo_limpio == empresa_id:
                        st.error("Ese correo ya es el correo base de la empresa.")
                    else:
                        try:
                            indice_ref = db.collection("empleados_index").document(correo_limpio)
                            indice_doc = indice_ref.get()
                            indice_actual = indice_doc.to_dict() or {}
                            cuenta_doc = db.collection("usuarios").document(correo_limpio).get()
                            cuenta = cuenta_doc.to_dict() or {}
                            empresa_cuenta = (cuenta.get("empresa_id") or cuenta.get("email") or "").lower().strip()
                            if indice_doc.exists and indice_actual.get("empresa_id") != empresa_id:
                                st.error("Ese correo ya está vinculado a otra empresa.")
                            elif cuenta_doc.exists and empresa_cuenta != empresa_id:
                                st.error("Ese correo ya tiene una cuenta personal o pertenece a otra empresa.")
                            else:
                                empleado_doc = {
                                    "email": correo_limpio,
                                    "nombre": nombre_nuevo.strip() or cuenta.get("institucion") or correo_limpio,
                                    "rol": rol_nuevo,
                                    "estado": "Activo",
                                    "fecha_invitacion": dt_module.datetime.now().isoformat(),
                                }
                                db.collection("usuarios").document(empresa_id).collection("empleados").document(correo_limpio).set(empleado_doc, merge=True)
                                indice_ref.set({"empresa_id": empresa_id, "rol": rol_nuevo, "nombre_empresa": nombre_empresa}, merge=True)
                                if cuenta_doc.exists:
                                    db.collection("usuarios").document(correo_limpio).set({"empresa_id": empresa_id, "rol": rol_nuevo}, merge=True)
                                registrar_auditoria("Vinculó integrante a empresa", f"{correo_limpio} — {rol_nuevo}")
                                st.success(f"{correo_limpio} quedó vinculado como {rol_nuevo}. Si aún no tiene cuenta, debe crearla con ese mismo correo.")
                                st.rerun()
                        except Exception as e:
                            st.error(f"No se pudo guardar la invitación: {e}")

        st.markdown("### Integrantes vinculados")
        if "equipo_mostrar_hasta" not in st.session_state:
            st.session_state.equipo_mostrar_hasta = 100
        limite_miembros = st.session_state.equipo_mostrar_hasta
        miembros = []
        try:
            miembros = [m.to_dict() or {} for m in db.collection("usuarios").document(empresa_id).collection("empleados").order_by("email").limit(limite_miembros).stream()]
            if not miembros:
                st.info("Aún no hay integrantes vinculados a esta nube empresarial.")
            for miembro in miembros:
                correo_miembro = miembro.get("email", "")
                nombre_miembro = miembro.get("nombre") or correo_miembro
                estado = miembro.get("estado", "Activo")
                with st.expander(f"{nombre_miembro} · {miembro.get('rol', 'Analista')} · {estado}"):
                    st.write(f"**Correo:** {correo_miembro}")
                    if estado == "Revocado":
                        st.caption("Acceso revocado.")
                        continue
                    rango_actual = miembro.get("rol", "Analista")
                    if not roles_asignables:
                        st.caption("No hay rangos disponibles para asignar.")
                        continue
                    indice_rango = roles_asignables.index(rango_actual) if rango_actual in roles_asignables else 0
                    rango_nuevo = st.selectbox("Rango", roles_asignables, index=indice_rango, key=f"rol_{correo_miembro}")
                    col_guardar, col_revocar = st.columns(2)
                    with col_guardar:
                        if st.button("Actualizar rango", key=f"actualizar_{correo_miembro}"):
                            try:
                                db.collection("usuarios").document(empresa_id).collection("empleados").document(correo_miembro).update({"rol": rango_nuevo})
                                db.collection("empleados_index").document(correo_miembro).set({"empresa_id": empresa_id, "rol": rango_nuevo, "nombre_empresa": nombre_empresa}, merge=True)
                                cuenta_miembro_ref = db.collection("usuarios").document(correo_miembro)
                                if cuenta_miembro_ref.get().exists:
                                    cuenta_miembro_ref.set({"rol": rango_nuevo, "empresa_id": empresa_id}, merge=True)
                                registrar_auditoria("Actualizó rango de integrante", f"{correo_miembro} → {rango_nuevo}")
                                st.success("Rango actualizado.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"No se pudo actualizar el rango: {e}")
                    with col_revocar:
                        if st.button("Revocar acceso", key=f"revocar_{correo_miembro}"):
                            try:
                                db.collection("usuarios").document(empresa_id).collection("empleados").document(correo_miembro).update({"estado": "Revocado"})
                                db.collection("empleados_index").document(correo_miembro).delete()
                                registrar_auditoria("Revocó acceso de integrante", correo_miembro)
                                st.success("Acceso revocado. La app lo bloqueará en la siguiente interacción.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"No se pudo revocar el acceso: {e}")
        except Exception as e:
            st.warning(f"No se pudo cargar el equipo: {e}")
        if len(miembros) == limite_miembros and st.button("Cargar más integrantes", key="cargar_mas_integrantes"):
            st.session_state.equipo_mostrar_hasta += 100
            st.rerun()

elif menu == "Acceso restringido":
    st.markdown('<p class="main-header">Acceso restringido</p>', unsafe_allow_html=True)
    st.info("Tu rango aún no tiene funciones asignadas. Contacta al propietario de la empresa.")
