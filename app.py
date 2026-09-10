import streamlit as st
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
            # Reemplaza caracteres escapados por saltos de línea reales
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
    "Jhonnattan Andrei Melo Salas",
    "Nicol Stefani Vanegas Cruz",
    "Luis Alejandro Martínez Rubio",
    "Iván Santiago Valencia Villamil"
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
        # Cargar Ingresos
        ing_docs = db.collection('ingresos').stream()
        ing_data = [doc.to_dict() for doc in ing_docs]
        if ing_data:
            st.session_state.ingresos_df = pd.DataFrame(ing_data)
        
        # Cargar Gastos
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
        page = doc.load_page(0)  # Primera página
        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5)) # Reducir resolución para miniatura
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=70)
        return buf.getvalue()
    except Exception as e:
        return None

# --- PANTALLAS DE AUTENTICACIÓN ---
if not st.session_state.logged_in:
    st.markdown(
        '<p class="main-header" style="text-align: center;">🏛️ Portal'
        " Financiero Institucional</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="sub-header" style="text-align: center;">Colegio Francisco'
        " de Paula Santander</p>",
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3 = st.tabs(
        ["Iniciar Sesión", "Crear Cuenta", "Olvidé mi Contraseña"]
    )

    with tab1:
        st.markdown("### Acceso Institucional con Google")
        st.write(
            "Usa tu cuenta autorizada para acceder de forma segura sin"
            " contraseñas."
        )

        auth_html = """
        <!DOCTYPE html>
        <html>
        <head>
            <script type="module">
                import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
                import { getAuth, GoogleAuthProvider, signInWithPopup } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";

                const firebaseConfig = {
                    apiKey: "TU_API_KEY",
                    authDomain: "proyecto-app-ffdb5.firebaseapp.com",
                    projectId: "proyecto-app-ffdb5",
                    storageBucket: "proyecto-app-ffdb5.appspot.com",
                    messagingSenderId: "TU_MESSAGING_SENDER_ID",
                    appId: "TU_APP_ID"
                };

                const app = initializeApp(firebaseConfig);
                const auth = getAuth(app);
                const provider = new GoogleAuthProvider();

                window.loginWithGoogle = function() {
                    signInWithPopup(auth, provider)
                        .then((result) => {
                            const user = result.user;
                            window.parent.postMessage({
                                type: 'streamlit:setComponentValue',
                                value: user.email
                            }, '*');
                        })
                        .catch((error) => {
                            console.error("Error en el login:", error);
                        });
                };
            </script>
        </head>
        <body>
            <button onclick="loginWithGoogle()" style="
                background-color: #4285F4;
                color: white;
                border: none;
                padding: 12px 20px;
                font-size: 16px;
                border-radius: 8px;
                cursor: pointer;
                font-family: sans-serif;
                font-weight: 600;
                width: 100%;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
            ">
                Acceder con Google
            </button>
        </body>
        </html>
        """

        resultado_auth = html(auth_html, height=60)

        if resultado_auth:
            email_ingresado = resultado_auth.lower()
            user_ref = db.collection("usuarios").document(email_ingresado)
            user_doc = user_ref.get()

            if user_doc.exists:
                st.session_state.logged_in = True
                st.session_state.user_data = user_doc.to_dict()
                cargar_datos_nube()
                st.success("¡Bienvenido/a!")
                st.rerun()
            else:
                nuevo_usuario = {
                    "institucion": email_ingresado.split("@")[0].capitalize(),
                    "email": email_ingresado,
                    "password": "GOOGLE_AUTH",
                    "fecha_creacion": datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                }
                db.collection("usuarios").document(email_ingresado).set(
                    nuevo_usuario
                )
                st.session_state.logged_in = True
                st.session_state.user_data = nuevo_usuario
                cargar_datos_nube()
                st.success("¡Cuenta registrada e iniciada con Google!")
                st.rerun()

        st.markdown("---")
        with st.expander("O usar contraseña tradicional"):
            with st.form("login_form_tradicional"):
                email_login = st.text_input("Correo Electrónico")
                pass_login = st.text_input("Contraseña", type="password")
                submit_login = st.form_submit_button("Entrar con Contraseña")

                if submit_login and db:
                    user_ref = db.collection("usuarios").document(
                        email_login.lower()
                    )
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
                        st.error("No existe una cuenta con este correo.")

    with tab2:
        with st.form("register_form"):
            inst_name = st.text_input("Nombre de la Institución / Persona")
            email_reg = st.text_input("Correo Electrónico")
            pass_reg = st.text_input(
                "Contraseña (Min. 6 caracteres, 1 mayúscula)", type="password"
            )
            submit_reg = st.form_submit_button("Registrar Cuenta")

            if submit_reg and db:
                if len(pass_reg) < 6 or not any(
                    c.isupper() for c in pass_reg
                ):
                    st.error(
                        "La contraseña debe tener al menos 6 caracteres y 1"
                        " letra mayúscula."
                    )
                elif not inst_name or not email_reg:
                    st.error("Todos los campos son obligatorios.")
                else:
                    email_exists = (
                        db.collection("usuarios")
                        .document(email_reg.lower())
                        .get()
                        .exists
                    )
                    name_query = (
                        db.collection("usuarios")
                        .where("institucion", "==", inst_name)
                        .get()
                    )

                    if email_exists:
                        st.error(
                            "Ya existe una cuenta con este correo electrónico."
                        )
                    elif len(name_query) > 0:
                        st.error(
                            "Ya existe una cuenta con este nombre de"
                            " institución/persona."
                        )
                    else:
                        nuevo_usuario = {
                            "institucion": inst_name,
                            "email": email_reg.lower(),
                            "password": hash_password(pass_reg),
                            "fecha_creacion": datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                        }
                        db.collection("usuarios").document(
                            email_reg.lower()
                        ).set(nuevo_usuario)
                        st.success(
                            "¡Cuenta creada exitosamente! Ya puedes iniciar"
                            " sesión."
                        )

    with tab3:
        with st.form("forgot_form"):
            st.info(
                "Ingresa tu correo y te enviaremos las instrucciones de"
                " recuperación."
            )
            email_forgot = st.text_input("Correo Electrónico registrado")
            submit_forgot = st.form_submit_button("Recuperar Contraseña")

            if submit_forgot and db:
                if (
                    db.collection("usuarios")
                    .document(email_forgot.lower())
                    .get()
                    .exists
                ):
                    st.success(
                        f"✅ Se ha enviado un correo con instrucciones de"
                        f" recuperación a {email_forgot}. (Simulación de"
                        " sistema)"
                    )
                else:
                    st.error(
                        "El correo no está registrado en nuestra base de"
                        " datos."
                    )

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
    "8. Reporte Final"
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
    st.info("💡 Haz clic para generar el comprobante con código QR.")

    if st.button("🚀 Generar Comprobante"):
        tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
        tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
        saldo = tot_ing - tot_gas
        rec_id = f"GEN-{datetime.now().strftime('%Y%m%d%H%M')}"
        
        texto_recibo = f"COMPROBANTE {rec_id}\nIngresos: ${tot_ing}\nGastos: ${tot_gas}\nSaldo: ${saldo}"
        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(texto_recibo)
        qr.make(fit=True)
        
        img_w, img_h = 650, 880
        base_img = Image.new("RGB", (img_w, img_h), color="#FFFFFF")
        draw = ImageDraw.Draw(base_img)
        try: font_title = ImageFont.truetype("arial.ttf", 22)
        except: font_title = ImageFont.load_default()
        
        draw.rectangle([(0, 0), (img_w, 110)], fill="#1E3A8A")
        draw.text((30, 40), f"Balance: {rec_id} - {datetime.now().strftime('%Y-%m-%d')}", fill="#FFFFFF", font=font_title)
        
        qr_resized = qr.make_image(fill_color="black", back_color="white").convert("RGB").resize((200, 200))
        base_img.paste(qr_resized, (225, 300))
        
        buf = BytesIO()
        base_img.save(buf, format="PNG")
        st.session_state.rec_img_bytes = buf.getvalue()
        
    if 'rec_img_bytes' in st.session_state:
        st.image(st.session_state.rec_img_bytes, width=400)
        st.download_button("📥 Descargar Recibo PNG", data=st.session_state.rec_img_bytes, file_name="Comprobante.png", mime="image/png")

elif menu == "7. Gestión de Archivos":
    st.markdown('<p class="main-header">📁 Repositorio de Documentos e Imágenes</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Sube tus recibos, facturas (PDF o Imagen) para guardarlos en la nube de forma permanente.</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    archivo_subido = st.file_uploader("Sube tu archivo (PDF, PNG, JPG)", type=["png", "jpg", "jpeg", "pdf"])
    
    if archivo_subido and db and bucket:
        if st.button("⬆️ Guardar Archivo en la Nube"):
            with st.spinner("Subiendo archivo..."):
                nombre_archivo = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{archivo_subido.name}"
                blob = bucket.blob(f"documentos/{nombre_archivo}")
                
                file_bytes = archivo_subido.getvalue()
                blob.upload_from_string(file_bytes, content_type=archivo_subido.type)
                blob.make_public()
                
                url_miniatura = blob.public_url
                es_pdf = archivo_subido.type == "application/pdf"
                
                if es_pdf:
                    miniatura_bytes = generar_miniatura_pdf(file_bytes)
                    if miniatura_bytes:
                        blob_min = bucket.blob(f"miniaturas/{nombre_archivo}.jpg")
                        blob_min.upload_from_string(miniatura_bytes, content_type="image/jpeg")
                        blob_min.make_public()
                        url_miniatura = blob_min.public_url
                
                doc_data = {
                    "nombre": archivo_subido.name,
                    "url_archivo": blob.public_url,
                    "url_miniatura": url_miniatura,
                    "tipo": "PDF" if es_pdf else "Imagen",
                    "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "subido_por": st.session_state.user_data['institucion']
                }
                db.collection("archivos").document(nombre_archivo).set(doc_data)
                st.success("✅ Archivo subido y guardado exitosamente.")
            
    st.markdown("### 🖼️ Galería de Archivos Guardados")
    if db:
        archivos_ref = db.collection("archivos").stream()
        archivos_lista = [a.to_dict() for a in archivos_ref]
        
        if archivos_lista:
            columnas = st.columns(4)
            for i, arch in enumerate(archivos_lista):
                with columnas[i % 4]:
                    st.markdown('<div class="file-card">', unsafe_allow_html=True)
                    st.image(arch['url_miniatura'], use_container_width=True)
                    st.markdown(f"**{arch['nombre'][:15]}...**")
                    st.caption(f"{arch['tipo']} | {arch['fecha']}")
                    st.markdown(f"[📥 Descargar original]({arch['url_archivo']})")
                    st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("Aún no hay archivos subidos en la nube.")
    else:
        st.warning("Conecta Firebase para ver la galería de archivos.")

elif menu == "8. Reporte Final":
    st.markdown('<p class="main-header">📑 Descarga de Excel</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        st.session_state.ingresos_df.drop(columns=['ID'], errors='ignore').to_excel(writer, sheet_name='Ingresos', index=False)
        st.session_state.gastos_df.drop(columns=['ID'], errors='ignore').to_excel(writer, sheet_name='Gastos', index=False)
    
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("⬇️ Descargar Backup Excel", data=f, file_name="Proyecto_Financiero_Cloud.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
