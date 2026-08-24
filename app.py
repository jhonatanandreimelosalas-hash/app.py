import streamlit as st
import pandas as pd
import openpyxl
import qrcode
from io import BytesIO
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from pdf2image import convert_from_bytes

st.set_page_config(page_title="Gestión Financiera - Prototipo Eventos", page_icon="💰", layout="wide")

EXCEL_FILE = "Proyecto_Financiero_Eventos_Actualizado (1).xlsx"

# Lista oficial de integrantes para menús desplegables
INTEGRANTES_LISTA = [
    "Ivan Santiago Valencia Villamil",
    "Nicol Vanegas Cruz",
    "Jhonatan Andrey Melo",
    "Alejandro Martinez Rubio"
]

@st.cache_data
def load_excel_data():
    xls = pd.ExcelFile(EXCEL_FILE)
    sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
    return sheets

def save_excel_data(df_dict):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    # Limpiar caché para forzar la lectura del Excel actualizado
    st.cache_data.clear()

# Función para generar un recibo profesional en formato PDF utilizando ReportLab
def generar_pdf_recibo(rec_id, fecha, tipo, concepto, valor, responsable):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    # Encabezado y Estilo
    c.setFillColorRGB(0.1, 0.2, 0.4)
    c.rect(0, height - 100, width, 100, fill=1, stroke=0)
    
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(50, height - 45, "COMPROBANTE OFICIAL DE EVENTO")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 70, "Sistema de Control y Gestión Financiera")
    
    # Detalles del Recibo
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 12)
    y_pos = height - 150
    
    c.drawString(50, y_pos, f"ID de Comprobante: {rec_id}")
    y_pos -= 30
    c.drawString(50, y_pos, f"Fecha de Emisión: {fecha}")
    y_pos -= 30
    c.drawString(50, y_pos, f"Tipo de Movimiento: {tipo}")
    y_pos -= 30
    c.drawString(50, y_pos, f"Concepto: {concepto}")
    y_pos -= 30
    c.drawString(50, y_pos, f"Valor Total: ${valor:,.0f} COP")
    y_pos -= 30
    c.drawString(50, y_pos, f"Responsable Emisor: {responsable}")
    y_pos -= 30
    c.drawString(50, y_pos, "Estado: Registrado, Verificado y Aprobado")
    
    # Pie de página
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.line(50, 100, width - 50, 100)
    c.setFont("Helvetica-Oblique", 9)
    c.setFillColorRGB(0.4, 0.4, 0.4)
    c.drawString(50, 80, "Este documento es un comprobante digital generado automáticamente por el sistema del proyecto.")
    
    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# Cargar datos
try:
    data_dict = load_excel_data()
except Exception as e:
    st.error(f"Error al cargar el archivo Excel: {e}")
    st.stop()

st.sidebar.title("📌 Menú de Navegación")
menu = st.sidebar.selectbox("Selecciona una sección:", [
    "1. Inicio", 
    "2. Registro de Ingresos", 
    "3. Registro de Gastos", 
    "4. Balance Financiero", 
    "5. Dashboard y Gráficos", 
    "6. Anexo de Recibos & QR", 
    "7. Reporte Final"
])

# --- 1. INICIO ---
if menu == "1. Inicio":
    st.title("🏛️ Proyecto de Control y Gestión Financiera")
    st.subheader("Sistema Genérico para Control de Eventos (Ingresos, Gastos y Balance)")
    st.markdown("---")
    st.markdown("### 📋 1. Información General del Proyecto")
    st.info("**Nombre del Proyecto:** Plantilla Estándar de Presupuesto y Balance de Eventos")
    st.info("**Objetivo:** Registrar recaudos y gastos para calcular la ganancia o saldo final.")
    st.info("**Versión:** v2.2 (Interactiva y Sincronizada)")
    st.info("**Estado:** En Desarrollo / Proyecto Base")
    
    st.markdown("### 👥 2. Equipo / Integrantes")
    integrantes_data = [
        {"N.°": 1, "Nombre Completo": "Ivan Santiago Valencia Villamil", "Rol / Responsabilidad": "Líder de Proyecto / Administración"},
        {"N.°": 2, "Nombre Completo": "Nicol Vanegas Cruz", "Rol / Responsabilidad": "Gestión de Registro e Ingresos"},
        {"N.°": 3, "Nombre Completo": "Jhonatan Andrey Melo", "Rol / Responsabilidad": "Control de Gastos e Insumos"},
        {"N.°": 4, "Nombre Completo": "Alejandro Martinez Rubio", "Rol / Responsabilidad": "Soportes y Control de Balance"},
    ]
    st.table(pd.DataFrame(integrantes_data))

# --- 2. REGISTRO DE INGRESOS ---
elif menu == "2. Registro de Ingresos":
    st.title("📈 Registro de Ingresos")
    df_ingresos = data_dict['Registro de Ingresos']
    
    st.markdown("### 📝 Tabla Interactiva (Modifica datos o elimina filas y guarda)")
    
    column_config = {}
    if 'Responsable' in df_ingresos.columns:
        column_config['Responsable'] = st.column_config.SelectboxColumn(
            "Responsable",
            options=INTEGRANTES_LISTA,
            required=True
        )

    edited_ingresos = st.data_editor(
        df_ingresos, 
        use_container_width=True, 
        num_rows="dynamic", 
        key="editor_ingresos",
        column_config=column_config
    )
    
    if st.button("💾 Guardar cambios en la tabla de Ingresos"):
        data_dict['Registro de Ingresos'] = edited_ingresos
        save_excel_data(data_dict)
        st.success("¡Cambios guardados y sincronizados correctamente!")
        st.rerun()
    
    st.markdown("---")
    st.markdown("### ➕ Agregar Nuevo Ingreso")
    with st.form("form_ingreso"):
        col1, col2 = st.columns(2)
        with col1:
            fecha = st.date_input("Fecha")
            concepto = st.text_input("Concepto (ej. Venta de boletas)")
            valor = st.number_input("Valor ($)", min_value=0.0, step=1000.0)
        with col2:
            responsable = st.selectbox("Responsable", INTEGRANTES_LISTA, index=1)
            observaciones = st.text_area("Observaciones")
        
        submitted = st.form_submit_button("Guardar Ingreso")
        if submitted:
            nuevo_reg = pd.DataFrame({
                'Fecha': [str(fecha)],
                'Concepto': [concepto],
                'Valor': [valor],
                'Responsable': [responsable],
                'Observaciones': [observaciones]
            })
            data_dict['Registro de Ingresos'] = pd.concat([df_ingresos, nuevo_reg], ignore_index=True)
            save_excel_data(data_dict)
            st.success("¡Ingreso agregado y guardado correctamente!")
            st.rerun()

# --- 3. REGISTRO DE GASTOS ---
elif menu == "3. Registro de Gastos":
    st.title("📉 Registro de Gastos")
    df_gastos = data_dict['Registro de Gastos']
    
    st.markdown("### 📝 Tabla Interactiva (Modifica datos o elimina filas y guarda)")
    
    column_config_g = {}
    if 'Responsable' in df_gastos.columns:
        column_config_g['Responsable'] = st.column_config.SelectboxColumn(
            "Responsable",
            options=INTEGRANTES_LISTA,
            required=True
        )
    if 'Categoría' in df_gastos.columns:
        column_config_g['Categoría'] = st.column_config.SelectboxColumn(
            "Categoría",
            options=["Logística", "Publicidad", "Alimentación", "Varios"],
            required=True
        )

    edited_gastos = st.data_editor(
        df_gastos, 
        use_container_width=True, 
        num_rows="dynamic", 
        key="editor_gastos",
        column_config=column_config_g
    )
    
    if st.button("💾 Guardar cambios en la tabla de Gastos"):
        data_dict['Registro de Gastos'] = edited_gastos
        save_excel_data(data_dict)
        st.success("¡Cambios de gastos guardados y sincronizados correctamente!")
        st.rerun()
        
    st.markdown("---")
    st.markdown("### ➕ Agregar Nuevo Gasto")
    with st.form("form_gasto"):
        col1, col2 = st.columns(2)
        with col1:
            fecha = st.date_input("Fecha Gasto")
            concepto = st.text_input("Concepto (ej. Alquiler de sonido)")
            categoria = st.selectbox("Categoría", ["Logística", "Publicidad", "Alimentación", "Varios"])
        with col2:
            valor = st.number_input("Valor ($)", min_value=0.0, step=1000.0)
            responsable = st.selectbox("Responsable", INTEGRANTES_LISTA, index=2)
            observaciones = st.text_area("Observaciones")
        
        submitted = st.form_submit_button("Guardar Gasto")
        if submitted:
            nuevo_reg = pd.DataFrame({
                'Fecha': [str(fecha)],
                'Concepto': [concepto],
                'Categoría': [categoria],
                'Valor': [valor],
                'Responsable': [responsable],
                'Observaciones': [observaciones]
            })
            data_dict['Registro de Gastos'] = pd.concat([df_gastos, nuevo_reg], ignore_index=True)
            save_excel_data(data_dict)
            st.success("¡Gasto agregado y guardado correctamente!")
            st.rerun()

# --- 4. BALANCE FINANCIERO ---
elif menu == "4. Balance Financiero":
    st.title("⚖️ Balance Financiero General")
    
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    total_ingresos = pd.to_numeric(df_ing['Valor'], errors='coerce').sum() if 'Valor' in df_ing.columns else 0
    total_gastos = pd.to_numeric(df_gas['Valor'], errors='coerce').sum() if 'Valor' in df_gas.columns else 0
    saldo = total_ingresos - total_gastos
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Ingresos", f"${total_ingresos:,.0f}")
    col2.metric("Total Gastos", f"${total_gastos:,.0f}")
    col3.metric("Saldo / Ganancia Neta", f"${saldo:,.0f}", delta=f"${saldo:,.0f}")
    
    st.markdown("---")
    balance_df = pd.DataFrame({
        "Concepto": ["Total de Ingresos", "Total de Gastos", "Saldo Final"],
        "Valor ($)": [total_ingresos, total_gastos, saldo]
    })
    st.table(balance_df)

# --- 5. DASHBOARD ---
elif menu == "5. Dashboard y Gráficos":
    st.title("📊 Dashboard y Resumen Visual")
    
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    total_ingresos = pd.to_numeric(df_ing['Valor'], errors='coerce').sum() if 'Valor' in df_ing.columns else 0
    total_gastos = pd.to_numeric(df_gas['Valor'], errors='coerce').sum() if 'Valor' in df_gas.columns else 0
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Comparativa Ingresos vs Gastos")
        chart_data = pd.DataFrame({
            "Tipo": ["Ingresos", "Gastos"],
            "Monto": [total_ingresos, total_gastos]
        }).set_index("Tipo")
        st.bar_chart(chart_data)
        
    with col2:
        st.subheader("Gastos por Categoría")
        if 'Categoría' in df_gas.columns and 'Valor' in df_gas.columns:
            df_gas_copy = df_gas.copy()
            df_gas_copy['Valor'] = pd.to_numeric(df_gas_copy['Valor'], errors='coerce')
            cat_grouped = df_gas_copy.groupby('Categoría')['Valor'].sum()
            st.bar_chart(cat_grouped)
        else:
            st.info("No hay suficientes datos de categorías.")

# --- 6. ANEXO DE RECIBOS & QR ---
elif menu == "6. Anexo de Recibos & QR":
    st.title("🧾 Generador Automático de Recibos PDF y Códigos QR")
    st.markdown("Crea un recibo oficial en formato **PDF** y su respectivo código QR estructurado para escaneo móvil.")
    
    df_anexo = data_dict['Anexo de recibos']
    st.dataframe(df_anexo, use_container_width=True)
    
    st.markdown("### ➕ Crear Recibo en PDF y QR")
    
    if 'ultimo_pdf_bytes' not in st.session_state:
        st.session_state.ultimo_pdf_bytes = None
    if 'ultimo_recibo_id' not in st.session_state:
        st.session_state.ultimo_recibo_id = None
    if 'ultimo_qr_img' not in st.session_state:
        st.session_state.ultimo_qr_img = None

    with st.form("form_qr_auto"):
        col1, col2 = st.columns(2)
        with col1:
            rec_id = st.text_input("ID Recibo", value=f"REC-00{len(df_anexo)+1}")
            fecha = st.date_input("Fecha")
            tipo = st.selectbox("Tipo de Movimiento", ["Ingreso", "Gasto"])
            concepto = st.text_input("Concepto (ej. Pago de sonido / Venta boleta)")
        with col2:
            valor = st.number_input("Valor ($)", min_value=0.0, step=1000.0)
            responsable = st.selectbox("Responsable que emite", INTEGRANTES_LISTA)
            
        submitted = st.form_submit_button("Generar Recibo PDF y QR")
        if submitted:
            # Generar el PDF real
            pdf_bytes = generar_pdf_recibo(rec_id, str(fecha), tipo, concepto, valor, responsable)
            
            # Formato optimizado para el QR (Estructura clara de datos del comprobante)
            texto_qr = f"COMPROBANTE_ID:{rec_id}|TIPO:{tipo}|CONCEPTO:{concepto}|VALOR:${valor:,.0f}COP|RESPONSABLE:{responsable}|ESTADO:VERIFICADO"
            
            # Configuración del QR
            qr = qrcode.QRCode(
                version=None,
                error_correction=qrcode.constants.ERROR_CORRECT_M,
                box_size=8,
                border=2,
            )
            qr.add_data(texto_qr)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            
            buf_qr = BytesIO()
            img.save(buf_qr, format="PNG")
            
            nuevo_anexo = pd.DataFrame({
                'ID': [rec_id],
                'Fecha': [str(fecha)],
                'Tipo (Ingreso/Gasto)': [tipo],
                'Concepto': [concepto],
                'Valor': [valor],
                'Nombre del archivo o enlace': [f"Recibo_PDF_{rec_id}.pdf"],
                'Código QR': [f"QR Oficial - {concepto}"]
            })
            data_dict['Anexo de recibos'] = pd.concat([df_anexo, nuevo_anexo], ignore_index=True)
            save_excel_data(data_dict)
            
            # Guardar en session state
            st.session_state.ultimo_pdf_bytes = pdf_bytes
            st.session_state.ultimo_recibo_id = rec_id
            st.session_state.ultimo_qr_img = buf_qr.getvalue()
            
            st.success("¡Recibo PDF y Código QR generados con éxito!")

    # Vista previa y botones de descarga seguros fuera del formulario
    if st.session_state.ultimo_pdf_bytes is not None:
        st.markdown("---")
        st.subheader(f"📄 Comprobante Generado: {st.session_state.ultimo_recibo_id}")
        
        col_prev1, col_prev2 = st.columns(2)
        with col_prev1:
            st.image(st.session_state.ultimo_qr_img, caption=f"Código QR Oficial - {st.session_state.ultimo_recibo_id}", width=200)
        with col_prev2:
            st.info("El código QR codifica toda la información oficial del comprobante financiero.")
            st.download_button(
                label="📥 Descargar Comprobante en formato PDF",
                data=st.session_state.ultimo_pdf_bytes,
                file_name=f"{st.session_state.ultimo_recibo_id}_comprobante.pdf",
                mime="application/pdf"
            )

# --- 7. REPORTE FINAL ---
elif menu == "7. Reporte Final":
    st.title("📑 Reporte Final del Evento")
    
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    total_ingresos = pd.to_numeric(df_ing['Valor'], errors='coerce').sum() if 'Valor' in df_ing.columns else 0
    total_gastos = pd.to_numeric(df_gas['Valor'], errors='coerce').sum() if 'Valor' in df_gas.columns else 0
    saldo = total_ingresos - total_gastos
    
    st.markdown("### Resumen Ejecutivo")
    st.write(f"- **Total Recaudado:** ${total_ingresos:,.0f} COP")
    st.write(f"- **Total Invertido/Gastado:** ${total_gastos:,.0f} COP")
    st.write(f"- **Ganancia Neta Obtenida:** ${saldo:,.0f} COP")
    
    st.markdown("### 💡 Evaluación de la Gestión")
    st.success("El proyecto se ejecutó exitosamente cumpliendo con los registros financieros y soportes digitales correspondientes.")
    
    st.markdown("### 📥 Descargar Archivo Excel Actualizado")
    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            label="Descargar Excel Completo",
            data=f,
            file_name="Proyecto_Financiero_Eventos_Actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
