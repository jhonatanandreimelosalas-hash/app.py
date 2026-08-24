import streamlit as st
import pandas as pd
import openpyxl
import qrcode
from io import BytesIO
from PIL import Image

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
    st.info("**Versión:** v2.1 (Aplicación Interactiva Mejorada)")
    st.info("**Estado:** En Desarrollo / Proyecto Base")
    
    st.markdown("### 👥 2. Equipo de Trabajo / Integrantes")
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
    
    # Limpiar filas nulas si las hay
    df_ingresos = df_ingresos.dropna(subset=['Fecha']) if 'Fecha' in df_ingresos.columns else df_ingresos
    
    st.markdown("### 📝 Tabla Interactiva (Puedes editar celdas o borrar filas directamente)")
    edited_ingresos = st.data_editor(df_ingresos, use_container_width=True, num_rows="dynamic", key="editor_ingresos")
    
    if st.button("💾 Guardar cambios de la tabla"):
        data_dict['Registro de Ingresos'] = edited_ingresos
        save_excel_data(data_dict)
        st.success("¡Cambios guardados correctamente en el Excel!")
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
            st.success("¡Ingreso agregado y guardado en el Excel correctamente!")
            st.rerun()

# --- 3. REGISTRO DE GASTOS ---
elif menu == "3. Registro de Gastos":
    st.title("📉 Registro de Gastos")
    df_gastos = data_dict['Registro de Gastos']
    df_gastos = df_gastos.dropna(subset=['Fecha']) if 'Fecha' in df_gastos.columns else df_gastos
    
    st.markdown("### 📝 Tabla Interactiva (Puedes editar celdas o borrar filas directamente)")
    edited_gastos = st.data_editor(df_gastos, use_container_width=True, num_rows="dynamic", key="editor_gastos")
    
    if st.button("💾 Guardar cambios de la tabla de gastos"):
        data_dict['Registro de Gastos'] = edited_gastos
        save_excel_data(data_dict)
        st.success("¡Cambios de gastos guardados correctamente en el Excel!")
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
            st.success("¡Gasto agregado y guardado en el Excel correctamente!")
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
            df_gas['Valor'] = pd.to_numeric(df_gas['Valor'], errors='coerce')
            cat_grouped = df_gas.groupby('Categoría')['Valor'].sum()
            st.bar_chart(cat_grouped)
        else:
            st.info("No hay suficientes datos de categorías.")

# --- 6. ANEXO DE RECIBOS & QR ---
elif menu == "6. Anexo de Recibos & QR":
    st.title("🧾 Anexo de Recibos y Generador de Códigos QR")
    
    df_anexo = data_dict['Anexo de recibos']
    st.dataframe(df_anexo, use_container_width=True)
    
    st.markdown("### ➕ Registrar Comprobante / Recibo y Generar QR")
    with st.form("form_qr"):
        col1, col2 = st.columns(2)
        with col1:
            rec_id = st.text_input("ID Recibo (ej. REC-009)", value=f"REC-00{len(df_anexo)+1}")
            fecha = st.date_input("Fecha Recibo")
            tipo = st.selectbox("Tipo", ["Ingreso", "Gasto"])
            concepto = st.text_input("Concepto del Recibo")
        with col2:
            valor = st.number_input("Valor ($)", min_value=0.0, step=1000.0)
            enlace = st.text_input("Enlace del Comprobante (Drive, ImgBB, URL)", value="https://drive.google.com/...")
            
        submitted = st.form_submit_button("Generar QR y Guardar en Anexo")
        if submitted:
            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(enlace)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            
            buf = BytesIO()
            img.save(buf, format="PNG")
            
            nuevo_anexo = pd.DataFrame({
                'ID': [rec_id],
                'Fecha': [str(fecha)],
                'Tipo (Ingreso/Gasto)': [tipo],
                'Concepto': [concepto],
                'Valor': [valor],
                'Nombre del archivo o enlace': [enlace],
                'Código QR': [f"[QR Generado para {enlace}]"]
            })
            data_dict['Anexo de recibos'] = pd.concat([df_anexo, nuevo_anexo], ignore_index=True)
            save_excel_data(data_dict)
            
            st.success("¡Recibo registrado con éxito y QR generado!")
            st.image(buf.getvalue(), caption=f"Código QR para {rec_id} - {concepto}", width=200)

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
            label="Descargar Excel Completo con Nuevos Registros",
            data=f,
            file_name="Proyecto_Financiero_Eventos_Actualizado (1).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
