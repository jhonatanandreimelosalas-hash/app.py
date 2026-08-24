import streamlit as st
import pandas as pd
import openpyxl
import qrcode
from io import BytesIO

st.set_page_config(page_title="Gestión Financiera - Prototipo Eventos", page_icon="💰", layout="wide")

EXCEL_FILE = "Proyecto_Financiero_Eventos_Actualizado (1).xlsx"

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
    st.cache_data.clear()

def buscar_columna(df, posibles_nombres):
    for col in df.columns:
        col_limpia = str(col).strip().lower()
        for p in posibles_nombres:
            if p.lower() in col_limpia:
                return col
    return None

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
    st.subheader("Sistema Genérico para Control de Eventos")
    st.markdown("---")
    st.markdown("### 👥 Equipo / Integrantes")
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
    
    edited_ingresos = st.data_editor(df_ingresos, use_container_width=True, num_rows="dynamic", key="editor_ingresos")
    
    if st.button("💾 Guardar cambios en la tabla de Ingresos"):
        data_dict['Registro de Ingresos'] = edited_ingresos
        save_excel_data(data_dict)
        st.success("¡Cambios guardados correctamente!")
        st.rerun()
    
    st.markdown("---")
    st.markdown("### ➕ Agregar Nuevo Ingreso")
    with st.form("form_ingreso"):
        col1, col2 = st.columns(2)
        with col1:
            fecha = st.date_input("Fecha")
            concepto = st.text_input("Concepto")
            valor = st.number_input("Valor ($)", min_value=0.0, step=1000.0)
        with col2:
            responsable = st.selectbox("Responsable", INTEGRANTES_LISTA)
            observaciones = st.text_area("Observaciones")
        
        if st.form_submit_button("Guardar Ingreso"):
            nuevo_reg = pd.DataFrame({'Fecha': [str(fecha)], 'Concepto': [concepto], 'Valor': [valor], 'Responsable': [responsable], 'Observaciones': [observaciones]})
            data_dict['Registro de Ingresos'] = pd.concat([df_ingresos, nuevo_reg], ignore_index=True)
            save_excel_data(data_dict)
            st.success("¡Ingreso agregado correctamente!")
            st.rerun()

# --- 3. REGISTRO DE GASTOS ---
elif menu == "3. Registro de Gastos":
    st.title("📉 Registro de Gastos")
    df_gastos = data_dict['Registro de Gastos']
    
    edited_gastos = st.data_editor(df_gastos, use_container_width=True, num_rows="dynamic", key="editor_gastos")
    
    if st.button("💾 Guardar cambios en la tabla de Gastos"):
        data_dict['Registro de Gastos'] = edited_gastos
        save_excel_data(data_dict)
        st.success("¡Cambios de gastos guardados correctamente!")
        st.rerun()
        
    st.markdown("---")
    st.markdown("### ➕ Agregar Nuevo Gasto")
    with st.form("form_gasto"):
        col1, col2 = st.columns(2)
        with col1:
            fecha = st.date_input("Fecha Gasto")
            concepto = st.text_input("Concepto")
            categoria = st.selectbox("Categoría", ["Logística", "Publicidad", "Alimentación", "Varios"])
        with col2:
            valor = st.number_input("Valor ($)", min_value=0.0, step=1000.0)
            responsable = st.selectbox("Responsable", INTEGRANTES_LISTA)
            observaciones = st.text_area("Observaciones")
        
        if st.form_submit_button("Guardar Gasto"):
            nuevo_reg = pd.DataFrame({'Fecha': [str(fecha)], 'Concepto': [concepto], 'Categoría': [categoria], 'Valor': [valor], 'Responsable': [responsable], 'Observaciones': [observaciones]})
            data_dict['Registro de Gastos'] = pd.concat([df_gastos, nuevo_reg], ignore_index=True)
            save_excel_data(data_dict)
            st.success("¡Gasto agregado correctamente!")
            st.rerun()

# --- 4. BALANCE FINANCIERO ---
elif menu == "4. Balance Financiero":
    st.title("⚖️ Balance Financiero General")
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    col_val_ing = buscar_columna(df_ing, ['valor', 'monto', 'precio'])
    col_val_gas = buscar_columna(df_gas, ['valor', 'monto', 'precio'])
    
    total_ingresos = pd.to_numeric(df_ing[col_val_ing], errors='coerce').sum() if col_val_ing else 0
    total_gastos = pd.to_numeric(df_gas[col_val_gas], errors='coerce').sum() if col_val_gas else 0
    saldo = total_ingresos - total_gastos
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Ingresos", f"${total_ingresos:,.0f}")
    col2.metric("Total Gastos", f"${total_gastos:,.0f}")
    col3.metric("Saldo / Ganancia Neta", f"${saldo:,.0f}", delta=f"${saldo:,.0f}")

# --- 5. DASHBOARD ---
elif menu == "5. Dashboard y Gráficos":
    st.title("📊 Dashboard y Resumen Visual")
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    col_v_ing = buscar_columna(df_ing, ['valor', 'monto'])
    col_v_gas = buscar_columna(df_gas, ['valor', 'monto'])
    
    tot_ing = pd.to_numeric(df_ing[col_v_ing], errors='coerce').sum() if col_v_ing else 0
    tot_gas = pd.to_numeric(df_gas[col_v_gas], errors='coerce').sum() if col_v_gas else 0
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Comparativa Ingresos vs Gastos")
        st.bar_chart(pd.DataFrame({"Tipo": ["Ingresos", "Gastos"], "Monto": [tot_ing, tot_gas]}).set_index("Tipo"))
    with col2:
        st.subheader("Gastos por Categoría")
        col_cat = buscar_columna(df_gas, ['categoría', 'categoria', 'tipo'])
        if col_cat and col_v_gas:
            df_g_copy = df_gas.copy()
            df_g_copy[col_v_gas] = pd.to_numeric(df_g_copy[col_v_gas], errors='coerce')
            st.bar_chart(df_g_copy.groupby(col_cat)[col_v_gas].sum())
        else:
            st.info("No se encontró columna de categoría.")

# --- 6. ANEXO DE RECIBOS & QR ---
elif menu == "6. Anexo de Recibos & QR":
    st.title("🧾 Generador de Comprobantes y QR desde Registros")
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    opciones = []
    
    # Mapear columnas dinámicamente
    c_f_ing = buscar_columna(df_ing, ['fecha'])
    c_c_ing = buscar_columna(df_ing, ['concepto', 'detalle', 'descripcion'])
    c_v_ing = buscar_columna(df_ing, ['valor', 'monto'])
    
    if not df_ing.empty and c_c_ing:
        for _, r in df_ing.iterrows():
            fec = r[c_f_ing] if c_f_ing else ""
            con = r[c_c_ing] if c_c_ing else ""
            val = pd.to_numeric(r[c_v_ing], errors='coerce') if c_v_ing else 0
            if pd.notna(con):
                opciones.append(f"[INGRESO] {fec} - {con} (${val:,.0f})")
                
    c_f_gas = buscar_columna(df_gas, ['fecha'])
    c_c_gas = buscar_columna(df_gas, ['concepto', 'detalle', 'descripcion'])
    c_v_gas = buscar_columna(df_gas, ['valor', 'monto'])
    
    if not df_gas.empty and c_c_gas:
        for _, r in df_gas.iterrows():
            fec = r[c_f_gas] if c_f_gas else ""
            con = r[c_c_gas] if c_c_gas else ""
            val = pd.to_numeric(r[c_v_gas], errors='coerce') if c_v_gas else 0
            if pd.notna(con):
                opciones.append(f"[GASTO] {fec} - {con} (${val:,.0f})")

    if not opciones:
        st.warning("No hay registros válidos encontrados en las tablas para certificar.")
    else:
        mov_sel = st.selectbox("Selecciona el movimiento:", opciones)
        
        if st.button("Generar Comprobante y QR"):
            is_ingreso = "[INGRESO]" in mov_sel
            rec_id = f"REC-{abs(hash(mov_sel)) % 10000:04d}"
            
            # Extraer datos de la fila seleccionada
            df_origen = df_ing if is_ingreso else df_gas
            c_con = c_c_ing if is_ingreso else c_c_gas
            c_val = c_v_ing if is_ingreso else c_v_gas
            c_fec = c_f_ing if is_ingreso else c_f_gas
            
            fila = df_origen[df_origen[c_con].astype(str).apply(lambda x: x in mov_sel)]
            
            if not fila.empty:
                f_data = fila.iloc[0]
                fecha = f_data[c_fec] if c_fec else "N/A"
                concepto = f_data[c_con] if c_con else "N/A"
                valor = pd.to_numeric(f_data[c_val], errors='coerce') if c_val else 0
                
                texto_recibo = (
                    f"=== COMPROBANTE OFICIAL DE EVENTO ===\n"
                    f"ID: {rec_id}\n"
                    f"Tipo: {'Ingreso' if is_ingreso else 'Gasto'}\n"
                    f"Fecha: {fecha}\n"
                    f"Concepto: {concepto}\n"
                    f"Valor: ${valor:,.0f} COP\n"
                    f"Estado: Verificado y Aprobado"
                )
                
                texto_qr = f"ID:{rec_id}|VALOR:${valor:,.0f}COP|CONCEPTO:{concepto}"
                
                qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=2)
                qr.add_data(texto_qr)
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")
                
                buf = BytesIO()
                img.save(buf, format="PNG")
                
                st.session_state.ultimo_recibo_texto = texto_recibo
                st.session_state.ultimo_recibo_id = rec_id
                st.session_state.ultimo_qr_img = buf.getvalue()
                st.success("¡Comprobante generado con éxito!")

    if 'ultimo_recibo_texto' in st.session_state and st.session_state.ultimo_recibo_texto is not None:
        st.markdown("---")
        st.subheader(f"📄 Vista Previa: {st.session_state.ultimo_recibo_id}")
        c1, c2 = st.columns(2)
        with c1: st.text(st.session_state.ultimo_recibo_texto)
        with c2: st.image(st.session_state.ultimo_qr_img, width=180)
        st.download_button("📥 Descargar Comprobante (.txt)", st.session_state.ultimo_recibo_texto, file_name=f"{st.session_state.ultimo_recibo_id}.txt")

# --- 7. REPORTE FINAL ---
elif menu == "7. Reporte Final":
    st.title("📑 Reporte Final del Evento")
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    col_v_ing = buscar_columna(df_ing, ['valor', 'monto'])
    col_v_gas = buscar_columna(df_gas, ['valor', 'monto'])
    
    tot_ing = pd.to_numeric(df_ing[col_v_ing], errors='coerce').sum() if col_v_ing else 0
    tot_gas = pd.to_numeric(df_gas[col_v_gas], errors='coerce').sum() if col_v_gas else 0
    saldo = tot_ing - tot_gas
    
    st.write(f"- **Total Recaudado:** ${tot_ing:,.0f} COP")
    st.write(f"- **Total Invertido/Gastado:** ${tot_gas:,.0f} COP")
    st.write(f"- **Ganancia Neta Obtenida:** ${saldo:,.0f} COP")
    
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("Descargar Excel Completo", data=f, file_name="Proyecto_Financiero_Eventos_Actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
