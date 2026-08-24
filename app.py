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
    st.info("**Versión:** v2.3 (Optimizado y Conectado)")
    
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
    
    column_config = {}
    if 'Responsable' in df_ingresos.columns:
        column_config['Responsable'] = st.column_config.SelectboxColumn("Responsable", options=INTEGRANTES_LISTA, required=True)

    edited_ingresos = st.data_editor(df_ingresos, use_container_width=True, num_rows="dynamic", key="editor_ingresos", column_config=column_config)
    
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
        
        if st.form_submit_button("Guardar Ingreso"):
            nuevo_reg = pd.DataFrame({'Fecha': [str(fecha)], 'Concepto': [concepto], 'Valor': [valor], 'Responsable': [responsable], 'Observaciones': [observaciones]})
            data_dict['Registro de Ingresos'] = pd.concat([df_ingresos, nuevo_reg], ignore_index=True)
            save_excel_data(data_dict)
            st.success("¡Ingreso agregado y guardado correctamente!")
            st.rerun()

# --- 3. REGISTRO DE GASTOS ---
elif menu == "3. Registro de Gastos":
    st.title("📉 Registro de Gastos")
    df_gastos = data_dict['Registro de Gastos']
    
    column_config_g = {}
    if 'Responsable' in df_gastos.columns:
        column_config_g['Responsable'] = st.column_config.SelectboxColumn("Responsable", options=INTEGRANTES_LISTA, required=True)
    if 'Categoría' in df_gastos.columns:
        column_config_g['Categoría'] = st.column_config.SelectboxColumn("Categoría", options=["Logística", "Publicidad", "Alimentación", "Varios"], required=True)

    edited_gastos = st.data_editor(df_gastos, use_container_width=True, num_rows="dynamic", key="editor_gastos", column_config=column_config_g)
    
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
        
        if st.form_submit_button("Guardar Gasto"):
            nuevo_reg = pd.DataFrame({'Fecha': [str(fecha)], 'Concepto': [concepto], 'Categoría': [categoria], 'Valor': [valor], 'Responsable': [responsable], 'Observaciones': [observaciones]})
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
    balance_df = pd.DataFrame({"Concepto": ["Total de Ingresos", "Total de Gastos", "Saldo Final"], "Valor ($)": [total_ingresos, total_gastos, saldo]})
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
        st.bar_chart(pd.DataFrame({"Tipo": ["Ingresos", "Gastos"], "Monto": [total_ingresos, total_gastos]}).set_index("Tipo"))
    with col2:
        st.subheader("Gastos por Categoría")
        if 'Categoría' in df_gas.columns and 'Valor' in df_gas.columns:
            df_gas_copy = df_gas.copy()
            df_gas_copy['Valor'] = pd.to_numeric(df_gas_copy['Valor'], errors='coerce')
            st.bar_chart(df_gas_copy.groupby('Categoría')['Valor'].sum())
        else:
            st.info("No hay suficientes datos de categorías.")

# --- 6. ANEXO DE RECIBOS & QR (CONECTADO A INGRESOS Y GASTOS) ---
elif menu == "6. Anexo de Recibos & QR":
    st.title("🧾 Generador de Comprobantes y QR desde Registros")
    st.markdown("Selecciona cualquier movimiento real registrado en tus ingresos o gastos para generarle su comprobante oficial y código QR.")
    
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    # Unificar ingresos y gastos en una sola lista seleccionable
    opciones_movimientos = []
    
    if not df_ing.empty and 'Concepto' in df_ing.columns:
        for idx, row in df_ing.iterrows():
            opciones_movimientos.append(f"[INGRESO] {row.get('Fecha', '')} - {row.get('Concepto', '')} (${row.get('Valor', 0):,.0f})")
            
    if not df_gas.empty and 'Concepto' in df_gas.columns:
        for idx, row in df_gas.iterrows():
            opciones_movimientos.append(f"[GASTO] {row.get('Fecha', '')} - {row.get('Concepto', '')} (${row.get('Valor', 0):,.0f})")

    if not opciones_movimientos:
        st.warning("Primero debes registrar al menos un ingreso o un gasto en las secciones anteriores.")
    else:
        movimiento_seleccionado = st.selectbox("Selecciona el Ingreso o Gasto a certificar:", opciones_movimientos)
        
        if st.button("Generar Comprobante y QR del Movimiento"):
            # Extraer datos de la opción seleccionada
            tipo_mov = "Ingreso" if "[INGRESO]" in movimiento_seleccionado else "Gasto"
            rec_id = f"REC-{abs(hash(movimiento_seleccionado)) % 10000:04d}"
            
            # Buscar el detalle exacto en los dataframes
            if tipo_mov == "Ingreso":
                fila_encontrada = df_ing[df_ing.apply(lambda r: f"[INGRESO] {r.get('Fecha', '')} - {r.get('Concepto', '')} (${r.get('Valor', 0):,.0f})" == movimiento_seleccionado, axis=1)]
            else:
                fila_encontrada = df_gas[df_gas.apply(lambda r: f"[GASTO] {r.get('Fecha', '')} - {r.get('Concepto', '')} (${r.get('Valor', 0):,.0f})" == movimiento_seleccionado, axis=1)]
            
            if not fila_encontrada.empty:
                f_data = fila_encontrada.iloc[0]
                fecha = f_data.get('Fecha', 'N/A')
                concepto = f_data.get('Concepto', 'N/A')
                valor = float(f_data.get('Valor', 0))
                responsable = f_data.get('Responsable', 'Equipo')
                obs = f_data.get('Observaciones', f_data.get('Categoría', 'General'))
                
                texto_recibo = (
                    f"=== COMPROBANTE OFICIAL DE EVENTO ===\n"
                    f"ID: {rec_id}\n"
                    f"Tipo: {tipo_mov}\n"
                    f"Fecha: {fecha}\n"
                    f"Concepto: {concepto}\n"
                    f"Categoría/Obs: {obs}\n"
                    f"Valor: ${valor:,.0f} COP\n"
                    f"Responsable: {responsable}\n"
                    f"Estado: Verificado y Aprobado"
                )
                
                texto_qr = f"ID:{rec_id}|TIPO:{tipo_mov}|CONCEPTO:{concepto}|VALOR:${valor:,.0f}COP|RESP:{responsable}|OK"
                
                qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=2)
                qr.add_data(texto_qr)
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")
                
                buf = BytesIO()
                img.save(buf, format="PNG")
                
                # Guardar en la hoja de anexo del excel
                df_anexo = data_dict['Anexo de recibos']
                nuevo_anexo = pd.DataFrame({
                    'ID': [rec_id],
                    'Fecha': [str(fecha)],
                    'Tipo (Ingreso/Gasto)': [tipo_mov],
                    'Concepto': [concepto],
                    'Valor': [valor],
                    'Nombre del archivo o enlace': [f"Comprobante_{rec_id}.txt"],
                    'Código QR': [f"QR Oficial - {concepto}"]
                })
                data_dict['Anexo de recibos'] = pd.concat([df_anexo, nuevo_anexo], ignore_index=True)
                save_excel_data(data_dict)
                
                st.session_state.ultimo_recibo_texto = texto_recibo
                st.session_state.ultimo_recibo_id = rec_id
                st.session_state.ultimo_qr_img = buf.getvalue()
                st.success("¡Comprobante vinculado y generado con éxito!")

    if 'ultimo_recibo_texto' in st.session_state and st.session_state.ultimo_recibo_texto is not None:
        st.markdown("---")
        st.subheader(f"📄 Vista Previa: {st.session_state.ultimo_recibo_id}")
        c1, c2 = st.columns(2)
        with c1: 
            st.text(st.session_state.ultimo_recibo_texto)
        with c2: 
            st.image(st.session_state.ultimo_qr_img, width=180)
        st.download_button(
            label="📥 Descargar Comprobante (.txt)", 
            data=st.session_state.ultimo_recibo_texto, 
            file_name=f"{st.session_state.ultimo_recibo_id}.txt", 
            mime="text/plain"
        )

# --- 7. REPORTE FINAL ---
elif menu == "7. Reporte Final":
    st.title("📑 Reporte Final del Evento")
    df_ing = data_dict['Registro de Ingresos']
    df_gas = data_dict['Registro de Gastos']
    
    total_ingresos = pd.to_numeric(df_ing['Valor'], errors='coerce').sum() if 'Valor' in df_ing.columns else 0
    total_gastos = pd.to_numeric(df_gas['Valor'], errors='coerce').sum() if 'Valor' in df_gas.columns else 0
    saldo = total_ingresos - total_gastos
    
    st.write(f"- **Total Recaudado:** ${total_ingresos:,.0f} COP")
    st.write(f"- **Total Invertido/Gastado:** ${total_gastos:,.0f} COP")
    st.write(f"- **Ganancia Neta Obtenida:** ${saldo:,.0f} COP")
    
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("Descargar Excel Completo", data=f, file_name="Proyecto_Financiero_Eventos_Actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
