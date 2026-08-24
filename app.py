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

def limpiar_y_obtener_datos(df, col_concepto, col_valor):
    """Filtra filas vacías y la fila de totales para trabajar solo con los registros reales"""
    if df.empty:
        return pd.DataFrame()
    
    # Copiar para no alterar el original visual
    df_temp = df.copy()
    
    # Asegurar nombres de columnas si vienen como índices o texto plano
    # Buscamos filas donde el concepto no sea nulo, no esté vacío y no sea una fila de "TOTAL"
    filtro_validos = []
    for idx, row in df_temp.iterrows():
        concepto_val = str(row.get(col_concepto, '')).strip().lower()
        if concepto_val != 'nan' and concepto_val != '' and 'total' not in concepto_val:
            filtro_validos.append(idx)
            
    return df_temp.loc[filtro_validos]

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

# Identificar nombres reales de columnas en las hojas
df_ing_raw = data_dict.get('Registro de Ingresos', pd.DataFrame())
df_gas_raw = data_dict.get('Registro de Gastos', pd.DataFrame())

col_c_ing = df_ing_raw.columns[1] if len(df_ing_raw.columns) > 1 else 'Concepto'
col_v_ing = df_ing_raw.columns[2] if len(df_ing_raw.columns) > 2 else 'Valor'

col_c_gas = df_gas_raw.columns[1] if len(df_gas_raw.columns) > 1 else 'Concepto'
col_v_gas = df_gas_raw.columns[3] if len(df_gas_raw.columns) > 3 else 'Valor'

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
    
    edited_ingresos = st.data_editor(df_ing_raw, use_container_width=True, num_rows="dynamic", key="editor_ingresos")
    
    if st.button("💾 Guardar cambios en la tabla de Ingresos"):
        data_dict['Registro de Ingresos'] = edited_ingresos
        save_excel_data(data_dict)
        st.success("¡Cambios guardados correctamente!")
        st.rerun()

# --- 3. REGISTRO DE GASTOS ---
elif menu == "3. Registro de Gastos":
    st.title("📉 Registro de Gastos")
    
    edited_gastos = st.data_editor(df_gas_raw, use_container_width=True, num_rows="dynamic", key="editor_gastos")
    
    if st.button("💾 Guardar cambios en la tabla de Gastos"):
        data_dict['Registro de Gastos'] = edited_gastos
        save_excel_data(data_dict)
        st.success("¡Cambios de gastos guardados correctamente!")
        st.rerun()

# --- 4. BALANCE FINANCIERO ---
elif menu == "4. Balance Financiero":
    st.title("⚖️ Balance Financiero General")
    
    df_ing_clean = limpiar_y_obtener_datos(df_ing_raw, col_c_ing, col_v_ing)
    df_gas_clean = limpiar_y_obtener_datos(df_gas_raw, col_c_gas, col_v_gas)
    
    total_ingresos = pd.to_numeric(df_ing_clean[col_v_ing], errors='coerce').sum() if not df_ing_clean.empty else 0
    total_gastos = pd.to_numeric(df_gas_clean[col_v_gas], errors='coerce').sum() if not df_gas_clean.empty else 0
    saldo = total_ingresos - total_gastos
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Ingresos", f"${total_ingresos:,.0f}")
    col2.metric("Total Gastos", f"${total_gastos:,.0f}")
    col3.metric("Saldo / Ganancia Neta", f"${saldo:,.0f}", delta=f"${saldo:,.0f}")

# --- 5. DASHBOARD ---
elif menu == "5. Dashboard y Gráficos":
    st.title("📊 Dashboard y Resumen Visual")
    
    df_ing_clean = limpiar_y_obtener_datos(df_ing_raw, col_c_ing, col_v_ing)
    df_gas_clean = limpiar_y_obtener_datos(df_gas_raw, col_c_gas, col_v_gas)
    
    tot_ing = pd.to_numeric(df_ing_clean[col_v_ing], errors='coerce').sum() if not df_ing_clean.empty else 0
    tot_gas = pd.to_numeric(df_gas_clean[col_v_gas], errors='coerce').sum() if not df_gas_clean.empty else 0
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Comparativa Ingresos vs Gastos")
        st.bar_chart(pd.DataFrame({"Tipo": ["Ingresos", "Gastos"], "Monto": [tot_ing, tot_gas]}).set_index("Tipo"))
    with col2:
        st.subheader("Gastos por Categoría")
        if not df_gas_clean.empty and len(df_gas_clean.columns) > 2:
            col_cat_gas = df_gas_clean.columns[2] # Ajusta según la columna de categoría en gastos
            df_g_copy = df_gas_clean.copy()
            df_g_copy[col_v_gas] = pd.to_numeric(df_g_copy[col_v_gas], errors='coerce')
            st.bar_chart(df_g_copy.groupby(col_cat_gas)[col_v_gas].sum())
        else:
            st.info("No hay suficientes datos de gastos para graficar por categoría.")

# --- 6. ANEXO DE RECIBOS & QR ---
elif menu == "6. Anexo de Recibos & QR":
    st.title("🧾 Generador de Comprobantes y QR desde Registros")
    
    df_ing_clean = limpiar_y_obtener_datos(df_ing_raw, col_c_ing, col_v_ing)
    df_gas_clean = limpiar_y_obtener_datos(df_gas_raw, col_c_gas, col_v_gas)
    
    opciones = []
    
    if not df_ing_clean.empty:
        col_f_ing = df_ing_raw.columns[0]
        for _, r in df_ing_clean.iterrows():
            fec = r.get(col_f_ing, '')
            con = r.get(col_c_ing, '')
            val = pd.to_numeric(r.get(col_v_ing, 0), errors='coerce')
            opciones.append(f"[INGRESO] {fec} - {con} (${val:,.0f})")
                
    if not df_gas_clean.empty:
        col_f_gas = df_gas_raw.columns[0]
        for _, r in df_gas_clean.iterrows():
            fec = r.get(col_f_gas, '')
            con = r.get(col_c_gas, '')
            val = pd.to_numeric(r.get(col_v_gas, 0), errors='coerce')
            opciones.append(f"[GASTO] {fec} - {con} (${val:,.0f})")

    if not opciones:
        st.warning("⚠️ No se encontraron registros válidos en las tablas.")
    else:
        mov_sel = st.selectbox("Selecciona el movimiento:", opciones)
        
        if st.button("Generar Comprobante y QR"):
            is_ingreso = "[INGRESO]" in mov_sel
            rec_id = f"REC-{abs(hash(mov_sel)) % 10000:04d}"
            
            df_origen = df_ing_clean if is_ingreso else df_gas_clean
            col_con_activo = col_c_ing if is_ingreso else col_c_gas
            col_val_activo = col_v_ing if is_ingreso else col_v_gas
            col_fec_activo = df_ing_raw.columns[0] if is_ingreso else df_gas_raw.columns[0]
            
            fila = df_origen[df_origen[col_con_activo].astype(str).apply(lambda x: x in mov_sel)]
            
            if not fila.empty:
                f_data = fila.iloc[0]
                fecha = f_data.get(col_fec_activo, 'N/A')
                concepto = f_data.get(col_con_activo, 'N/A')
                valor = pd.to_numeric(f_data.get(col_val_activo, 0), errors='coerce')
                
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
    
    df_ing_clean = limpiar_y_obtener_datos(df_ing_raw, col_c_ing, col_v_ing)
    df_gas_clean = limpiar_y_obtener_datos(df_gas_raw, col_c_gas, col_v_gas)
    
    tot_ing = pd.comprobantes if False else (pd.to_numeric(df_ing_clean[col_v_ing], errors='coerce').sum() if not df_ing_clean.empty else 0)
    tot_gas = pd.to_numeric(df_gas_clean[col_v_gas], errors='coerce').sum() if not df_gas_clean.empty else 0
    saldo = tot_ing - tot_gas
    
    st.write(f"- **Total Recaudado:** ${tot_ing:,.0f} COP")
    st.write(f"- **Total Invertido/Gastado:** ${tot_gas:,.0f} COP")
    st.write(f"- **Ganancia Neta Obtenida:** ${saldo:,.0f} COP")
    
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("Descargar Excel Completo", data=f, file_name="Proyecto_Financiero_Eventos_Actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
