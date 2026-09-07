import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import qrcode
from io import BytesIO
from datetime import datetime
import plotly.express as px

# Configuración inicial de la página
st.set_page_config(
    page_title="Gestión Financiera - Prototipo Eventos", 
    page_icon="💰", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS modernos y personalizados
st.markdown("""
    <style>
        .main-header { font-size: 2.3rem; color: #1E3A8A; font-weight: 800; margin-bottom: 0px; letter-spacing: -0.5px; }
        .sub-header { font-size: 1.1rem; color: #4B5563; margin-bottom: 20px; }
        .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; background-color: #1E3A8A; color: white; transition: 0.3s; }
        .stButton>button:hover { background-color: #2563EB; border-color: #2563EB; }
        div.stMetric { background-color: #F8FAFC; padding: 15px 20px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid #E2E8F0; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Proyecto_Financiero_Eventos_Actualizado (1).xlsx"

INTEGRANTES_LISTA = [
    "Ivan Santiago Valencia Villamil",
    "Nicol Vanegas Cruz",
    "Jhonatan Andrey Melo",
    "Alejandro Martinez Rubio"
]

# --- INICIALIZAR ESTADO DE DATOS LIMPIOS EN SESSION_STATE ---
if 'ingresos_df' not in st.session_state:
    st.session_state.ingresos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Valor", "Responsable", "Observaciones"])

if 'gastos_df' not in st.session_state:
    st.session_state.gastos_df = pd.DataFrame(columns=["Fecha", "Concepto", "Categoría", "Valor", "Responsable"])

def guardar_todo_en_excel():
    """Sincroniza los DataFrames actuales con el archivo Excel manteniendo el formato"""
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        st.session_state.ingresos_df.to_excel(writer, sheet_name='Registro de Ingresos', index=False)
        st.session_state.gastos_df.to_excel(writer, sheet_name='Registro de Gastos', index=False)
    
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        font_body = Font(name="Arial", size=10)
        thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
                for cell in row:
                    cell.border = thin_border
                    if row_idx == 1:
                        cell.font = font_header
                        cell.fill = fill_header
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    else:
                        cell.font = font_body
                        cell.alignment = Alignment(horizontal="left", vertical="center")
            for col in ws.columns:
                max_len = max([len(str(cell.value or '')) for cell in col])
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        wb.save(EXCEL_FILE)
    except Exception:
        pass

# --- MENÚ LATERAL Y BACKUP ---
st.sidebar.markdown("### 💰 Control Financiero")
st.sidebar.markdown("---")

# Configuración de Presupuesto Máximo en la barra lateral
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
    "7. Reporte Final"
])
st.sidebar.markdown("---")

# Botón de Copia de Seguridad (Backup) en la barra lateral
st.sidebar.markdown("📦 **Copias de Seguridad**")
guardar_todo_en_excel()
try:
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="📥 Descargar Backup Diario",
            data=f,
            file_name=f"Backup_Financiero_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except Exception:
    pass

# --- 1. INICIO ---
if menu == "1. Inicio":
    st.markdown('<p class="main-header">🏛️ Proyecto de Control y Gestión Financiera</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Plataforma centralizada para la administración y supervisión de recursos en eventos</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("### 🎯 Objetivo del Sistema")
        st.write("Control transparente y automatizado de los movimientos monetarios, auditoría en tiempo real, gestión de presupuestos y generación de comprobantes con Plotly.")
    with col2:
        st.success("✅ **Estado del Sistema:** Operativo y Sincronizado.")

    st.markdown("---")
    st.markdown("### 👥 Equipo de Trabajo")
    integrantes_data = [
        {"N.°": 1, "Nombre Completo": "Ivan Santiago Valencia Villamil", "Rol / Responsabilidad": "Líder de Proyecto / Administración"},
        {"N.°": 2, "Nombre Completo": "Nicol Vanegas Cruz", "Rol / Responsabilidad": "Gestión de Registro e Ingresos"},
        {"N.°": 3, "Nombre Completo": "Jhonatan Andrey Melo", "Rol / Responsabilidad": "Control de Gastos e Insumos"},
        {"N.°": 4, "Nombre Completo": "Alejandro Martinez Rubio", "Rol / Responsabilidad": "Soportes y Control de Balance"},
    ]
    st.dataframe(pd.DataFrame(integrantes_data), use_container_width=True, hide_index=True)

# --- 2. REGISTRO DE INGRESOS ---
elif menu == "2. Registro de Ingresos":
    st.markdown('<p class="main-header">📈 Registro de Ingresos</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Agrega y administra las entradas económicas</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    with st.expander("➕ Agregar Nuevo Ingreso", expanded=True):
        with st.form("form_nuevo_ingreso"):
            c1, c2 = st.columns(2)
            with c1:
                f_ing = st.date_input("Fecha", value=datetime.now())
                con_ing = st.text_input("Concepto (ej. Venta de boletería)")
            with c2:
                resp_ing = st.selectbox("Responsable", INTEGRANTES_LISTA)
                val_ing = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
            obs_ing = st.text_area("Observaciones (Opcional)")
            
            btn_guardar_ing = st.form_submit_button("Guardar Ingreso")
            if btn_guardar_ing:
                if con_ing.strip() == "":
                    st.error("⚠️ El concepto no puede estar vacío.")
                else:
                    nuevo_reg = {
                        "Fecha": f_ing.strftime("%Y-%m-%d"),
                        "Concepto": con_ing,
                        "Valor": float(val_ing),
                        "Responsable": resp_ing,
                        "Observaciones": obs_ing
                    }
                    st.session_state.ingresos_df = pd.concat([st.session_state.ingresos_df, pd.DataFrame([nuevo_reg])], ignore_index=True)
                    guardar_todo_en_excel()
                    st.success("¡Ingreso agregado y sumado exitosamente!")
                    st.rerun()

    st.markdown("### 📋 Listado Actual de Ingresos y Filtros")
    if not st.session_state.ingresos_df.empty:
        # Filtros Rápidos
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            filtro_resp_i = st.selectbox("Filtrar por Responsable (Ingresos):", ["Todos"] + INTEGRANTES_LISTA)
        
        df_mostrar_i = st.session_state.ingresos_df.copy()
        if filtro_resp_i != "Todos":
            df_mostrar_i = df_mostrar_i[df_mostrar_i["Responsable"] == filtro_resp_i]

        st.dataframe(df_mostrar_i, use_container_width=True, hide_index=True)
        total_ing = st.session_state.ingresos_df["Valor"].astype(float).sum()
        st.metric(label="💵 TOTAL INGRESOS", value=f"${total_ing:,.0f} COP")
    else:
        st.info("No hay ingresos registrados todavía. Usa el formulario de arriba para agregar uno.")

# --- 3. REGISTRO DE GASTOS ---
elif menu == "3. Registro de Gastos":
    st.markdown('<p class="main-header">📉 Registro de Gastos</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Controla los egresos y compras del evento con alertas de presupuesto</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Alerta visual de presupuesto tope
    current_total_gastos = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    if presupuesto_tope > 0 and current_total_gastos > presupuesto_tope:
        st.error(f"🚨 ¡ATENCIÓN! Los gastos actuales (${current_total_gastos:,.0f}) superan el presupuesto límite configurado (${presupuesto_tope:,.0f}).")
    elif presupuesto_tope > 0:
        st.info(f"ℹ️ Presupuesto disponible: ${(presupuesto_tope - current_total_gastos):,.0f} COP de un tope de ${presupuesto_tope:,.0f} COP.")

    with st.expander("➕ Agregar Nuevo Gasto", expanded=True):
        with st.form("form_nuevo_gasto"):
            c1, c2 = st.columns(2)
            with c1:
                f_gas = st.date_input("Fecha Gasto", value=datetime.now())
                con_gas = st.text_input("Concepto (ej. Alquiler de sonido)")
                cat_gas = st.selectbox("Categoría", ["Logística", "Publicidad", "Alimentación", "Varios"])
            with c2:
                val_gas = st.number_input("Valor ($)", min_value=0.0, step=1000.0, format="%.2f")
                resp_gas = st.selectbox("Responsable", INTEGRANTES_LISTA)
            
            btn_guardar_gas = st.form_submit_button("Guardar Gasto")
            if btn_guardar_gas:
                if con_gas.strip() == "":
                    st.error("⚠️ El concepto no puede estar vacío.")
                else:
                    nuevo_reg_g = {
                        "Fecha": f_gas.strftime("%Y-%m-%d"),
                        "Concepto": con_gas,
                        "Categoría": cat_gas,
                        "Valor": float(val_gas),
                        "Responsable": resp_gas
                    }
                    st.session_state.gastos_df = pd.concat([st.session_state.gastos_df, pd.DataFrame([nuevo_reg_g])], ignore_index=True)
                    guardar_todo_en_excel()
                    st.success("¡Gasto agregado y sumado exitosamente!")
                    st.rerun()

    st.markdown("### 📋 Listado Actual de Gastos y Filtros")
    if not st.session_state.gastos_df.empty:
        # Filtros Rápidos
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            filtro_resp_g = st.selectbox("Filtrar por Responsable (Gastos):", ["Todos"] + INTEGRANTES_LISTA)
        with f_col2:
            filtro_cat_g = st.selectbox("Filtrar por Categoría:", ["Todas", "Logística", "Publicidad", "Alimentación", "Varios"])

        df_mostrar_g = st.session_state.gastos_df.copy()
        if filtro_resp_g != "Todos":
            df_mostrar_g = df_mostrar_g[df_mostrar_g["Responsable"] == filtro_resp_g]
        if filtro_cat_g != "Todas":
            df_mostrar_g = df_mostrar_g[df_mostrar_g["Categoría"] == filtro_cat_g]

        st.dataframe(df_mostrar_g, use_container_width=True, hide_index=True)
        total_gas = st.session_state.gastos_df["Valor"].astype(float).sum()
        st.metric(label="💸 TOTAL GASTOS", value=f"${total_gas:,.0f} COP")
    else:
        st.info("No hay gastos registrados todavía. Usa el formulario de arriba para agregar uno.")

# --- 4. BALANCE FINANCIERO ---
elif menu == "4. Balance Financiero":
    st.markdown('<p class="main-header">⚖️ Balance Financiero General</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Resumen contable actualizado automáticamente</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    saldo = tot_ing - tot_gas
    
    c1, c2, c3 = st.columns(3)
    c1.metric("💵 Total Ingresos", f"${tot_ing:,.0f} COP")
    c2.metric("💸 Total Gastos", f"${tot_gas:,.0f} COP")
    c3.metric("💰 Ganancia Neta", f"${saldo:,.0f} COP", delta=f"${saldo:,.0f} COP")

# --- 5. DASHBOARD ---
elif menu == "5. Dashboard y Gráficos":
    st.markdown('<p class="main-header">📊 Dashboard y Resumen Visual</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Análisis gráfico avanzado con gráficos interactivos (Plotly)</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### ⚖️ Comparativa Ingresos vs Gastos")
        df_comp = pd.DataFrame({"Tipo": ["Ingresos", "Gastos"], "Monto": [tot_ing, tot_gas]})
        fig_bar = px.bar(df_comp, x="Tipo", y="Monto", color="Tipo", text_auto=True, color_discrete_sequence=["#10B981", "#EF4444"])
        fig_bar.update_layout(showlegend=False, margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("#### 🍩 Gastos por Categoría (Gráfico Circular)")
        if not st.session_state.gastos_df.empty:
            df_cat = st.session_state.gastos_df.groupby("Categoría")["Valor"].sum().reset_index()
            fig_pie = px.pie(df_cat, names="Categoría", values="Valor", hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3)
            fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20))
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No hay datos de gastos suficientes para graficar la distribución.")

# --- 6. ANEXO DE RECIBOS & QR ---
elif menu == "6. Anexo de Recibos & QR":
    st.markdown('<p class="main-header">🧾 Generador de Comprobantes y Códigos QR</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Emite soportes oficiales de cada movimiento</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    opciones = []
    for _, r in st.session_state.ingresos_df.iterrows():
        opciones.append(f"[INGRESO] {r['Fecha']} - {r['Concepto']} (${float(r['Valor']):,.0f})")
    for _, r in st.session_state.gastos_df.iterrows():
        opciones.append(f"[GASTO] {r['Fecha']} - {r['Concepto']} (${float(r['Valor']):,.0f})")

    if not opciones:
        st.warning("⚠️ No hay movimientos registrados para generar comprobantes.")
    else:
        mov_sel = st.selectbox("🔍 Selecciona el movimiento:", opciones)
        if st.button("🚀 Generar Comprobante"):
            is_ing = "[INGRESO]" in mov_sel
            rec_id = f"REC-{abs(hash(mov_sel)) % 10000:04d}"
            
            if is_ing:
                fila = st.session_state.ingresos_df[st.session_state.ingresos_df.apply(lambda x: f"[INGRESO] {x['Fecha']} - {x['Concepto']} (${float(x['Valor']):,.0f})" == mov_sel, axis=1)].iloc[0]
            else:
                fila = st.session_state.gastos_df[st.session_state.gastos_df.apply(lambda x: f"[GASTO] {x['Fecha']} - {x['Concepto']} (${float(x['Valor']):,.0f})" == mov_sel, axis=1)].iloc[0]
            
            texto_recibo = f"=== COMPROBANTE OFICIAL ===\nID: {rec_id}\nFecha: {fila['Fecha']}\nConcepto: {fila['Concepto']}\nValor: ${float(fila['Valor']):,.0f} COP"
            
            qr = qrcode.QRCode(box_size=8, border=2)
            qr.add_data(texto_recibo)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            buf = BytesIO()
            img.save(buf, format="PNG")
            
            st.session_state.rec_txt = texto_recibo
            st.session_state.rec_id = rec_id
            st.session_state.rec_qr = buf.getvalue()
            st.success("¡Comprobante generado!")

    if 'rec_txt' in st.session_state:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.text_area("Comprobante", st.session_state.rec_txt, height=150)
        with c2:
            st.image(st.session_state.rec_qr, width=160)

# --- 7. REPORTE FINAL ---
elif menu == "7. Reporte Final":
    st.markdown('<p class="main-header">📑 Reporte Final del Evento</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Consolidado general y descarga</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    tot_ing = st.session_state.ingresos_df["Valor"].astype(float).sum() if not st.session_state.ingresos_df.empty else 0.0
    tot_gas = st.session_state.gastos_df["Valor"].astype(float).sum() if not st.session_state.gastos_df.empty else 0.0
    saldo = tot_ing - tot_gas
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Recaudado", f"${tot_ing:,.0f} COP")
    c2.metric("Total Gastado", f"${tot_gas:,.0f} COP")
    c3.metric("Ganancia Neta", f"${saldo:,.0f} COP", delta=f"${saldo:,.0f} COP")
    
    st.markdown("---")
    guardar_todo_en_excel()
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("⬇️ Descargar Excel Completo", data=f, file_name="Proyecto_Financiero_Eventos_Actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
