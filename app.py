import streamlit as st
import pandas as pd
from pydantic import BaseModel, Field
from typing import List
from google import genai
from google.genai import types
import io
import os
from datetime import datetime

# Configuración de página
st.set_page_config(
    page_title="Gestión de Compras y Facturas - Obra",
    layout="wide",
    page_icon="🏗️"
)

# Carpeta para almacenamiento físico de sustentación
FOLDER_SUSTENTOS = "sustentos"
os.makedirs(FOLDER_SUSTENTOS, exist_ok=True)

# 1. Esquema estructurado de salida (Pydantic)
class ItemComprobante(BaseModel):
    descripcion_material: str = Field(description="Nombre o descripcion detallada del material o insumo")
    cantidad: float = Field(description="Cantidad adquirida")
    unidad_medida: str = Field(description="Unidad de medida (ej. BOL, VAR, M3, KG, GLN, UND)")
    precio_unitario: float = Field(description="Precio por unidad")
    precio_parcial: float = Field(description="Subtotal o precio parcial del item")

class FacturaData(BaseModel):
    tipo_comprobante: str = Field(description="Factura, Boleta de Venta, Guia de Remision o Nota de Venta")
    numero_comprobante: str = Field(description="Serie y numero del documento (ej. F001-00012345)")
    nombre_empresa: str = Field(description="Razon Social o Nombre Comercial del proveedor")
    ruc_empresa: str = Field(description="RUC o documento de identidad del proveedor")
    fecha_emision: str = Field(description="Fecha de emision en formato YYYY-MM-DD o vacio")
    moneda: str = Field(description="PEN o USD")
    items: List[ItemComprobante]
    monto_total: float = Field(description="Importe total del comprobante")

# 2. Configuración de API Key
API_KEY = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))

if not API_KEY:
    st.warning("⚠️ Configura tu GEMINI_API_KEY en .streamlit/secrets.toml o en Streamlit Secrets")
    st.stop()

client = genai.Client(api_key=API_KEY)

# 3. Función de extracción con IA usando los modelos vigentes
def extraer_datos_comprobante(file_bytes: bytes, mime_type: str) -> FacturaData:
    prompt = """
    Analiza detalladamente este comprobante de pago de construcción o materiales (factura, boleta o nota).
    Extrae con máxima precisión la información del encabezado y la tabla completa de ítems.
    Verifica que:
    - Cantidad sea numérica.
    - Precio parcial = Cantidad * Precio Unitario (o el subtotal impreso en el documento).
    - Moneda sea PEN o USD.
    """
    archivo_part = types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
    
    modelos_a_probar = [
        "gemini-3.7-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.1-flash-lite"
    ]
    
    ultimo_error = None
    for mod in modelos_a_probar:
        try:
            response = client.models.generate_content(
                model=mod,
                contents=[prompt, archivo_part],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=FacturaData,
                    temperature=0.1
                )
            )
            return FacturaData.model_validate_json(response.text)
        except Exception as e:
            ultimo_error = e
            continue
            
    raise RuntimeError(f"Error al procesar con IA: {str(ultimo_error)}")

# 4. Gestión del Archivo Madre (Con migración automática de columnas)
EXCEL_PATH = "registro_madre_comprobantes.xlsx"
COLUMNAS_MADRE = [
    "ID_Registro", "Fecha_Hora_Registro", "Proyecto", "Tipo_Comprobante",
    "Nro_Comprobante", "Empresa_Proveedor", "RUC_Proveedor",
    "Fecha_Emision", "Moneda", "Descripcion_Material", "Cantidad",
    "Unidad", "Precio_Unitario", "Precio_Parcial", "Total_Comprobante", "Ruta_Sustento"
]

def cargar_archivo_madre() -> pd.DataFrame:
    if os.path.exists(EXCEL_PATH):
        try:
            df = pd.read_excel(EXCEL_PATH)
            # Asegurar que todas las columnas requeridas existan
            for col in COLUMNAS_MADRE:
                if col not in df.columns:
                    df[col] = ""
            return df[COLUMNAS_MADRE]
        except Exception:
            pass
    return pd.DataFrame(columns=COLUMNAS_MADRE)

def guardar_en_archivo_madre(datos: FacturaData, proyecto: str, ruta_sustento: str):
    df_actual = cargar_archivo_madre()
    fecha_reg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    filas = []
    
    nuevo_id_base = len(df_actual) + 1
    for idx, item in enumerate(datos.items):
        filas.append({
            "ID_Registro": f"REG-{nuevo_id_base + idx:04d}",
            "Fecha_Hora_Registro": fecha_reg,
            "Proyecto": proyecto,
            "Tipo_Comprobante": datos.tipo_comprobante,
            "Nro_Comprobante": datos.numero_comprobante,
            "Empresa_Proveedor": datos.nombre_empresa,
            "RUC_Proveedor": datos.ruc_empresa,
            "Fecha_Emision": datos.fecha_emision if datos.fecha_emision else datetime.now().strftime("%Y-%m-%d"),
            "Moneda": datos.moneda,
            "Descripcion_Material": item.descripcion_material,
            "Cantidad": item.cantidad,
            "Unidad": item.unidad_medida,
            "Precio_Unitario": item.precio_unitario,
            "Precio_Parcial": item.precio_parcial,
            "Total_Comprobante": datos.monto_total,
            "Ruta_Sustento": ruta_sustento
        })
    
    df_nuevo = pd.DataFrame(filas)
    df_final = pd.concat([df_actual, df_nuevo], ignore_index=True)
    df_final.to_excel(EXCEL_PATH, index=False)
    return df_final

# 5. Barra Lateral
with st.sidebar:
    st.header("⚙️ Control de Obra")
    proyecto_seleccionado = st.text_input("Proyecto Activo", value="Proyecto Residencial")
    st.divider()
    
    df_madre = cargar_archivo_madre()
    st.metric("Total Líneas Registradas", len(df_madre))
    if not df_madre.empty:
        total_acumulado = pd.to_numeric(df_madre["Precio_Parcial"], errors="coerce").fillna(0).sum()
        st.metric("Gasto Total Acumulado", f"S/ {total_acumulado:,.2f}")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_madre.to_excel(writer, index=False, sheet_name="Archivo_Madre")
        st.download_button(
            label="📥 Descargar Archivo Madre (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"Archivo_Madre_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# 6. Estructura de Pestañas
tab_registro, tab_dashboard = st.tabs(["📥 Registro & Escaneo", "📊 Dashboard de Compras"])

# ==========================================
# PESTAÑA 1: REGISTRO Y ESCANEO
# ==========================================
with tab_registro:
    st.subheader("Captura y Almacenamiento de Comprobantes")
    st.caption("Carga tus documentos (PDF, JPG, PNG). Se almacenará el sustento físico y se verificarán duplicados.")
    
    archivo_subido = st.file_uploader("📂 Seleccionar comprobante (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg", "webp"])

    if archivo_subido is not None:
        mime_type = archivo_subido.type
        if not mime_type:
            if archivo_subido.name.lower().endswith(".pdf"):
                mime_type = "application/pdf"
            elif archivo_subido.name.lower().endswith(".png"):
                mime_type = "image/png"
            else:
                mime_type = "image/jpeg"

        col_prev, col_act = st.columns([1, 2])
        with col_prev:
            if "pdf" in mime_type:
                st.info(f"📄 **Documento PDF:** {archivo_subido.name}")
            else:
                st.image(archivo_subido, caption="Comprobante cargado", use_container_width=True)
            
            btn_extraer = st.button("🔍 Extraer Datos con IA", type="primary", use_container_width=True)

        if btn_extraer:
            with st.spinner("Analizando documento y extrayendo datos..."):
                try:
                    bytes_data = archivo_subido.getvalue()
                    datos_extraidos = extraer_datos_comprobante(bytes_data, mime_type)
                    
                    # Guardar archivo físicamente en carpeta sustentos
                    nombre_archivo_sustento = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{archivo_subido.name}"
                    ruta_guardada = os.path.join(FOLDER_SUSTENTOS, nombre_archivo_sustento)
                    with open(ruta_guardada, "wb") as f:
                        f.write(bytes_data)
                    
                    st.session_state["datos_temp"] = datos_extraidos
                    st.session_state["ruta_sustento_temp"] = ruta_guardada
                    st.success("¡Extracción completada!")
                except Exception as e:
                    st.error(f"Error en la extracción: {str(e)}")

    # Sección de Validación y Detección de Duplicados
    if "datos_temp" in st.session_state:
        datos: FacturaData = st.session_state["datos_temp"]
        st.divider()
        st.subheader("📋 Validación de Encabezado y Verificación de Serie")
        
        # VERIFICACIÓN DE DUPLICADOS EN BASE DE DATOS
        df_existente = cargar_archivo_madre()
        num_doc = datos.numero_comprobante.strip()
        
        if not df_existente.empty and "Nro_Comprobante" in df_existente.columns:
            coincidencias = df_existente[df_existente["Nro_Comprobante"].astype(str).str.strip().str.upper() == num_doc.upper()]
            if not coincidencias.empty:
                fecha_prev = coincidencias.iloc[0]["Fecha_Hora_Registro"]
                proveedor_prev = coincidencias.iloc[0]["Empresa_Proveedor"]
                st.error(f"⚠️ **¡ALERTA DE DUPLICADO!** La factura/boleta con Serie/N° **{num_doc}** ya fue registrada anteriormente en el sistema el {fecha_prev} para el proveedor **{proveedor_prev}**.")
            else:
                st.success(f"✅ N° Comprobante **{num_doc}** verificado. Es un documento nuevo.")
        
        c1, c2, c3, c4 = st.columns(4)
        v_tipo = c1.text_input("Tipo de Comprobante", value=datos.tipo_comprobante)
        v_nro = c2.text_input("N° Comprobante (Serie)", value=datos.numero_comprobante)
        v_empresa = c3.text_input("Proveedor / Empresa", value=datos.nombre_empresa)
        v_ruc = c4.text_input("RUC / ID", value=datos.ruc_empresa)
        
        items_dict = [it.model_dump() for it in datos.items]
        df_items = pd.DataFrame(items_dict)
        
        st.markdown("**Desglose de Partidas y Materiales:**")
        df_editado = st.data_editor(
            df_items,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "descripcion_material": st.column_config.TextColumn("Descripción del Material", required=True),
                "cantidad": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                "unidad_medida": st.column_config.TextColumn("Unidad"),
                "precio_unitario": st.column_config.NumberColumn("Precio Unitario", format="S/ %.2f"),
                "precio_parcial": st.column_config.NumberColumn("Precio Parcial", format="S/ %.2f"),
            }
        )
        
        df_editado["precio_parcial"] = df_editado["cantidad"] * df_editado["precio_unitario"]
        total_calc = df_editado["precio_parcial"].sum()
        
        st.markdown(f"**Total Ítems:** `S/ {total_calc:,.2f}` | **Total Comprobante:** `S/ {datos.monto_total:,.2f}`")
        
        if st.button("💾 Confirmar y Registrar en Archivo Madre", type="primary", use_container_width=True):
            datos.items = [ItemComprobante(**row) for row in df_editado.to_dict(orient="records")]
            datos.monto_total = total_calc
            datos.tipo_comprobante = v_tipo
            datos.numero_comprobante = v_nro
            datos.nombre_empresa = v_empresa
            datos.ruc_empresa = v_ruc
            
            ruta_sustento = st.session_state.get("ruta_sustento_temp", "")
            guardar_en_archivo_madre(datos, proyecto_seleccionado, ruta_sustento)
            
            st.success("✅ Documento y partidas guardados en el Archivo Madre con éxito.")
            del st.session_state["datos_temp"]
            if "ruta_sustento_temp" in st.session_state:
                del st.session_state["ruta_sustento_temp"]
            st.rerun()

# ==========================================
# PESTAÑA 2: DASHBOARD DE COMPRAS
# ==========================================
with tab_dashboard:
    st.subheader("📊 Resumen Ejecutivo y Control Semanal de Gastos")
    
    df_madre = cargar_archivo_madre()
    
    # Asegurar conversión numérica para métricas
    df_madre["Precio_Parcial"] = pd.to_numeric(df_madre["Precio_Parcial"], errors="coerce").fillna(0)
    
    if df_madre.empty or df_madre["ID_Registro"].isnull().all() or len(df_madre) == 0:
        st.info("ℹ️ Aún no hay compras registradas. Registra tus primeros comprobantes en la pestaña 'Registro & Escaneo'.")
    else:
        # KPIs
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        total_gastado = df_madre["Precio_Parcial"].sum()
        total_facturas = df_madre["Nro_Comprobante"].nunique()
        total_proveedores = df_madre["Empresa_Proveedor"].nunique()
        total_materiales = len(df_madre)
        
        kpi1.metric("Gasto Total (S/)", f"S/ {total_gastado:,.2f}")
        kpi2.metric("Comprobantes Registrados", f"{total_facturas} docs")
        kpi3.metric("Proveedores Distintos", f"{total_proveedores}")
        kpi4.metric("Ítems / Partidas", f"{total_materiales} líneas")
        
        st.divider()
        
        # Histograma de Gastos Semanales + Top Proveedores
        g_col1, g_col2 = st.columns(2)
        
        with g_col1:
            st.markdown("##### 📅 Histograma de Gastos Semanales del Proyecto")
            df_madre["Fecha_DT"] = pd.to_datetime(df_madre["Fecha_Emision"], errors="coerce")
            df_madre["Año_Semana"] = df_madre["Fecha_DT"].dt.strftime('%Y-W%U').fillna("Sin Fecha")
            
            gastos_semanales = df_madre.groupby("Año_Semana")["Precio_Parcial"].sum()
            st.bar_chart(gastos_semanales, color="#27AE60")
            
        with g_col2:
            st.markdown("##### 🏢 Gasto por Proveedor (Top 10)")
            gasto_prov = df_madre.groupby("Empresa_Proveedor")["Precio_Parcial"].sum().sort_values(ascending=False).head(10)
            st.bar_chart(gasto_prov, color="#2E86C1")
        
        st.divider()
        
        # Tabla del Archivo Madre
        st.markdown("##### 🗃️ Archivo Madre Consolidado")
        
        st.dataframe(
            df_madre[COLUMNAS_MADRE],
            use_container_width=True,
            hide_index=True
        )
