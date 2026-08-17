import streamlit as st
import pandas as pd
from pydantic import BaseModel, Field
from typing import List
from google import genai
from google.genai import types
import os
from datetime import datetime

# Configuración de página
st.set_page_config(page_title="Registro de Comprobantes - Obra", layout="wide", page_icon="🏗️")

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

# 2. Configurar el Cliente de Gemini
API_KEY = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))

if not API_KEY:
    st.warning("⚠️ Configura tu GEMINI_API_KEY en .streamlit/secrets.toml")
    st.stop()

client = genai.Client(api_key=API_KEY)

# 3. Función de extracción compatible con PDF e Imágenes
def extraer_datos_comprobante(file_bytes: bytes, mime_type: str) -> FacturaData:
    prompt = """
    Analiza detalladamente este comprobante de pago de construcción o materiales (factura, boleta o guía).
    Extrae con máxima precisión la información del encabezado y la tabla completa de ítems.
    Verifica que:
    - Cantidad sea numérica.
    - Precio parcial = Cantidad * Precio Unitario (o el subtotal impreso en el documento).
    """
    
    # Crear el objeto de archivo compatible con PDF o Imagen
    archivo_part = types.Part.from_bytes(
        data=file_bytes,
        mime_type=mime_type
    )
    
    # Modelos vigentes con soporte nativo de PDF y Visión
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
            
    raise RuntimeError(f"Error procesando documento: {str(ultimo_error)}")

# 4. Gestión del Archivo Madre en Excel
EXCEL_PATH = "registro_madre_comprobantes.xlsx"

def guardar_en_archivo_madre(datos: FacturaData, proyecto: str):
    filas = []
    fecha_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for item in datos.items:
        filas.append({
            "Fecha_Registro_Sistema": fecha_registro,
            "Proyecto": proyecto,
            "Tipo_Comprobante": datos.tipo_comprobante,
            "Nro_Comprobante": datos.numero_comprobante,
            "Empresa_Proveedor": datos.nombre_empresa,
            "RUC_Proveedor": datos.ruc_empresa,
            "Fecha_Emision": datos.fecha_emision,
            "Moneda": datos.moneda,
            "Descripcion_Material": item.descripcion_material,
            "Cantidad": item.cantidad,
            "Unidad": item.unidad_medida,
            "Precio_Unitario": item.precio_unitario,
            "Precio_Parcial": item.precio_parcial,
            "Total_Comprobante": datos.monto_total
        })
    
    df_nuevo = pd.DataFrame(filas)
    
    if os.path.exists(EXCEL_PATH):
        df_existente = pd.read_excel(EXCEL_PATH)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
        
    df_final.to_excel(EXCEL_PATH, index=False)
    return df_final

# 5. Interfaz de Usuario
st.title("🏗️ Registro Inteligente de Facturas y Materiales")
st.caption("Captura de comprobantes (PDF, JPG, PNG) con extracción automática y consolidación al archivo madre.")

with st.sidebar:
    st.header("⚙️ Parámetros de Obra")
    proyecto_seleccionado = st.text_input("Nombre del Proyecto / Obra", value="Proyecto Nuevo")
    st.divider()
    if os.path.exists(EXCEL_PATH):
        df_madre = pd.read_excel(EXCEL_PATH)
        st.metric("Total Registros Guardados", len(df_madre))
        st.metric("Gasto Total Acumulado", f"S/ {df_madre['Precio_Parcial'].sum():,.2f}")
        with open(EXCEL_PATH, "rb") as f:
            st.download_button("📥 Descargar Archivo Madre (.xlsx)", f, file_name=EXCEL_PATH)

# Entrada de archivos (PDF o Imágenes)
col_input1, col_input2 = st.columns(2)
with col_input1:
    archivo_subido = st.file_uploader("Subir Factura/Boleta (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg", "webp"])
with col_input2:
    foto_camara = st.camera_input("O tomar foto con la cámara (Móvil/Webcam)")

archivo_activo = archivo_subido if archivo_subido is not None else foto_camara

if archivo_activo:
    # Determinar tipo MIME
    mime_type = archivo_activo.type
    if not mime_type:
        if archivo_activo.name.lower().endswith(".pdf"):
            mime_type = "application/pdf"
        elif archivo_activo.name.lower().endswith(".png"):
            mime_type = "image/png"
        else:
            mime_type = "image/jpeg"
            
    col_prev, col_proc = st.columns([1, 2])
    with col_prev:
        if "pdf" in mime_type:
            st.success(f"📄 Archivo PDF cargado: **{archivo_activo.name}**")
        else:
            st.image(archivo_activo, caption="Comprobante cargado", use_container_width=True)
            
        btn_procesar = st.button("🔍 Extraer Datos con IA", type="primary", use_container_width=True)
    
    if btn_procesar:
        with st.spinner("Procesando documento con IA y cuadrando importes..."):
            try:
                # Leer bytes del archivo subido
                file_bytes = archivo_activo.getvalue()
                datos_extraidos = extraer_datos_comprobante(file_bytes, mime_type)
                st.session_state["datos_temp"] = datos_extraidos
                st.success("¡Datos extraídos exitosamente!")
            except Exception as e:
                st.error(f"Error en la extracción: {str(e)}")

# Validación previa antes de confirmar en el archivo madre
if "datos_temp" in st.session_state:
    datos: FacturaData = st.session_state["datos_temp"]
    
    st.divider()
    st.subheader("📋 Validación y Confirmación de Datos")
    
    col_c1, col_c2, col_c3, col_c4 = st.columns(4)
    col_c1.text_input("Tipo", value=datos.tipo_comprobante, key="val_tipo")
    col_c2.text_input("N° Comprobante", value=datos.numero_comprobante, key="val_nro")
    col_c3.text_input("Proveedor", value=datos.nombre_empresa, key="val_empresa")
    col_c4.text_input("RUC", value=datos.ruc_empresa, key="val_ruc")
    
    # Preparar DataFrame editable para los ítems
    items_list = [item.model_dump() for item in datos.items]
    df_items = pd.DataFrame(items_list)
    
    st.markdown("**Detalle de Materiales e Insumos:**")
    df_editado = st.data_editor(
        df_items,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "precio_unitario": st.column_config.NumberColumn(format="S/ %.2f"),
            "precio_parcial": st.column_config.NumberColumn(format="S/ %.2f"),
            "cantidad": st.column_config.NumberColumn(format="%.2f"),
        }
    )
    
    # Recalcular parciales si el usuario cambió valores en la tabla
    df_editado["precio_parcial"] = df_editado["cantidad"] * df_editado["precio_unitario"]
    total_recalculado = df_editado["precio_parcial"].sum()
    st.markdown(f"**Total Calculado Ítems:** `S/ {total_recalculado:,.2f}` | **Total Comprobante:** `S/ {datos.monto_total:,.2f}`")
    
    if st.button("💾 Confirmar y Registrar en Archivo Madre", type="primary"):
        nuevos_items = [ItemComprobante(**row) for row in df_editado.to_dict(orient="records")]
        datos.items = nuevos_items
        datos.monto_total = total_recalculado
        datos.tipo_comprobante = st.session_state["val_tipo"]
        datos.numero_comprobante = st.session_state["val_nro"]
        datos.nombre_empresa = st.session_state["val_empresa"]
        datos.ruc_empresa = st.session_state["val_ruc"]
        
        guardar_en_archivo_madre(datos, proyecto_seleccionado)
        st.success("✅ Registro guardado en el archivo madre con éxito.")
        del st.session_state["datos_temp"]
        st.rerun()