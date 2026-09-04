from datetime import datetime
import io
import os
import re
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Servicios Agrícolas Cumbre Ltda", page_icon="🍇", layout="centered"
)

# --- ARCHIVO DE PERSISTENCIA (BASE DE DATOS COMPARTIDA) ---
DB_FILE = "inventario_cumbre.xlsx"

def inicializar_bd():
    if os.path.exists(DB_FILE):
        try:
            inventario = pd.read_excel(DB_FILE, sheet_name="Inventario_Actual")
            movimientos = pd.read_excel(DB_FILE, sheet_name="Movimientos")
            inventario["Factura_Guia"] = inventario["Factura_Guia"].fillna("").astype(str)
            return inventario, movimientos
        except Exception:
            pass
    
    # Datos iniciales base si no existe el archivo
    inventario_inicial = pd.DataFrame(
        {
            "Codigo_Barras": ["U1", "103690", "U2", "2905507", "800004005185"],
            "Factura_Guia": ["56599", "", "", "", ""],
            "Nombre_Producto": ["Urea", "Fascinate 150 sl", "Azufre", "Tebuconazol 430 sc", "Bloqueador"],
            "Tipo": ["Fertilizante", "Herbicida", "Fungicida", "Insecticida", "Insumos"],
            "Stock_Actual": [0.0, 0.0, 0.0, 0.0, 0.0],
            "Unidad_Medida": ["Kg", "Lt", "Kg", "Lt", "Unidades"],
            "Bodega": ["Huingan", "Huingan", "Huingan", "Huingan", "Huingan"],
            "Proveedor": ["Copeval", "M&Valdivieso", "Gmt", "Otro", "Copeval"],
        }
    )
    movimientos_iniciales = pd.DataFrame(
        columns=["Codigo_Barras", "Fecha", "Producto", "Cantidad", "Unidad_Medida", "Campo", "Cuartel", "Usuario"]
    )
    return inventario_inicial, movimientos_iniciales

def guardar_bd():
    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
        st.session_state.inventario.to_excel(writer, index=False, sheet_name="Inventario_Actual")
        st.session_state.movimientos.to_excel(writer, index=False, sheet_name="Movimientos")

if "inventario" not in st.session_state or "movimientos" not in st.session_state:
    inv, mov = inicializar_bd()
    st.session_state.inventario = inv
    st.session_state.movimientos = mov

# --- CABECERA ---
try:
    st.image("LOGO CUMBRE_2.jpg", use_container_width=True)
except:
    st.title("🍇 S.Agrícolas Cumbre Ltda")

st.markdown("Control inteligente de ingresos, salidas y stock en tiempo real.")

# Métricas superiores
total_productos = len(st.session_state.inventario)
sin_stock = len(
    st.session_state.inventario[st.session_state.inventario["Stock_Actual"] == 0]
)

col_m1, col_m2 = st.columns(2)
with col_m1:
    st.metric("📦 Productos", total_productos)
with col_m2:
    st.metric("⚠️ Sin stock", sin_stock)

st.markdown("---")

# --- MENÚ DE ACCIONES (CELULAR) ---
st.subheader("Menú Principal")

accion = st.radio(
    "Seleccione una operación:",
    ["📋 Ver Inventario General", "📥 Ingreso a bodega", "📤 Salida / uso"],
)

st.markdown("---")

def widget_codigo_barras(sufijo):
    cam_key = f"activar_cam_{sufijo}"
    if cam_key not in st.session_state:
        st.session_state[cam_key] = False

    codigo_input_key = f"input_codigo_{sufijo}"
    
    col_input, col_btn = st.columns([3, 1])
    
    with col_input:
        codigo_ingresado = st.text_input("Código de barras / SKU", key=codigo_input_key)
        
    with col_btn:
        st.write("") 
        st.write("")
        if st.button("📸 Cámara", key=f"btn_cam_{sufijo}", help="Usar cámara del celular"):
            st.session_state[cam_key] = not st.session_state[cam_key]

    if st.session_state[cam_key]:
        st.info("📱 Alinea el código de barras frente a la cámara y toma una foto.")
        foto = st.camera_input("Capturar código con cámara", key=f"cam_input_{sufijo}")
        if foto is not None:
            st.success("✅ ¡Foto capturada! Anota el código visualizado en la casilla superior.")
            if st.button("Cerrar cámara", key=f"cerrar_cam_{sufijo}"):
                st.session_state[cam_key] = False
                st.rerun()

    return codigo_ingresado

# --- 1. INGRESO A BODEGA ---
if accion == "📥 Ingreso a bodega":
    st.markdown("### 📥 Registrar entrada de productos")
    codigo_ingreso = widget_codigo_barras("ingreso")
    factura_guia = st.text_input("Factura o Guía de compra")
    
    prod_sugerido = ""
    tipo_sug = "-- Seleccionar --"
    unidad_sug = "-- Seleccionar --"
    
    if codigo_ingreso:
        match = st.session_state.inventario[
            st.session_state.inventario["Codigo_Barras"] == codigo_ingreso
        ]
        if not match.empty:
            prod_sugerido = match.iloc[0]["Nombre_Producto"]
            tipo_sug = match.iloc[0]["Tipo"]
            unidad_sug = match.iloc[0]["Unidad_Medida"]

    nombre = st.text_input("Nombre del producto", value=prod_sugerido)
    cantidad_ingreso = st.number_input("Cantidad a ingresar", min_value=0.0, value=None, format="%.2f")
    
    tipos_disponibles = ["-- Seleccionar --", "Fertilizante", "Herbicida", "Fungicida", "Insecticida", "Insumos", "Otros"]
    try:
        idx_tipo = tipos_disponibles.index(tipo_sug)
    except ValueError:
        idx_tipo = 0

    tipo = st.selectbox("Tipo", tipos_disponibles, index=idx_tipo)
    
    unidades_disponibles = ["-- Seleccionar --", "Kg", "gr", "Lt", "cc", "Unidades"]
    try:
        idx_unidad = unidades_disponibles.index(unidad_sug)
    except ValueError:
        idx_unidad = 0

    unidad = st.selectbox("Unidad de medida", unidades_disponibles, index=idx_unidad)
    bodega = st.selectbox("Bodega", ["-- Seleccionar --", "Huingan", "Bucalemu", "Rinconada", "San Felipe"], index=0)
    proveedor = st.selectbox("Proveedor", ["-- Seleccionar --", "Copeval", "Gmt", "M&Valdivieso", "Otro"], index=0)

    if st.button("Guardar Ingreso"):
        if codigo_ingreso and nombre and cantidad_ingreso is not None and cantidad_ingreso > 0:
            inv = st.session_state.inventario
            if codigo_ingreso in inv["Codigo_Barras"].values:
                idx = inv[inv["Codigo_Barras"] == codigo_ingreso].index[0]
                st.session_state.inventario.at[idx, "Stock_Actual"] += cantidad_ingreso
                st.session_state.inventario.at[idx, "Factura_Guia"] = factura_guia
                guardar_bd()
                st.success(f"¡Stock actualizado! Se sumaron {cantidad_ingreso} {unidad} a {nombre}.")
            else:
                nuevo_prod = pd.DataFrame(
                    {
                        "Codigo_Barras": [codigo_ingreso],
                        "Factura_Guia": [factura_guia],
                        "Nombre_Producto": [nombre],
                        "Tipo": [tipo],
                        "Stock_Actual": [cantidad_ingreso],
                        "Unidad_Medida": [unidad],
                        "Bodega": [bodega],
                        "Proveedor": [proveedor],
                    }
                )
                st.session_state.inventario = pd.concat(
                    [st.session_state.inventario, nuevo_prod], ignore_index=True
                )
                guardar_bd()
                st.success(f"¡Nuevo producto '{nombre}' registrado e ingresado con éxito!")
        else:
            st.warning("Completa los campos obligatorios y una cantidad mayor a 0.")

# --- 2. SALIDA / USO ---
elif accion == "📤 Salida / uso":
    st.markdown("### 📤 Registrar salida / aplicación")
    lista_productos_disponibles = ["-- Seleccionar desde inventario --"] + [
        f"{row['Nombre_Producto']} (SKU: {row['Codigo_Barras']} - Stock: {row['Stock_Actual']} {row['Unidad_Medida']})"
        for _, row in st.session_state.inventario.iterrows()
    ]
    
    sel_producto = st.selectbox("📦 Selección rápida de producto", lista_productos_disponibles)
    codigo_sugerido_select = ""
    if sel_producto != "-- Seleccionar desde inventario --":
        match_sku = re.search(r"SKU: (.*?) -", sel_producto)
        if match_sku:
            codigo_sugerido_select = match_sku.group(1)

    codigo_salida = widget_codigo_barras("salida")
    if not codigo_salida and codigo_sugerido_select:
        codigo_salida = codigo_sugerido_select

    prod_nombre_encontrado = ""
    unidad_medida_sugerida = "-- Seleccionar --"

    if codigo_salida:
        match = st.session_state.inventario[
            st.session_state.inventario["Codigo_Barras"] == codigo_salida
        ]
        if not match.empty:
            prod_nombre_encontrado = match.iloc[0]["Nombre_Producto"]
            unidad_medida_sugerida = match.iloc[0]["Unidad_Medida"]
        else:
            st.warning("⚠️ Código de barras no encontrado en el inventario.")

    nombre_producto_salida = st.text_input("Nombre del producto", value=prod_nombre_encontrado)
    cantidad_retirada = st.number_input("Cantidad que ocuparé", min_value=0.0, value=None, format="%.2f")

    unidades_disponibles = ["-- Seleccionar --", "Kg", "gr", "Lt", "cc", "Unidades"]
    try:
        index_default = unidades_disponibles.index(unidad_medida_sugerida)
    except ValueError:
        index_default = 0

    unidad_medida = st.selectbox("Unidad de medida", unidades_disponibles, index=index_default)
    campo = st.selectbox("Campo", ["-- Seleccionar --", "El Huingan", "Bucalemu", "Rinconada", "San Felipe"], index=0)
    cuartel = st.selectbox("Cuartel", ["-- Seleccionar --", "Cuartel 1 Timpson", "Cuartel 2", "Cuartel 3", "General / Bodega"], index=0)
    usuario = st.selectbox("Usuario / Responsable", ["-- Seleccionar --", "Bruno Hernández", "Manuel Muñoz", "Bodeguero"], index=0)

    if st.button("Guardar Salida"):
        if codigo_salida and nombre_producto_salida and cantidad_retirada is not None and cantidad_retirada > 0:
            inv = st.session_state.inventario
            if codigo_salida in inv["Codigo_Barras"].values:
                idx = inv[inv["Codigo_Barras"] == codigo_salida].index[0]
                stock_actual = inv.at[idx, "Stock_Actual"]

                if stock_actual >= cantidad_retirada:
                    st.session_state.inventario.at[idx, "Stock_Actual"] = stock_actual - cantidad_retirada
                    nuevo_mov = pd.DataFrame(
                        {
                            "Codigo_Barras": [codigo_salida],
                            "Fecha": [datetime.now().strftime("%m/%d/%Y %H:%M")],
                            "Producto": [nombre_producto_salida],
                            "Cantidad": [f"{cantidad_retirada} {unidad_medida}"],
                            "Unidad_Medida": [unidad_medida],
                            "Campo": [campo],
                            "Cuartel": [cuartel],
                            "Usuario": [usuario],
                        }
                    )
                    st.session_state.movimientos = pd.concat(
                        [st.session_state.movimientos, nuevo_mov], ignore_index=True
                    )
                    guardar_bd()
                    st.success("¡Salida registrada con éxito y respaldada en la nube!")
                else:
                    st.error(f"Stock insuficiente. Solo hay {stock_actual} disponibles.")
            else:
                st.error("El código ingresado no existe en el inventario actual.")
        else:
            st.error("Completa los datos correctamente.")

# --- 3. INVENTARIO GENERAL ---
else:
    tab1, tab2 = st.tabs(["📦 Stock Actual (Productos)", "📊 Registro de Movimientos"])
    with tab1:
        st.markdown("### Tabla de Stock en Bodega")
        buscar_prod = st.text_input("🔍 Buscar por nombre o código:")
        df_stock = st.session_state.inventario
        if buscar_prod:
            df_stock = df_stock[
                df_stock["Nombre_Producto"].str.contains(buscar_prod, case=False, na=False)
                | df_stock["Codigo_Barras"].str.contains(buscar_prod, case=False, na=False)
            ]
        st.dataframe(df_stock, use_container_width=True)

    with tab2:
        st.markdown("### Historial de Movimientos / Aplicaciones")
        st.dataframe(st.session_state.movimientos, use_container_width=True)

# --- BOTÓN PARA DESCARGAR EL EXCEL GENERAL ---
st.markdown("---")
st.subheader("📊 Exportar Datos")

if os.path.exists(DB_FILE):
    with open(DB_FILE, "rb") as f:
        excel_bytes = f.read()
else:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.inventario.to_excel(writer, index=False, sheet_name='Inventario_Actual')
        st.session_state.movimientos.to_excel(writer, index=False, sheet_name='Movimientos')
    excel_bytes = output.getvalue()

st.download_button(
    label="📥 Descargar Base de Datos Completa (Excel)",
    data=excel_bytes,
    file_name="Inventario_Servicios_Cumbre.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)