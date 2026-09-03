import streamlit as st
import pandas as pd
from datetime import datetime

# Simulación inicial del estado si no existe en la sesión
if "inventario" not in st.session_state:
    st.session_state.inventario = pd.DataFrame(
        columns=["Codigo_Barras", "Nombre_Producto", "Stock_Actual", "Unidad_Medida"]
    )

if "movimientos" not in st.session_state:
    st.session_state.movimientos = pd.DataFrame(
        columns=["Codigo_Barras", "Fecha", "Producto", "Cantidad", "Unidad_Medida", "Campo", "Cuartel", "Usuario"]
    )

# Componente para manejo de código de barras (con soporte para cámara web/móvil usando streamlit-barcode-scanner si está disponible o entrada manual inteligente)
def widget_codigo_barras(key):
    st.markdown("##### 📷 Escaneo o Ingreso de Código de Barras")
    col1, col2 = st.columns([3, 1])
    
    codigo_manual = col1.text_input("Ingrese o escanee el código aquí", key=f"input_bc_{key}", placeholder="Ej: 780123456789")
    
    # Integración opcional del componente de cámara para celulares si se usa streamlit-qrcode-scanner o streamlit-barcode-scanner
    codigo_camara = ""
    try:
        from streamlit_barcode_scanner import streamlit_barcode_scanner
        if col2.button("📱 Usar Cámara", key=f"btn_cam_{key}"):
            barcode_scan = streamlit_barcode_scanner()
            if barcode_scan:
                codigo_camara = barcode_scan
                st.success(f"¡Código detectado por cámara: {codigo_camara}!")
    except ImportError:
        col2.info("Cámara manual vía texto/pistola")

    return codigo_camara if codigo_camara else codigo_manual

# Menú de navegación principal en la barra lateral
accion = st.sidebar.selectbox(
    "Selecciona una opción",
    ["📥 Ingreso / nuevo stock", "📤 Salida / uso"]
)

# --- VISTA 1: INGRESO / NUEVO STOCK (AZUL) ---
if accion == "📥 Ingreso / nuevo stock":
    st.markdown("### 📥 Registrar ingreso de stock")
    st.markdown(
        "_Escanea el código de barras. Si el producto ya existe, se autocompletará y sumará stock; si es nuevo, podrás registrarlo._"
    )

    codigo_ingreso = widget_codigo_barras("ingreso")

    # Autoreconocimiento automático en la base de datos para Ingreso
    prod_nombre_sugerido = ""
    unidad_sugerida = "-- Seleccionar --"
    
    if codigo_ingreso:
        match_inv = st.session_state.inventario[
            st.session_state.inventario["Codigo_Barras"] == codigo_ingreso
        ]
        if not match_inv.empty:
            prod_nombre_sugerido = match_inv.iloc[0]["Nombre_Producto"]
            unidad_sugerida = match_inv.iloc[0]["Unidad_Medida"]
            st.info(f"🔍 **¡Producto reconocido en la base de datos!** Se actualizará el stock de: **{prod_nombre_sugerido}**")
        else:
            st.warning("⚠️ Código nuevo. Se registrará como un producto completamente nuevo.")

    nombre_producto = st.text_input("Nombre del producto", value=prod_nombre_sugerido)
    
    unidades_disponibles = ["Litros", "cc", "Kg", "gr", "Unidades"]
    try:
        idx_unidad = unidades_disponibles.index(unidad_sugerida) + 1
    except ValueError:
        idx_unidad = 0

    unidad_medida_ingreso = st.selectbox("Unidad de medida", ["-- Seleccionar --"] + unidades_disponibles, index=idx_unidad)

    cantidad_ingreso = st.number_input("Cantidad a ingresar")
    
    proveedor = st.text_input("Proveedor")
    
    campo_ingreso = st.selectbox("Campo", ["-- Seleccionar --", "El Huingan", "Bucalemu", "Rinconada", "San Felipe"], index=0)
    cuartel_ingreso = st.selectbox("Cuartel", ["-- Seleccionar --", "Cuartel 1 Timpson", "Cuartel 2", "Cuartel 3", "General / Bodega"], index=0)
    usuario_ingreso = st.selectbox("Usuario / Responsable", ["-- Seleccionar --", "Bruno Hernández", "Manuel Muñoz", "Bodeguero"], index=0)

    if st.button("Guardar Ingreso"):
        if codigo_ingreso and nombre_producto and unidad_medida_ingreso != "-- Seleccionar --":
            inv = st.session_state.inventario
            
            if codigo_ingreso in inv["Codigo_Barras"].values:
                idx = inv[inv["Codigo_Barras"] == codigo_ingreso].index[0]
                stock_actual = inv.at[idx, "Stock_Actual"]
                st.session_state.inventario.at[idx, "Stock_Actual"] = stock_actual + cantidad_ingreso
                st.success(f"¡Stock actualizado! Se sumaron {cantidad_ingreso} {unidad_medida_ingreso} al producto existente.")
            else:
                nuevo_prod = pd.DataFrame(
                    {
                        "Codigo_Barras": [codigo_ingreso],
                        "Nombre_Producto": [nombre_producto],
                        "Stock_Actual": [cantidad_ingreso],
                        "Unidad_Medida": [unidad_medida_ingreso],
                    }
                )
                st.session_state.inventario = pd.concat([st.session_state.inventario, nuevo_prod], ignore_index=True)
                st.success("¡Nuevo producto registrado en el inventario con éxito!")
            
            nuevo_mov = pd.DataFrame(
                {
                    "Codigo_Barras": [codigo_ingreso],
                    "Fecha": [datetime.now().strftime("%m/%d/%Y")],
                    "Producto": [nombre_producto],
                    "Cantidad": [f"+{cantidad_ingreso} {unidad_medida_ingreso}"],
                    "Unidad_Medida": [unidad_medida_ingreso],
                    "Campo": [campo_ingreso if campo_ingreso != "-- Seleccionar --" else ""],
                    "Cuartel": [cuartel_ingreso if cuartel_ingreso != "-- Seleccionar --" else ""],
                    "Usuario": [usuario_ingreso if usuario_ingreso != "-- Seleccionar --" else ""],
                }
            )
            st.session_state.movimientos = pd.concat(
                [st.session_state.movimientos, nuevo_mov], ignore_index=True
            )
        else:
            st.error("Por favor completa los campos obligatorios: Código, Nombre del producto y Unidad de medida.")

# --- VISTA 2: SALIDA / USO (VERDE) ---
elif accion == "📤 Salida / uso":
    st.markdown("### 📤 Registrar salida / uso")
    st.markdown(
        "_Escanea con el celular, pistola o selecciona el producto para el reconocimiento automático._"
    )

    lista_productos_disponibles = ["-- Seleccionar desde inventario --"] + [
        f"{row['Nombre_Producto']} (SKU: {row['Codigo_Barras']} - Stock: {row['Stock_Actual']} {row['Unidad_Medida']})"
        for _, row in st.session_state.inventario.iterrows()
    ]
    
    sel_producto = st.selectbox("📦 Selección rápida de producto", lista_productos_disponibles, index=0)
    
    codigo_sugerido_select = ""
    if sel_producto != "-- Seleccionar desde inventario --":
        import re
        match_sku = re.search(r"SKU: (.*?) -", sel_producto)
        if match_sku:
            codigo_sugerido_select = match_sku.group(1)

    codigo_salida = widget_codigo_barras("salida")
    
    if not codigo_salida and codigo_sugerido_select:
        codigo_salida = codigo_sugerido_select

    prod_nombre_encontrado = ""
    stock_disponible = 0.0
    unidad_medida_sugerida = "Litros"

    if codigo_salida:
        match = st.session_state.inventario[
            st.session_state.inventario["Codigo_Barras"] == codigo_salida
        ]
        if not match.empty:
            prod_nombre_encontrado = match.iloc[0]["Nombre_Producto"]
            stock_disponible = match.iloc[0]["Stock_Actual"]
            unidad_medida_sugerida = match.iloc[0]["Unidad_Medida"]
            st.success(f"✅ **Reconocido:** {prod_nombre_encontrado} | Stock disponible: {stock_disponible} {unidad_medida_sugerida}")
        else:
            st.warning("⚠️ Código de barras no encontrado en el inventario actual.")

    nombre_producto_salida = st.text_input("Nombre del producto", value=prod_nombre_encontrado)

    cantidad_retirada = st.number_input("Cantidad que ocuparé",)

    unidades_disponibles = ["Litros", "cc", "Kg", "gr", "Unidades"]
    try:
        index_default = unidades_disponibles.index(unidad_medida_sugerida)
    except ValueError:
        index_default = 0

    unidad_medida = st.selectbox("Unidad de medida", unidades_disponibles, index=index_default)
    
    campo = st.selectbox("Campo", ["-- Seleccionar --", "El Huingan", "Bucalemu", "Rinconada", "San Felipe"], index=0)
    cuartel = st.selectbox("Cuartel", ["-- Seleccionar --", "Cuartel 1 Timpson", "Cuartel 2", "Cuartel 3", "General / Bodega"], index=0)
    usuario = st.selectbox("Usuario / Responsable", ["-- Seleccionar --", "Bruno Hernández", "Manuel Muñoz", "Bodeguero"], index=0)

    if st.button("Guardar Salida"):
        if codigo_salida and nombre_producto_salida:
            inv = st.session_state.inventario
            if codigo_salida in inv["Codigo_Barras"].values:
                idx = inv[inv["Codigo_Barras"] == codigo_salida].index[0]
                stock_actual = inv.at[idx, "Stock_Actual"]

                if stock_actual >= cantidad_retirada:
                    st.session_state.inventario.at[idx, "Stock_Actual"] = stock_actual - cantidad_retirada
                    
                    nuevo_mov = pd.DataFrame(
                        {
                            "Codigo_Barras": [codigo_salida],
                            "Fecha": [datetime.now().strftime("%m/%d/%Y")],
                            "Producto": [nombre_producto_salida],
                            "Cantidad": [f"{cantidad_retirada} {unidad_medida}"],
                            "Unidad_Medida": [unidad_medida],
                            "Campo": [campo if campo != "-- Seleccionar --" else ""],
                            "Cuartel": [cuartel if cuartel != "-- Seleccionar --" else ""],
                            "Usuario": [usuario if usuario != "-- Seleccionar --" else ""],
                        }
                    )
                    st.session_state.movimientos = pd.concat(
                        [st.session_state.movimientos, nuevo_mov], ignore_index=True
                    )
                    st.success("¡Salida registrada con éxito! Stock actualizado en tiempo real.")
                else:
                    st.error(f"Stock insuficiente. Solo hay {stock_actual} disponibles.")
            else:
                st.error("El código ingresado no existe en el inventario actual.")
        else:
            st.error("Debe ingresar o seleccionar un producto válido para realizar la salida.")