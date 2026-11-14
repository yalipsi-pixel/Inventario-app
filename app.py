import os
from datetime import datetime

import pandas as pd
import streamlit as st
from PIL import Image
from pyzbar import pyzbar

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA Y LOGO
# ==========================================
st.set_page_config(
    page_title="Inventario - Servicios Agrícolas Cumbre",
    page_icon="📦",
    layout="wide",
)

ruta_logo = "LOGO CUMBRE_2.jpg"
if os.path.exists(ruta_logo):
    st.image(ruta_logo, width=250)
else:
    st.warning(
        f"⚠️ No se encontró el archivo '{ruta_logo}'. Por favor, colócalo en la misma carpeta que este programa para visualizarlo."
    )

st.title("📦 Sistema de Inventario Web")
st.subheader("Servicios Agrícolas Cumbre Ltda.")

home_dir = os.path.expanduser("~")
ruta_escritorio = os.path.join(home_dir, "Desktop")
if not os.path.exists(ruta_escritorio):
    ruta_escritorio = os.path.join(home_dir, "Escritorio")

CARPETA_BBDD = os.path.join(ruta_escritorio, "inventario cumbre")
EXCEL_FILE = os.path.join(CARPETA_BBDD, "Inventario_Servicios_Cumbre.xlsx")

CATEGORIAS = ["Herbicida", "Insecticida", "Fertilizante", "Fungicida", "Otros"]
UNIDADES = ["Lt", "cc", "Kg", "gr", "Unidad"]
BODEGAS = ["Huimgal", "San Felipe"]
PROVEEDORES = ["M&Valdivieso", "Copeval", "Gmt", "Otros"]
CAMPOS = ["San Felipe", "Santa Maria", "Rinconada", "Bucalemu", "Rinconada Galvez"]
RESPONSABLES = ["Manuel Muñoz", "Bruno Hernandez", "Otro"]


def inicializar_entorno():
    os.makedirs(CARPETA_BBDD, exist_ok=True)

    if not os.path.exists(EXCEL_FILE):
        df_inv = pd.DataFrame(columns=[
            "Código",
            "Código_Barra",
            "Producto",
            "Categoría",
            "Stock_Actual",
            "Unidad_Medida",
            "Bodega",
            "Ultima_Actualizacion",
        ])
        df_catalogo = pd.DataFrame(columns=[
            "Código_Barra",
            "Producto",
            "Categoría",
            "Unidad_Medida",
            "Proveedor",
            "Bodega",
        ])
        df_mov = pd.DataFrame(columns=[
            "Fecha",
            "Tipo_Movimiento",
            "Código_Barra",
            "Producto",
            "Cantidad",
            "Unidad_Medida",
            "Categoría",
            "Factura_Guia",
            "Bodega",
            "Proveedor",
            "Campo",
            "Variedad",
            "Cuartel",
            "Responsable",
            "Observaciones",
        ])

        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_inv.to_excel(writer, sheet_name="Inventario_Actual", index=False)
            df_catalogo.to_excel(writer, sheet_name="Catalogo_Productos", index=False)
            df_mov.to_excel(writer, sheet_name="Movimientos", index=False)


def asegurar_columnas(df, columnas):
    for col in columnas:
        if col not in df.columns:
            df[col] = ""
    return df


def cargar_datos():
    inicializar_entorno()

    df_inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventario_Actual")
    df_catalogo = pd.read_excel(EXCEL_FILE, sheet_name="Catalogo_Productos")
    df_mov = pd.read_excel(EXCEL_FILE, sheet_name="Movimientos")

    df_inv = asegurar_columnas(df_inv, [
        "Código",
        "Código_Barra",
        "Producto",
        "Categoría",
        "Stock_Actual",
        "Unidad_Medida",
        "Bodega",
        "Ultima_Actualizacion",
    ])
    df_catalogo = asegurar_columnas(df_catalogo, [
        "Código_Barra",
        "Producto",
        "Categoría",
        "Unidad_Medida",
        "Proveedor",
        "Bodega",
    ])
    df_mov = asegurar_columnas(df_mov, [
        "Fecha",
        "Tipo_Movimiento",
        "Código_Barra",
        "Producto",
        "Cantidad",
        "Unidad_Medida",
        "Categoría",
        "Factura_Guia",
        "Bodega",
        "Proveedor",
        "Campo",
        "Variedad",
        "Cuartel",
        "Responsable",
        "Observaciones",
    ])

    return df_inv, df_catalogo, df_mov


def guardar_datos(df_inv, df_catalogo, df_mov):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="Inventario_Actual", index=False)
        df_catalogo.to_excel(writer, sheet_name="Catalogo_Productos", index=False)
        df_mov.to_excel(writer, sheet_name="Movimientos", index=False)


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def leer_codigo_desde_imagen(imagen):
    if imagen is None:
        return ""

    try:
        image = Image.open(imagen)
        codigos = pyzbar.decode(image)
        if not codigos:
            st.warning("⚠️ No se detectó ningún código de barras en la imagen. Intenta otra foto.")
            return ""
        return codigos[0].data.decode("utf-8", errors="ignore")
    except Exception:
        st.warning("⚠️ No se pudo leer el código de barras. Intenta nuevamente.")
        return ""


def buscar_producto_por_codigo(codigo, df_inv, df_catalogo):
    codigo_norm = limpiar_texto(codigo).upper()
    if not codigo_norm:
        return None

    producto_en_inventario = df_inv[df_inv["Código_Barra"].astype(str).str.upper().str.strip() == codigo_norm]
    if not producto_en_inventario.empty:
        return producto_en_inventario.iloc[0].to_dict()

    producto_en_catalogo = df_catalogo[df_catalogo["Código_Barra"].astype(str).str.upper().str.strip() == codigo_norm]
    if not producto_en_catalogo.empty:
        return producto_en_catalogo.iloc[0].to_dict()

    return None


def obtener_siguiente_codigo(df_inv):
    codigo_actuales = pd.to_numeric(df_inv["Código"].astype(str).str.extract(r"(\d+)")[0], errors="coerce").fillna(0)
    ultimo = int(codigo_actuales.max()) if not codigo_actuales.empty and codigo_actuales.notna().any() else 0
    return f"PROD-{ultimo + 1:04d}"


def actualizar_catalogo(df_catalogo, codigo, nombre, categoria, unidad, proveedor, bodega):
    if not limpiar_texto(codigo):
        return df_catalogo

    codigo_norm = limpiar_texto(codigo).upper()
    nombre_limpio = limpiar_texto(nombre)

    registro = {
        "Código_Barra": codigo_norm,
        "Producto": nombre_limpio,
        "Categoría": categoria,
        "Unidad_Medida": unidad,
        "Proveedor": proveedor,
        "Bodega": bodega,
    }

    df_catalogo = df_catalogo.copy()
    idx = df_catalogo[df_catalogo["Código_Barra"].astype(str).str.upper().str.strip() == codigo_norm].index
    if idx.empty:
        df_catalogo = pd.concat([df_catalogo, pd.DataFrame([registro])], ignore_index=True)
    else:
        for columna, valor in registro.items():
            df_catalogo.at[idx[0], columna] = valor

    return df_catalogo


# CARGA INICIAL

df_inv, df_catalogo, df_mov = cargar_datos()

# ==========================================
# MENÚ DE NAVEGACIÓN (PESTAÑAS)
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(["➕ Entrada", "➖ Salida", "📦 Inventario", "📊 Historial"])

# --- PESTAÑA 1: ENTRADA ---
with tab1:
    st.markdown("### Registro de Entrada")

    with st.form("form_entrada", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            codigo_barras_entrada = st.text_input(
                "Código de barra",
                value=st.session_state.get("codigo_entrada", ""),
                key="codigo_entrada",
                help="Puedes escanear con la cámara del celular o escribir el código manualmente.",
            )
            foto_entrada = st.camera_input("📷 Escanear código de barras", key="camara_entrada")
            if foto_entrada is not None:
                codigo_leido = leer_codigo_desde_imagen(foto_entrada)
                if codigo_leido:
                    st.session_state["codigo_entrada"] = codigo_leido
                    codigo_barras_entrada = codigo_leido

            fecha_ingreso = st.date_input("Fecha de ingreso", datetime.now().date(), disabled=True)
            factura_guia = st.text_input("Número de factura / guía")
            nombre_producto = st.text_input("Nombre del producto")
            categoria_entrada = st.selectbox("Categoría", CATEGORIAS)

        with col2:
            cantidad_entrada = st.number_input("Cantidad", min_value=0.0, step=0.01, format="%.2f")
            unidad_entrada = st.selectbox("Unidad de medida", UNIDADES)
            bodega_entrada = st.selectbox("Bodega", BODEGAS)
            proveedor_entrada = st.selectbox("Proveedor", PROVEEDORES)
            observaciones_entrada = st.text_area("Observaciones", height=110)

        submit_entrada = st.form_submit_button("Guardar Entrada")

        if submit_entrada:
            codigo_barras_entrada = limpiar_texto(codigo_barras_entrada).upper()
            nombre_producto = limpiar_texto(nombre_producto)

            if not nombre_producto:
                st.error("⚠️ Debes ingresar el nombre del producto.")
            elif cantidad_entrada <= 0:
                st.error("⚠️ La cantidad debe ser mayor que cero.")
            else:
                producto_encontrado = buscar_producto_por_codigo(codigo_barras_entrada, df_inv, df_catalogo)
                if producto_encontrado is not None:
                    nombre_producto = limpiar_texto(producto_encontrado.get("Producto", nombre_producto))
                    categoria_entrada = producto_encontrado.get("Categoría", categoria_entrada)
                    unidad_entrada = producto_encontrado.get("Unidad_Medida", unidad_entrada)
                    bodega_entrada = producto_encontrado.get("Bodega", bodega_entrada)

                fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                prod_key = nombre_producto.upper()

                if producto_encontrado is not None:
                    producto_existente = df_inv[
                        df_inv["Producto"].astype(str).str.upper().str.strip() == prod_key
                    ]
                    if not producto_existente.empty:
                        idx = producto_existente.index[0]
                        stock_previo = float(df_inv.at[idx, "Stock_Actual"] if not pd.isna(df_inv.at[idx, "Stock_Actual"]) else 0)
                        if df_inv.at[idx, "Unidad_Medida"] != unidad_entrada:
                            st.error(
                                f"⚠️ El producto ya está registrado en {df_inv.at[idx, 'Unidad_Medida']}. Usa la misma unidad de medida."
                            )
                        else:
                            df_inv.at[idx, "Stock_Actual"] = stock_previo + float(cantidad_entrada)
                            df_inv.at[idx, "Ultima_Actualizacion"] = fecha_actual
                            df_inv.at[idx, "Bodega"] = bodega_entrada
                            df_inv.at[idx, "Categoría"] = categoria_entrada
                            df_inv.at[idx, "Código_Barra"] = codigo_barras_entrada or df_inv.at[idx, "Código_Barra"]
                    else:
                        nuevo_registro = pd.DataFrame([{
                            "Código": obtener_siguiente_codigo(df_inv),
                            "Código_Barra": codigo_barras_entrada,
                            "Producto": nombre_producto,
                            "Categoría": categoria_entrada,
                            "Stock_Actual": cantidad_entrada,
                            "Unidad_Medida": unidad_entrada,
                            "Bodega": bodega_entrada,
                            "Ultima_Actualizacion": fecha_actual,
                        }])
                        df_inv = pd.concat([df_inv, nuevo_registro], ignore_index=True)
                else:
                    nuevo_registro = pd.DataFrame([{
                        "Código": obtener_siguiente_codigo(df_inv),
                        "Código_Barra": codigo_barras_entrada,
                        "Producto": nombre_producto,
                        "Categoría": categoria_entrada,
                        "Stock_Actual": cantidad_entrada,
                        "Unidad_Medida": unidad_entrada,
                        "Bodega": bodega_entrada,
                        "Ultima_Actualizacion": fecha_actual,
                    }])
                    df_inv = pd.concat([df_inv, nuevo_registro], ignore_index=True)

                df_catalogo = actualizar_catalogo(
                    df_catalogo,
                    codigo_barras_entrada,
                    nombre_producto,
                    categoria_entrada,
                    unidad_entrada,
                    proveedor_entrada,
                    bodega_entrada,
                )

                nuevo_mov = pd.DataFrame([{
                    "Fecha": fecha_actual,
                    "Tipo_Movimiento": "Entrada",
                    "Código_Barra": codigo_barras_entrada,
                    "Producto": nombre_producto,
                    "Cantidad": cantidad_entrada,
                    "Unidad_Medida": unidad_entrada,
                    "Categoría": categoria_entrada,
                    "Factura_Guia": factura_guia,
                    "Bodega": bodega_entrada,
                    "Proveedor": proveedor_entrada,
                    "Campo": "",
                    "Variedad": "",
                    "Cuartel": "",
                    "Responsable": "",
                    "Observaciones": observaciones_entrada,
                }])
                df_mov = pd.concat([df_mov, nuevo_mov], ignore_index=True)

                guardar_datos(df_inv, df_catalogo, df_mov)
                st.success(f"✅ Entrada registrada correctamente: {nombre_producto} ({cantidad_entrada:g} {unidad_entrada}).")
                st.rerun()

# --- PESTAÑA 2: SALIDA ---
with tab2:
    st.markdown("### Registro de Salida")

    with st.form("form_salida", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            codigo_barras_salida = st.text_input(
                "Código de barra",
                value=st.session_state.get("codigo_salida", ""),
                key="codigo_salida",
                help="Puedes escanear con la cámara del celular o escribir el código manualmente.",
            )
            foto_salida = st.camera_input("📷 Escanear código de barras", key="camara_salida")
            if foto_salida is not None:
                codigo_leido = leer_codigo_desde_imagen(foto_salida)
                if codigo_leido:
                    st.session_state["codigo_salida"] = codigo_leido
                    codigo_barras_salida = codigo_leido

            fecha_salida = st.date_input("Fecha de salida", datetime.now().date(), disabled=True)
            nombre_producto_salida = st.text_input("Nombre del producto")
            categoria_salida = st.selectbox("Categoría", CATEGORIAS)

        with col2:
            cantidad_salida = st.number_input("Cantidad", min_value=0.0, step=0.01, format="%.2f")
            unidad_salida = st.selectbox("Unidad de medida", UNIDADES)
            campo_salida = st.selectbox("Campo", CAMPOS)
            variedad_salida = st.text_input("Variedad")
            cuartel_salida = st.text_input("Cuartel")
            responsable_salida = st.selectbox("Responsable", RESPONSABLES)
            observaciones_salida = st.text_area("Observaciones", height=110)

        submit_salida = st.form_submit_button("Guardar Salida")

        if submit_salida:
            codigo_barras_salida = limpiar_texto(codigo_barras_salida).upper()
            nombre_producto_salida = limpiar_texto(nombre_producto_salida)

            if not nombre_producto_salida:
                st.error("⚠️ Debes ingresar el nombre del producto.")
            elif cantidad_salida <= 0:
                st.error("⚠️ La cantidad debe ser mayor que cero.")
            else:
                producto_encontrado = buscar_producto_por_codigo(codigo_barras_salida, df_inv, df_catalogo)
                if producto_encontrado is not None:
                    nombre_producto_salida = limpiar_texto(producto_encontrado.get("Producto", nombre_producto_salida))
                    categoria_salida = producto_encontrado.get("Categoría", categoria_salida)
                    unidad_salida = producto_encontrado.get("Unidad_Medida", unidad_salida)

                producto_existente = df_inv[
                    df_inv["Producto"].astype(str).str.upper().str.strip() == nombre_producto_salida.upper()
                ]
                if producto_existente.empty and codigo_barras_salida:
                    producto_existente = df_inv[
                        df_inv["Código_Barra"].astype(str).str.upper().str.strip() == codigo_barras_salida
                    ]

                if producto_existente.empty:
                    st.error("⚠️ El producto no existe en el inventario. Primero registra una entrada.")
                else:
                    idx = producto_existente.index[0]
                    stock_previo = float(df_inv.at[idx, "Stock_Actual"] if not pd.isna(df_inv.at[idx, "Stock_Actual"]) else 0)
                    if stock_previo < float(cantidad_salida):
                        st.error(f"⚠️ Stock insuficiente. Disponible: {stock_previo:g} {unidad_salida}.")
                    else:
                        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        df_inv.at[idx, "Stock_Actual"] = stock_previo - float(cantidad_salida)
                        df_inv.at[idx, "Ultima_Actualizacion"] = fecha_actual
                        df_inv.at[idx, "Código_Barra"] = codigo_barras_salida or df_inv.at[idx, "Código_Barra"]

                        nuevo_mov = pd.DataFrame([{
                            "Fecha": fecha_actual,
                            "Tipo_Movimiento": "Salida",
                            "Código_Barra": codigo_barras_salida,
                            "Producto": nombre_producto_salida,
                            "Cantidad": cantidad_salida,
                            "Unidad_Medida": unidad_salida,
                            "Categoría": categoria_salida,
                            "Factura_Guia": "",
                            "Bodega": "",
                            "Proveedor": "",
                            "Campo": campo_salida,
                            "Variedad": variedad_salida,
                            "Cuartel": cuartel_salida,
                            "Responsable": responsable_salida,
                            "Observaciones": observaciones_salida,
                        }])
                        df_mov = pd.concat([df_mov, nuevo_mov], ignore_index=True)
                        guardar_datos(df_inv, df_catalogo, df_mov)
                        st.success(f"✅ Salida registrada correctamente: {nombre_producto_salida} ({cantidad_salida:g} {unidad_salida}).")
                        st.rerun()

# --- PESTAÑA 3: INVENTARIO ACTUAL ---
with tab3:
    st.markdown("### Inventario Actual")
    if df_inv.empty:
        st.info("El inventario está vacío. Registra una entrada para comenzar.")
    else:
        df_mostrar = df_inv.copy()
        df_mostrar = df_mostrar.rename(columns={
            "Código": "Código",
            "Código_Barra": "Código de Barra",
            "Producto": "Producto",
            "Categoría": "Categoría",
            "Stock_Actual": "Stock Actual",
            "Unidad_Medida": "Unidad",
            "Bodega": "Bodega",
            "Ultima_Actualizacion": "Última Actualización",
        })
        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

# --- PESTAÑA 4: HISTORIAL DE MOVIMIENTOS ---
with tab4:
    st.markdown("### Historial de Movimientos")
    if df_mov.empty:
        st.info("No hay movimientos registrados.")
    else:
        df_historial = df_mov.sort_values(by="Fecha", ascending=False)
        st.dataframe(df_historial, use_container_width=True, hide_index=True)


# Mostrar detalle final del estado de la sesión de usuario
st.caption("📌 Datos guardados en: " + EXCEL_FILE)
