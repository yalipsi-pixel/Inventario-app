import os
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

st.set_page_config(page_title="Gestión de Inventario y Trazabilidad", page_icon="📦", layout="wide")

RUTA_LOGO_PREDETERMINADA = Path.home() / "OneDrive - Personal" / "CYO gestion traz e inventario" / "Gestion de inventario y trazabilidad.png"
LOGO_FILE = Path(os.environ.get("INVENTARIO_LOGO", str(RUTA_LOGO_PREDETERMINADA))).expanduser()

st.markdown("""
<style>
:root { color-scheme: dark; }
[data-testid="stAppViewContainer"], [data-testid="stHeader"] { background:#101010; color:#f5f5f5; }
[data-testid="stHeader"] { background:#101010; }
[data-testid="stSidebar"] { background:#181818; }
h1,h2,h3,p,label,[data-testid="stCaptionContainer"] { color:#f2f2f2 !important; }
[data-testid="stForm"],[data-testid="stDataFrame"] { background:#181818; border:1px solid #333; border-radius:10px; padding:12px; }
[data-baseweb="input"]>div,[data-baseweb="textarea"]>div,[data-baseweb="select"]>div,[data-testid="stNumberInput"] div { background:#222 !important; color:#f5f5f5 !important; border-color:#444 !important; }
input,textarea { color:#f5f5f5 !important; }
[data-testid="stTabs"] [role="tablist"] { gap:8px; border-bottom:1px solid #444; }
[data-testid="stTabs"] button[role="tab"] { color:#ddd; background:#1b1b1b; border-radius:8px 8px 0 0; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color:#fff; background:#333; }
[data-testid="stFormSubmitButton"] button,[data-testid="stButton"] button { background:#303030; color:#fff; border:1px solid #555; border-radius:8px; }
[data-testid="stFormSubmitButton"] button:hover,[data-testid="stButton"] button:hover { background:#444; border-color:#888; color:#fff; }
</style>
""", unsafe_allow_html=True)

if LOGO_FILE.is_file():
    st.image(str(LOGO_FILE), width=700)
st.title("Gestión de Inventario y Trazabilidad")
st.caption("Registra entradas y salidas, consulta existencias y encuentra productos rápidamente.")

RUTA_PREDETERMINADA = Path.home() / "OneDrive - Personal" / "CYO gestion traz e inventario" / "Inventario_Servicios_Cumbre (version 1).xlsx"
EXCEL_FILE = Path(os.environ.get("INVENTARIO_EXCEL", str(RUTA_PREDETERMINADA))).expanduser()
HOJA_INGRESOS, HOJA_STOCK, HOJA_SALIDAS = "Historial ingreso", "Stock", "Movimientos-Salida"
ENCABEZADOS_INGRESO = ["Codigo Barras", "Fecha y hora", "Factura / Guia", "Nombre Producto", "Tipo", "Cantidad", "Unidad Medida", "Bodega", "Proveedor"]
ENCABEZADOS_STOCK = ["Codigo Barras", "Factura / Guia", "Nombre Producto", "Tipo", "Stock Actual", "Unidad Medida", "Bodega", "Proveedor"]
ENCABEZADOS_SALIDA = ["Codigo Barras", "Fecha", "Nombre del Producto", "Cantidad", "Unidad de Medida", "Campo", "Variedad", "Cuartel", "Responsable"]


def normalizar(texto):
    return " ".join(str(texto or "").strip().casefold().split())


def obtener_encabezados(hoja):
    return {normalizar(celda.value): indice for indice, celda in enumerate(hoja[1], start=1) if celda.value is not None}


def validar_hojas(libro):
    faltantes = [nombre for nombre in [HOJA_INGRESOS, HOJA_STOCK, HOJA_SALIDAS] if nombre not in libro.sheetnames]
    if faltantes:
        raise ValueError("No se encontraron estas pestañas: " + ", ".join(faltantes))
    for nombre, esperadas in [(HOJA_INGRESOS, ENCABEZADOS_INGRESO), (HOJA_STOCK, ENCABEZADOS_STOCK), (HOJA_SALIDAS, ENCABEZADOS_SALIDA)]:
        encabezados = obtener_encabezados(libro[nombre])
        faltan = [col for col in esperadas if normalizar(col) not in encabezados]
        if faltan:
            raise ValueError(f"En '{nombre}' faltan columnas: {', '.join(faltan)}. No se modificó el archivo.")


def abrir_libro():
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"No se encontró el Excel: {EXCEL_FILE}")
    libro = load_workbook(EXCEL_FILE)
    validar_hojas(libro)
    return libro


def valor_fila(hoja, encabezados, fila, columna):
    indice = encabezados.get(normalizar(columna))
    return hoja.cell(fila, indice).value if indice else None


def filas_con_datos(hoja, columna_clave=1):
    return [fila for fila in range(2, hoja.max_row + 1) if hoja.cell(fila, columna_clave).value not in (None, "")]


def anexar_registro(hoja, encabezados, columnas, registro):
    filas = filas_con_datos(hoja)
    fila_nueva = (max(filas) if filas else 1) + 1
    for nombre in columnas:
        hoja.cell(fila_nueva, encabezados[normalizar(nombre)], registro.get(nombre))
    for tabla in hoja.tables.values():
        min_col, min_fila, max_col, max_fila = range_boundaries(tabla.ref)
        if min_fila <= 1 <= max_fila and fila_nueva > max_fila:
            tabla.ref = f"{get_column_letter(min_col)}{min_fila}:{get_column_letter(max_col)}{fila_nueva}"
    return fila_nueva


def guardar_libro(libro):
    # El respaldo se crea una sola vez; nunca se reemplaza por una versión posterior.
    respaldo = EXCEL_FILE.with_name(f"{EXCEL_FILE.stem}.respaldo{EXCEL_FILE.suffix}")
    if not respaldo.exists():
        shutil.copy2(EXCEL_FILE, respaldo)
    libro.save(EXCEL_FILE)


def leer_registros(hoja, columnas):
    encabezados = obtener_encabezados(hoja)
    return [{nombre: valor_fila(hoja, encabezados, fila, nombre) for nombre in columnas} for fila in filas_con_datos(hoja)]


def encontrar_producto(hoja, encabezados, codigo, nombre):
    codigo_busqueda, nombre_busqueda = normalizar(codigo), normalizar(nombre)
    for fila in filas_con_datos(hoja):
        cod = normalizar(valor_fila(hoja, encabezados, fila, "Codigo Barras"))
        nom = normalizar(valor_fila(hoja, encabezados, fila, "Nombre Producto"))
        if codigo_busqueda and cod == codigo_busqueda:
            return fila
        if not codigo_busqueda and nombre_busqueda and nom == nombre_busqueda:
            return fila
    return None


def filtrar_df(df, consulta):
    if not consulta.strip() or df.empty:
        return df
    mascara = df.astype(str).apply(lambda columna: columna.str.contains(consulta.strip(), case=False, na=False, regex=False)).any(axis=1)
    return df[mascara]


def mostrar_error(error):
    st.error(f"No se pudo abrir, validar o guardar el Excel: {error}")
    st.info("Revisa que el archivo esté cerrado en Excel, que OneDrive esté sincronizado y que la ruta y los encabezados sean correctos.")


if not EXCEL_FILE.is_file():
    st.error(f"No se encontró el Excel: `{EXCEL_FILE}`")
    st.info("Revisa que OneDrive esté sincronizado o configura la variable INVENTARIO_EXCEL con la ruta completa del archivo.")
    st.stop()
try:
    abrir_libro()
except Exception as error:
    mostrar_error(error)
    st.stop()

st.caption(f"📂 Archivo conectado: {EXCEL_FILE}")
st.caption("Tab-Graf no se edita desde la app. Cada guardado crea antes una copia de respaldo del libro.")

# Descargar el libro actual como una copia local: este botón no escribe ni altera el original.
try:
    st.download_button(
        "⬇️ Descargar copia del Excel",
        data=EXCEL_FILE.read_bytes(),
        file_name=f"{EXCEL_FILE.stem}_copia_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Descarga una copia del archivo; no modifica el Excel de OneDrive.",
    )
except OSError as error:
    st.warning(f"No se pudo preparar la descarga: {error}")

tab_entrada, tab_salida, tab_stock, tab_historial = st.tabs(["➕ Entrada", "➖ Salida", "📦 Stock", "🧾 Historiales"])

with tab_entrada:
    st.subheader("Registrar ingreso de productos")
    with st.form("form_entrada", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            codigo = st.text_input("Código de barras")
            factura = st.text_input("Factura / Guía")
            nombre = st.text_input("Nombre del producto*")
        with col2:
            tipo = st.text_input("Tipo", placeholder="Ej.: Fertilizante, Herbicida")
            cantidad = st.number_input("Cantidad ingresada*", min_value=0.01, step=1.0)
            unidad = st.text_input("Unidad de medida*", placeholder="Kg, Lt, gr...")
        with col3:
            bodega = st.text_input("Bodega")
            proveedor = st.text_input("Proveedor")
        guardar_entrada = st.form_submit_button("Guardar entrada", type="primary")
    if guardar_entrada:
        if not nombre.strip() or not unidad.strip():
            st.error("Completa el nombre del producto y la unidad de medida.")
        else:
            try:
                libro = abrir_libro()
                hoja_ingreso, hoja_stock = libro[HOJA_INGRESOS], libro[HOJA_STOCK]
                hi, hs = obtener_encabezados(hoja_ingreso), obtener_encabezados(hoja_stock)
                registro = {"Codigo Barras": codigo.strip(), "Fecha y hora": datetime.now(), "Factura / Guia": factura.strip(), "Nombre Producto": nombre.strip(), "Tipo": tipo.strip(), "Cantidad": float(cantidad), "Unidad Medida": unidad.strip(), "Bodega": bodega.strip(), "Proveedor": proveedor.strip()}
                fila = encontrar_producto(hoja_stock, hs, codigo, nombre)
                if fila is not None:
                    unidad_actual = valor_fila(hoja_stock, hs, fila, "Unidad Medida")
                    if normalizar(unidad_actual) != normalizar(unidad):
                        st.error(f"La unidad registrada para este producto es '{unidad_actual}'. No se guardó el movimiento.")
                        st.stop()
                    previo = pd.to_numeric(valor_fila(hoja_stock, hs, fila, "Stock Actual"), errors="coerce")
                    stock_nuevo = (0 if pd.isna(previo) else float(previo)) + float(cantidad)
                    hoja_stock.cell(fila, hs[normalizar("Stock Actual")], stock_nuevo)
                    for campo, valor in [("Factura / Guia", factura.strip()), ("Tipo", tipo.strip()), ("Bodega", bodega.strip()), ("Proveedor", proveedor.strip())]:
                        if valor:
                            hoja_stock.cell(fila, hs[normalizar(campo)], valor)
                else:
                    stock_nuevo = float(cantidad)
                    registro_stock = {"Codigo Barras": codigo.strip(), "Factura / Guia": factura.strip(), "Nombre Producto": nombre.strip(), "Tipo": tipo.strip(), "Stock Actual": stock_nuevo, "Unidad Medida": unidad.strip(), "Bodega": bodega.strip(), "Proveedor": proveedor.strip()}
                    anexar_registro(hoja_stock, hs, ENCABEZADOS_STOCK, registro_stock)
                anexar_registro(hoja_ingreso, hi, ENCABEZADOS_INGRESO, registro)
                guardar_libro(libro)
                st.success(f"Entrada registrada. Stock de {nombre}: {stock_nuevo:g} {unidad}.")
                st.rerun()
            except Exception as error:
                mostrar_error(error)

with tab_salida:
    st.subheader("Registrar salida de productos")
    try:
        libro = abrir_libro()
        productos = leer_registros(libro[HOJA_STOCK], ENCABEZADOS_STOCK)
    except Exception as error:
        mostrar_error(error)
        productos = []
    productos = [p for p in productos if p.get("Nombre Producto")]
    if not productos:
        st.info("No hay productos registrados en Stock.")
    else:
        opciones = {f"{p.get('Nombre Producto')} | Código: {p.get('Codigo Barras') or 'sin código'} | Stock: {p.get('Stock Actual') or 0} {p.get('Unidad Medida') or ''}": p for p in productos}
        with st.form("form_salida", clear_on_submit=True):
            seleccion = st.selectbox("Producto*", list(opciones))
            producto = opciones[seleccion]
            c1, c2 = st.columns(2)
            with c1:
                cantidad_salida = st.number_input("Cantidad que sale*", min_value=0.01, step=1.0)
                campo = st.text_input("Campo")
                variedad = st.text_input("Variedad")
            with c2:
                cuartel = st.text_input("Cuartel")
                responsable = st.text_input("Responsable")
            guardar_salida = st.form_submit_button("Guardar salida", type="primary")
        if guardar_salida:
            try:
                libro = abrir_libro()
                hoja_stock, hoja_salida = libro[HOJA_STOCK], libro[HOJA_SALIDAS]
                hs, ho = obtener_encabezados(hoja_stock), obtener_encabezados(hoja_salida)
                fila = encontrar_producto(hoja_stock, hs, producto.get("Codigo Barras"), producto.get("Nombre Producto"))
                if fila is None:
                    st.error("No se encontró el producto en Stock. Recarga la página e inténtalo otra vez.")
                else:
                    previo = pd.to_numeric(valor_fila(hoja_stock, hs, fila, "Stock Actual"), errors="coerce")
                    stock_actual = 0 if pd.isna(previo) else float(previo)
                    if float(cantidad_salida) > stock_actual:
                        st.error(f"Stock insuficiente. Disponible: {stock_actual:g} {producto.get('Unidad Medida') or ''}.")
                    else:
                        salida = {"Codigo Barras": producto.get("Codigo Barras"), "Fecha": datetime.now(), "Nombre del Producto": producto.get("Nombre Producto"), "Cantidad": float(cantidad_salida), "Unidad de Medida": producto.get("Unidad Medida"), "Campo": campo.strip(), "Variedad": variedad.strip(), "Cuartel": cuartel.strip(), "Responsable": responsable.strip()}
                        anexar_registro(hoja_salida, ho, ENCABEZADOS_SALIDA, salida)
                        stock_nuevo = stock_actual - float(cantidad_salida)
                        hoja_stock.cell(fila, hs[normalizar("Stock Actual")], stock_nuevo)
                        guardar_libro(libro)
                        st.success(f"Salida registrada. Stock restante de {producto.get('Nombre Producto')}: {stock_nuevo:g} {producto.get('Unidad Medida') or ''}.")
                        st.rerun()
            except Exception as error:
                mostrar_error(error)

with tab_stock:
    st.subheader("Existencias actuales")
    try:
        libro = abrir_libro()
        registros = leer_registros(libro[HOJA_STOCK], ENCABEZADOS_STOCK)
        df_stock = pd.DataFrame(registros, columns=ENCABEZADOS_STOCK)
        consulta_stock = st.text_input("🔎 Buscar en stock", placeholder="Producto, código, bodega, tipo o proveedor", key="buscar_stock")
        df_stock_filtrado = filtrar_df(df_stock, consulta_stock)
        st.caption(f"Mostrando {len(df_stock_filtrado)} de {len(df_stock)} productos")
        st.dataframe(df_stock_filtrado.rename(columns={"Codigo Barras":"Código de barras", "Factura / Guia":"Factura / Guía", "Nombre Producto":"Nombre del producto", "Stock Actual":"Stock actual", "Unidad Medida":"Unidad de medida"}), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Editar un producto")
        if registros:
            opciones_edicion = {f"{r.get('Nombre Producto')} | Código: {r.get('Codigo Barras') or 'sin código'}": r for r in registros}
            elegido = st.selectbox("Selecciona el producto que quieres editar", list(opciones_edicion), key="producto_editar")
            producto_editar = opciones_edicion[elegido]
            if st.button("✏️ Editar producto", key="abrir_edicion"):
                st.session_state["mostrar_form_edicion"] = True
            if st.session_state.get("mostrar_form_edicion", False):
                st.warning("Al guardar, estos campos se actualizarán directamente en la pestaña Stock. El cambio manual de stock no crea un movimiento en los historiales.")
                with st.form("form_editar_stock"):
                    e1, e2 = st.columns(2)
                    with e1:
                        edit_codigo = st.text_input("Código de barras", value=str(producto_editar.get("Codigo Barras") or ""))
                        edit_factura = st.text_input("Factura / Guía", value=str(producto_editar.get("Factura / Guia") or ""))
                        edit_nombre = st.text_input("Nombre del producto", value=str(producto_editar.get("Nombre Producto") or ""))
                        edit_tipo = st.text_input("Tipo", value=str(producto_editar.get("Tipo") or ""))
                    with e2:
                        valor_stock = pd.to_numeric(producto_editar.get("Stock Actual"), errors="coerce")
                        edit_stock = st.number_input("Stock actual", min_value=0.0, value=max(0.0, float(valor_stock) if not pd.isna(valor_stock) else 0.0), step=1.0)
                        edit_unidad = st.text_input("Unidad de medida", value=str(producto_editar.get("Unidad Medida") or ""))
                        edit_bodega = st.text_input("Bodega", value=str(producto_editar.get("Bodega") or ""))
                        edit_proveedor = st.text_input("Proveedor", value=str(producto_editar.get("Proveedor") or ""))
                    guardar_edicion = st.form_submit_button("Guardar cambios", type="primary")
                if guardar_edicion:
                    if not edit_nombre.strip() or not edit_unidad.strip():
                        st.error("Nombre y unidad de medida son obligatorios.")
                    else:
                        try:
                            libro = abrir_libro()
                            hoja = libro[HOJA_STOCK]
                            encabezados = obtener_encabezados(hoja)
                            fila = encontrar_producto(hoja, encabezados, producto_editar.get("Codigo Barras"), producto_editar.get("Nombre Producto"))
                            if fila is None:
                                st.error("No se encontró el producto. Actualiza la página e inténtalo de nuevo.")
                            else:
                                cambios = {"Codigo Barras":edit_codigo.strip(), "Factura / Guia":edit_factura.strip(), "Nombre Producto":edit_nombre.strip(), "Tipo":edit_tipo.strip(), "Stock Actual":float(edit_stock), "Unidad Medida":edit_unidad.strip(), "Bodega":edit_bodega.strip(), "Proveedor":edit_proveedor.strip()}
                                for campo, valor in cambios.items():
                                    hoja.cell(fila, encabezados[normalizar(campo)], valor)
                                guardar_libro(libro)
                                st.session_state["mostrar_form_edicion"] = False
                                st.success("Producto actualizado en Stock.")
                                st.rerun()
                        except Exception as error:
                            mostrar_error(error)
        else:
            st.info("No hay productos para editar.")
    except Exception as error:
        mostrar_error(error)

with tab_historial:
    try:
        libro = abrir_libro()
        st.subheader("Historial de ingresos")
        ingresos = leer_registros(libro[HOJA_INGRESOS], ENCABEZADOS_INGRESO)
        df_ingresos = pd.DataFrame(ingresos, columns=ENCABEZADOS_INGRESO)
        buscar_ingresos = st.text_input("🔎 Buscar ingresos", placeholder="Producto, código, factura, bodega o proveedor", key="buscar_ingresos")
        filtrado_ingresos = filtrar_df(df_ingresos, buscar_ingresos)
        st.caption(f"Mostrando {len(filtrado_ingresos)} de {len(df_ingresos)} ingresos")
        st.dataframe(filtrado_ingresos, use_container_width=True, hide_index=True)

        st.subheader("Movimientos de salida")
        salidas = leer_registros(libro[HOJA_SALIDAS], ENCABEZADOS_SALIDA)
        df_salidas = pd.DataFrame(salidas, columns=ENCABEZADOS_SALIDA)
        buscar_salidas = st.text_input("🔎 Buscar salidas", placeholder="Producto, código, campo, variedad o responsable", key="buscar_salidas")
        filtrado_salidas = filtrar_df(df_salidas, buscar_salidas)
        st.caption(f"Mostrando {len(filtrado_salidas)} de {len(df_salidas)} salidas")
        st.dataframe(filtrado_salidas, use_container_width=True, hide_index=True)
    except Exception as error:
        mostrar_error(error)
