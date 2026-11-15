import os
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="Gestión de Inventario y Trazabilidad",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Gestión de Inventario y Trazabilidad")
st.caption("Registro de ingresos, salidas y existencias del inventario.")

# Puedes definir INVENTARIO_EXCEL como variable de entorno si guardas el archivo
# en otra ubicación. Por defecto se busca en la carpeta de OneDrive indicada.
RUTA_PREDETERMINADA = (
    Path.home()
    / "OneDrive - Personal"
    / "CYO gestion traz e inventario"
    / "Inventario_Servicios_Cumbre (version 1).xlsx"
)
EXCEL_FILE = Path(os.environ.get("INVENTARIO_EXCEL", str(RUTA_PREDETERMINADA))).expanduser()

HOJA_INGRESOS = "Historial ingreso"
HOJA_STOCK = "Stock"
HOJA_SALIDAS = "Movimientos-Salida"

ENCABEZADOS_INGRESO = [
    "Codigo Barras", "Fecha y hora", "Factura / Guia", "Nombre Producto",
    "Tipo", "Cantidad", "Unidad Medida", "Bodega", "Proveedor",
]
ENCABEZADOS_STOCK = [
    "Codigo Barras", "Factura / Guia", "Nombre Producto", "Tipo",
    "Stock Actual", "Unidad Medida", "Bodega", "Proveedor",
]
ENCABEZADOS_SALIDA = [
    "Codigo Barras", "Fecha", "Nombre del Producto", "Cantidad",
    "Unidad de Medida", "Campo", "Variedad", "Cuartel", "Responsable",
]


def normalizar(texto):
    return " ".join(str(texto or "").strip().casefold().split())


def obtener_encabezados(hoja):
    return {normalizar(celda.value): indice for indice, celda in enumerate(hoja[1], start=1) if celda.value is not None}


def validar_hojas(libro):
    necesarias = [HOJA_INGRESOS, HOJA_STOCK, HOJA_SALIDAS]
    faltantes = [nombre for nombre in necesarias if nombre not in libro.sheetnames]
    if faltantes:
        raise ValueError(
            "No se encontraron estas pestañas en el Excel: " + ", ".join(faltantes)
            + ". Revisa que sus nombres coincidan exactamente."
        )

    requeridas = {
        HOJA_INGRESOS: ENCABEZADOS_INGRESO,
        HOJA_STOCK: ENCABEZADOS_STOCK,
        HOJA_SALIDAS: ENCABEZADOS_SALIDA,
    }
    for nombre, esperadas in requeridas.items():
        encabezados = obtener_encabezados(libro[nombre])
        faltantes = [col for col in esperadas if normalizar(col) not in encabezados]
        if faltantes:
            raise ValueError(
                f"En la pestaña '{nombre}' faltan columnas: {', '.join(faltantes)}. "
                "No se modificó el archivo."
            )


def abrir_libro():
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(
            f"No se encontró el Excel en: {EXCEL_FILE}. "
            "Confirma la ruta o define la variable de entorno INVENTARIO_EXCEL."
        )
    libro = load_workbook(EXCEL_FILE)
    validar_hojas(libro)
    return libro


def valor_fila(hoja, encabezados, fila, nombre_columna):
    columna = encabezados.get(normalizar(nombre_columna))
    return hoja.cell(fila, columna).value if columna else None


def filas_con_datos(hoja, columna_clave=1):
    return [
        fila for fila in range(2, hoja.max_row + 1)
        if hoja.cell(fila, columna_clave).value not in (None, "")
    ]


def anexar_registro(hoja, encabezados, columnas, registro):
    filas = filas_con_datos(hoja)
    fila_nueva = (max(filas) if filas else 1) + 1
    for nombre in columnas:
        hoja.cell(fila_nueva, encabezados[normalizar(nombre)], registro.get(nombre))
    # Amplía el rango de las tablas de Excel para incluir la fila añadida.
    for tabla in hoja.tables.values():
        min_col, min_fila, max_col, max_fila = __import__("openpyxl").utils.range_boundaries(tabla.ref)
        if min_fila <= 1 <= max_fila and fila_nueva > max_fila:
            tabla.ref = f"{__import__('openpyxl').utils.get_column_letter(min_col)}{min_fila}:{__import__('openpyxl').utils.get_column_letter(max_col)}{fila_nueva}"
    return fila_nueva


def guardar_libro(libro):
    # Conserva una copia anterior del libro antes de guardar cambios.
    respaldo = EXCEL_FILE.with_name(f"{EXCEL_FILE.stem}.respaldo{EXCEL_FILE.suffix}")
    if not respaldo.exists():
        shutil.copy2(EXCEL_FILE, respaldo)
    libro.save(EXCEL_FILE)


def leer_registros(hoja, columnas):
    encabezados = obtener_encabezados(hoja)
    registros = []
    for fila in filas_con_datos(hoja):
        registros.append({
            nombre: valor_fila(hoja, encabezados, fila, nombre)
            for nombre in columnas
        })
    return registros


def cargar_stock():
    libro = abrir_libro()
    hoja = libro[HOJA_STOCK]
    registros = leer_registros(hoja, ENCABEZADOS_STOCK)
    return libro, hoja, registros


def encontrar_producto(hoja, encabezados, codigo, nombre):
    codigo_busqueda = normalizar(codigo)
    nombre_busqueda = normalizar(nombre)
    for fila in filas_con_datos(hoja):
        codigo_fila = normalizar(valor_fila(hoja, encabezados, fila, "Codigo Barras"))
        nombre_fila = normalizar(valor_fila(hoja, encabezados, fila, "Nombre Producto"))
        if codigo_busqueda and codigo_fila == codigo_busqueda:
            return fila
        if not codigo_busqueda and nombre_busqueda and nombre_fila == nombre_busqueda:
            return fila
    return None


def mostrar_error_excel(error):
    st.error(f"No se pudo abrir o validar el archivo Excel: {error}")
    st.info("El archivo no se crea ni se reemplaza automáticamente. Verifica la ruta y los encabezados antes de continuar.")


if not EXCEL_FILE.is_file():
    st.error(f"No se encontró el archivo Excel: `{EXCEL_FILE}`")
    st.info("Revisa que OneDrive esté sincronizado. Si el archivo está en otra ubicación, define la variable de entorno `INVENTARIO_EXCEL` con la ruta completa.")
    st.stop()

try:
    libro_actual = abrir_libro()
except Exception as error:
    mostrar_error_excel(error)
    st.stop()

st.caption(f"📂 Archivo conectado: {EXCEL_FILE}")
st.caption("La pestaña 'Tab-Graf' se conserva y no se modifica desde la app.")

tab_entrada, tab_salida, tab_stock, tab_historial = st.tabs(
    ["➕ Entrada", "➖ Salida", "📦 Stock", "🧾 Historiales"]
)

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
                hoja_ingreso = libro[HOJA_INGRESOS]
                hoja_stock = libro[HOJA_STOCK]
                encabezados_ingreso = obtener_encabezados(hoja_ingreso)
                encabezados_stock = obtener_encabezados(hoja_stock)
                fecha = datetime.now()
                registro_ingreso = {
                    "Codigo Barras": codigo.strip(), "Fecha y hora": fecha,
                    "Factura / Guia": factura.strip(), "Nombre Producto": nombre.strip(),
                    "Tipo": tipo.strip(), "Cantidad": float(cantidad),
                    "Unidad Medida": unidad.strip(), "Bodega": bodega.strip(),
                    "Proveedor": proveedor.strip(),
                }
                anexar_registro(hoja_ingreso, encabezados_ingreso, ENCABEZADOS_INGRESO, registro_ingreso)

                fila_stock = encontrar_producto(hoja_stock, encabezados_stock, codigo, nombre)
                if fila_stock is None:
                    registro_stock = {
                        "Codigo Barras": codigo.strip(), "Factura / Guia": factura.strip(),
                        "Nombre Producto": nombre.strip(), "Tipo": tipo.strip(),
                        "Stock Actual": float(cantidad), "Unidad Medida": unidad.strip(),
                        "Bodega": bodega.strip(), "Proveedor": proveedor.strip(),
                    }
                    anexar_registro(hoja_stock, encabezados_stock, ENCABEZADOS_STOCK, registro_stock)
                    stock_nuevo = float(cantidad)
                else:
                    unidad_actual = valor_fila(hoja_stock, encabezados_stock, fila_stock, "Unidad Medida")
                    if normalizar(unidad_actual) != normalizar(unidad):
                        st.error(f"La unidad registrada para este producto es '{unidad_actual}'. No se guardó el movimiento para evitar mezclar unidades.")
                        st.stop()
                    stock_actual = pd.to_numeric(valor_fila(hoja_stock, encabezados_stock, fila_stock, "Stock Actual"), errors="coerce")
                    stock_actual = 0.0 if pd.isna(stock_actual) else float(stock_actual)
                    stock_nuevo = stock_actual + float(cantidad)
                    hoja_stock.cell(fila_stock, encabezados_stock[normalizar("Stock Actual")], stock_nuevo)
                    for campo, valor in [("Factura / Guia", factura.strip()), ("Tipo", tipo.strip()), ("Bodega", bodega.strip()), ("Proveedor", proveedor.strip())]:
                        if valor:
                            hoja_stock.cell(fila_stock, encabezados_stock[normalizar(campo)], valor)
                guardar_libro(libro)
                st.success(f"Entrada registrada: {cantidad:g} {unidad} de {nombre}. Stock actualizado: {stock_nuevo:g} {unidad}.")
                st.rerun()
            except Exception as error:
                mostrar_error_excel(error)

with tab_salida:
    st.subheader("Registrar salida de productos")
    try:
        _, _, productos_stock = cargar_stock()
    except Exception as error:
        mostrar_error_excel(error)
        productos_stock = []

    disponibles = [p for p in productos_stock if p.get("Nombre Producto")]
    if not disponibles:
        st.info("No hay productos registrados en la pestaña Stock.")
    else:
        opciones = {
            f"{p.get('Nombre Producto')} | Código: {p.get('Codigo Barras') or 'sin código'} | Stock: {p.get('Stock Actual') or 0} {p.get('Unidad Medida') or ''}": p
            for p in disponibles
        }
        with st.form("form_salida", clear_on_submit=True):
            seleccion = st.selectbox("Producto*", list(opciones.keys()))
            producto_seleccionado = opciones[seleccion]
            col1, col2 = st.columns(2)
            with col1:
                cantidad_salida = st.number_input("Cantidad que sale*", min_value=0.01, step=1.0)
                campo = st.text_input("Campo")
                variedad = st.text_input("Variedad")
            with col2:
                cuartel = st.text_input("Cuartel")
                responsable = st.text_input("Responsable")
            guardar_salida = st.form_submit_button("Guardar salida", type="primary")

        if guardar_salida:
            try:
                libro = abrir_libro()
                hoja_stock = libro[HOJA_STOCK]
                hoja_salida = libro[HOJA_SALIDAS]
                encabezados_stock = obtener_encabezados(hoja_stock)
                encabezados_salida = obtener_encabezados(hoja_salida)
                codigo_producto = producto_seleccionado.get("Codigo Barras")
                nombre_producto = producto_seleccionado.get("Nombre Producto")
                fila_stock = encontrar_producto(hoja_stock, encabezados_stock, codigo_producto, nombre_producto)
                if fila_stock is None:
                    st.error("No se encontró el producto en Stock. Recarga la página y vuelve a intentarlo.")
                else:
                    stock_actual = pd.to_numeric(valor_fila(hoja_stock, encabezados_stock, fila_stock, "Stock Actual"), errors="coerce")
                    stock_actual = 0.0 if pd.isna(stock_actual) else float(stock_actual)
                    if float(cantidad_salida) > stock_actual:
                        st.error(f"Stock insuficiente. Disponible: {stock_actual:g} {producto_seleccionado.get('Unidad Medida') or ''}.")
                    else:
                        registro_salida = {
                            "Codigo Barras": codigo_producto, "Fecha": datetime.now(),
                            "Nombre del Producto": nombre_producto, "Cantidad": float(cantidad_salida),
                            "Unidad de Medida": producto_seleccionado.get("Unidad Medida"),
                            "Campo": campo.strip(), "Variedad": variedad.strip(),
                            "Cuartel": cuartel.strip(), "Responsable": responsable.strip(),
                        }
                        anexar_registro(hoja_salida, encabezados_salida, ENCABEZADOS_SALIDA, registro_salida)
                        stock_nuevo = stock_actual - float(cantidad_salida)
                        hoja_stock.cell(fila_stock, encabezados_stock[normalizar("Stock Actual")], stock_nuevo)
                        guardar_libro(libro)
                        st.success(f"Salida registrada: {cantidad_salida:g} {producto_seleccionado.get('Unidad Medida') or ''} de {nombre_producto}. Stock restante: {stock_nuevo:g}.")
                        st.rerun()
            except Exception as error:
                mostrar_error_excel(error)

with tab_stock:
    st.subheader("Existencias actuales (pestaña Stock)")
    try:
        _, _, registros_stock = cargar_stock()
        if registros_stock:
            df_stock = pd.DataFrame(registros_stock).rename(columns={
                "Codigo Barras": "Código de barras",
                "Factura / Guia": "Factura / Guía",
                "Nombre Producto": "Nombre del producto",
                "Stock Actual": "Stock actual",
                "Unidad Medida": "Unidad de medida",
            })
            st.dataframe(df_stock, use_container_width=True, hide_index=True)
        else:
            st.info("La pestaña Stock todavía no contiene productos.")
    except Exception as error:
        mostrar_error_excel(error)

with tab_historial:
    st.subheader("Historial de ingresos")
    try:
        libro = abrir_libro()
        ingresos = leer_registros(libro[HOJA_INGRESOS], ENCABEZADOS_INGRESO)
        if ingresos:
            st.dataframe(pd.DataFrame(ingresos), use_container_width=True, hide_index=True)
        else:
            st.info("Todavía no hay ingresos registrados.")

        st.subheader("Movimientos de salida")
        salidas = leer_registros(libro[HOJA_SALIDAS], ENCABEZADOS_SALIDA)
        if salidas:
            st.dataframe(pd.DataFrame(salidas), use_container_width=True, hide_index=True)
        else:
            st.info("Todavía no hay salidas registradas.")
    except Exception as error:
        mostrar_error_excel(error)
