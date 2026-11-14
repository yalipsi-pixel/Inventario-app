import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA Y LOGO
# ==========================================
st.set_page_config(
    page_title="Inventario - Servicios Agrícolas Cumbre",
    page_icon="📦",
    layout="wide"
)

# RESTAURACIÓN DEL LOGO
ruta_logo = "LOGO CUMBRE_2.jpg"
if os.path.exists(ruta_logo):
    st.image(ruta_logo, width=250)
else:
    st.warning(f"⚠️ No se encontró el archivo '{ruta_logo}'. Por favor, colócalo en la misma carpeta que este programa para visualizarlo.")

st.title("📦 Sistema de Inventario Web")
st.subheader("Servicios Agrícolas Cumbre Ltda.")

# ==========================================
# CONFIGURACIÓN DE RUTA HACIA EL ESCRITORIO
# ==========================================
home_dir = os.path.expanduser("~")
ruta_escritorio = os.path.join(home_dir, "Desktop")
if not os.path.exists(ruta_escritorio):
    ruta_escritorio = os.path.join(home_dir, "Escritorio")

CARPETA_BBDD = os.path.join(ruta_escritorio, "inventario cumbre")
EXCEL_FILE = os.path.join(CARPETA_BBDD, "Inventario_Servicios_Cumbre.xlsx")


def inicializar_entorno():
    if not os.path.exists(CARPETA_BBDD):
        os.makedirs(CARPETA_BBDD)

    if not os.path.exists(EXCEL_FILE):
        df_inv = pd.DataFrame(columns=[
            "Código", "Producto", "Categoría", "Stock_Actual", "Unidad_Medida", "Ultima_Actualizacion"
        ])
        df_mov = pd.DataFrame(columns=[
            "Fecha", "Tipo_Movimiento", "Producto", "Cantidad", "Unidad_Medida", "Factura_Guia", "Observaciones"
        ])
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_inv.to_excel(writer, sheet_name="Inventario_Actual", index=False)
            df_mov.to_excel(writer, sheet_name="Movimientos", index=False)


def cargar_datos():
    inicializar_entorno()
    df_inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventario_Actual")
    df_mov = pd.read_excel(EXCEL_FILE, sheet_name="Movimientos")
    return df_inv, df_mov


def guardar_datos(df_inv, df_mov):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="Inventario_Actual", index=False)
        df_mov.to_excel(writer, sheet_name="Movimientos", index=False)


def preparar_inventario_para_mostrar(df_inv, df_mov):
    """Agrega al inventario las cantidades acumuladas de entradas y salidas."""
    df_mostrar = df_inv.copy()
    columnas_cantidades = ["Total_Entradas", "Total_Salidas"]

    if df_mostrar.empty:
        for columna in columnas_cantidades:
            df_mostrar[columna] = pd.Series(dtype="float64")
        return df_mostrar

    if df_mov.empty:
        df_mostrar["Total_Entradas"] = 0
        df_mostrar["Total_Salidas"] = 0
        return df_mostrar

    movimientos = df_mov.copy()
    movimientos["Cantidad"] = pd.to_numeric(movimientos["Cantidad"], errors="coerce").fillna(0)
    movimientos["Producto"] = movimientos["Producto"].astype(str).str.strip().str.upper()
    movimientos["Tipo_Movimiento"] = movimientos["Tipo_Movimiento"].astype(str).str.strip().str.capitalize()

    resumen = movimientos.pivot_table(
        index="Producto",
        columns="Tipo_Movimiento",
        values="Cantidad",
        aggfunc="sum",
        fill_value=0
    )
    resumen = resumen.rename(columns={"Entrada": "Total_Entradas", "Salida": "Total_Salidas"})
    for columna in columnas_cantidades:
        if columna not in resumen.columns:
            resumen[columna] = 0

    df_mostrar["Producto"] = df_mostrar["Producto"].astype(str).str.strip().str.upper()
    df_mostrar = df_mostrar.merge(
        resumen[columnas_cantidades],
        left_on="Producto",
        right_index=True,
        how="left"
    )
    df_mostrar[columnas_cantidades] = df_mostrar[columnas_cantidades].fillna(0)
    return df_mostrar


df_inv, df_mov = cargar_datos()

# ==========================================
# MENÚ DE NAVEGACIÓN (PESTAÑAS)
# ==========================================
tab1, tab2, tab3 = st.tabs(["➕ Registrar Movimiento", "📦 Inventario Actual", "📊 Historial de Movimientos"])

# --- PESTAÑA 1: REGISTRAR MOVIMIENTO ---
with tab1:
    st.markdown("### Ingresar nueva Entrada o Salida")

    with st.form("form_movimiento"):
        col1, col2 = st.columns(2)

        with col1:
            tipo_mov = st.radio("Tipo de Movimiento:", ["Entrada 📥", "Salida 📤"])
            producto = st.text_input("Nombre del Producto")
            categoria = st.selectbox("Categoría", ["Agroquímicos", "Fertilizantes", "Herramientas", "Semillas", "EPP (Seguridad)", "Otros"])

            unidades_limpias = ["-- Seleccionar --", "Kg", "gr", "cc", "Litros", "Unidades"]
            unidad = st.selectbox("Unidad de Medida", unidades_limpias)

        with col2:
            cantidad = st.number_input("Cantidad", min_value=0.01, step=1.0)
            factura_guia = st.text_input("N° Factura o Guía (Opcional)")
            observaciones = st.text_area("Observaciones (Opcional)")

        submit_btn = st.form_submit_button("Guardar Movimiento")

        if submit_btn:
            if unidad == "-- Seleccionar --":
                st.error("⚠️ Por favor, selecciona una Unidad de Medida válida.")
            elif not producto.strip():
                st.error("⚠️ El nombre del producto no puede estar vacío.")
            else:
                fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                tipo_limpio = "Entrada" if "Entrada" in tipo_mov else "Salida"
                prod_upper = producto.strip().upper()

                producto_existente = df_inv[df_inv["Producto"].astype(str).str.strip().str.upper() == prod_upper]
                if not producto_existente.empty:
                    idx = producto_existente.index[0]
                    unidad_existente = str(df_inv.at[idx, "Unidad_Medida"])
                    stock_previo = float(pd.to_numeric(df_inv.at[idx, "Stock_Actual"], errors="coerce") or 0)

                    if unidad_existente != unidad:
                        st.error(f"⚠️ Este producto está registrado en {unidad_existente}. Usa la misma unidad de medida.")
                    elif tipo_limpio == "Salida" and cantidad > stock_previo:
                        st.error(f"⚠️ Stock insuficiente. Disponible: {stock_previo:g} {unidad}.")
                    else:
                        nuevo_mov = pd.DataFrame([{
                            "Fecha": fecha_actual,
                            "Tipo_Movimiento": tipo_limpio,
                            "Producto": prod_upper,
                            "Cantidad": cantidad,
                            "Unidad_Medida": unidad,
                            "Factura_Guia": factura_guia,
                            "Observaciones": observaciones
                        }])
                        df_mov = pd.concat([df_mov, nuevo_mov], ignore_index=True)
                        cambio_stock = cantidad if tipo_limpio == "Entrada" else -cantidad
                        df_inv.at[idx, "Stock_Actual"] = stock_previo + cambio_stock
                        df_inv.at[idx, "Ultima_Actualizacion"] = fecha_actual
                        guardar_datos(df_inv, df_mov)
                        st.success(f"✅ {tipo_limpio} de {cantidad:g} {unidad} registrada. Stock actual de {prod_upper}: {stock_previo + cambio_stock:g} {unidad}.")
                        st.rerun()
                else:
                    if tipo_limpio == "Salida":
                        st.error("⚠️ No se puede registrar una salida: el producto aún no existe en el inventario.")
                    else:
                        nuevo_mov = pd.DataFrame([{
                            "Fecha": fecha_actual,
                            "Tipo_Movimiento": tipo_limpio,
                            "Producto": prod_upper,
                            "Cantidad": cantidad,
                            "Unidad_Medida": unidad,
                            "Factura_Guia": factura_guia,
                            "Observaciones": observaciones
                        }])
                        nuevo_inv = pd.DataFrame([{
                            "Código": f"PROD-{len(df_inv)+1:04d}",
                            "Producto": prod_upper,
                            "Categoría": categoria,
                            "Stock_Actual": cantidad,
                            "Unidad_Medida": unidad,
                            "Ultima_Actualizacion": fecha_actual
                        }])
                        df_mov = pd.concat([df_mov, nuevo_mov], ignore_index=True)
                        df_inv = pd.concat([df_inv, nuevo_inv], ignore_index=True)
                        guardar_datos(df_inv, df_mov)
                        st.success(f"✅ Entrada de {cantidad:g} {unidad} registrada. Stock actual de {prod_upper}: {cantidad:g} {unidad}.")
                        st.rerun()

# --- PESTAÑA 2: INVENTARIO ACTUAL ---
with tab2:
    st.markdown("### Stock en Bodega")
    st.caption(f"📂 Conectado a: {EXCEL_FILE}")
    if df_inv.empty:
        st.info("El inventario está vacío. Registra un movimiento para comenzar.")
    else:
        df_inv_mostrar = preparar_inventario_para_mostrar(df_inv, df_mov)
        df_inv_mostrar = df_inv_mostrar.rename(columns={
            "Stock_Actual": "Stock Actual",
            "Unidad_Medida": "Unidad",
            "Total_Entradas": "Total Ingresado",
            "Total_Salidas": "Total Salido",
            "Ultima_Actualizacion": "Última Actualización"
        })
        columnas = [
            "Código", "Producto", "Categoría", "Total Ingresado", "Total Salido",
            "Stock Actual", "Unidad", "Última Actualización"
        ]
        st.dataframe(df_inv_mostrar[columnas], use_container_width=True, hide_index=True)

# --- PESTAÑA 3: HISTORIAL DE MOVIMIENTOS ---
with tab3:
    st.markdown("### Registro Histórico")
    if df_mov.empty:
        st.info("No hay movimientos registrados.")
    else:
        df_mov_mostrar = df_mov.sort_values(by="Fecha", ascending=False)
        st.dataframe(df_mov_mostrar, use_container_width=True, hide_index=True)
