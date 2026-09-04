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
    st.warning("⚠️ No se encontró el archivo 'logo.png'. Por favor, colócalo en la misma carpeta que este programa para visualizarlo.")

st.title("📦 Sistema de Inventario Web")
st.subheader("Servicios Agrícolas Cumbre Ltda.")

# ==========================================
# CONFIGURACIÓN DE RUTA HACIA EL ESCRITORIO
# ==========================================
# Detectar el directorio del usuario (Ej: C:\Users\TuNombre)
home_dir = os.path.expanduser("~")

# Manejar la diferencia entre Windows en Inglés (Desktop) y Español (Escritorio)
ruta_escritorio = os.path.join(home_dir, "Desktop")
if not os.path.exists(ruta_escritorio):
    ruta_escritorio = os.path.join(home_dir, "Escritorio")

# Definir la carpeta y el archivo exacto solicitado
CARPETA_BBDD = os.path.join(ruta_escritorio, "inventario cumbre")
EXCEL_FILE = os.path.join(CARPETA_BBDD, "Inventario_Servicios_Cumbre.xlsx")

def inicializar_entorno():
    """Crea la carpeta en el escritorio y el archivo Excel si no existen."""
    # 1. Crear carpeta si no existe
    if not os.path.exists(CARPETA_BBDD):
        os.makedirs(CARPETA_BBDD)
        
    # 2. Crear Excel con las hojas necesarias si no existe
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
    """Carga los datos asegurando que el entorno esté creado."""
    inicializar_entorno()
    df_inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventario_Actual")
    df_mov = pd.read_excel(EXCEL_FILE, sheet_name="Movimientos")
    return df_inv, df_mov

def guardar_datos(df_inv, df_mov):
    """Guarda los datos en la ruta del Escritorio."""
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="Inventario_Actual", index=False)
        df_mov.to_excel(writer, sheet_name="Movimientos", index=False)

# Cargar datos al iniciar la app
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
            
            # UNIDADES DE MEDIDA CORRECTAS
            unidades_limpias = ["Litros", "Kilos", "Unidades", "Sacos", "Cajas", "Bidones", "Gramos", "Metros"]
            unidad = st.selectbox("Unidad de Medida", unidades_limpias)
            
        with col2:
            cantidad = st.number_input("Cantidad", min_value=0.01, step=1.0)
            factura_guia = st.text_input("N° Factura o Guía (Opcional)")
            observaciones = st.text_area("Observaciones (Opcional)")
            
        submit_btn = st.form_submit_button("Guardar Movimiento")
        
        if submit_btn:
            if not producto.strip():
                st.error("El nombre del producto no puede estar vacío.")
            else:
                fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                tipo_limpio = "Entrada" if "Entrada" in tipo_mov else "Salida"
                
                # 1. Guardar en Historial
                nuevo_mov = pd.DataFrame([{
                    "Fecha": fecha_actual,
                    "Tipo_Movimiento": tipo_limpio,
                    "Producto": producto.strip().upper(),
                    "Cantidad": cantidad,
                    "Unidad_Medida": unidad,
                    "Factura_Guia": factura_guia,
                    "Observaciones": observaciones
                }])
                df_mov = pd.concat([df_mov, nuevo_mov], ignore_index=True)
                
                # 2. Actualizar Inventario
                prod_upper = producto.strip().upper()
                if prod_upper in df_inv["Producto"].values:
                    idx = df_inv.index[df_inv["Producto"] == prod_upper].tolist()[0]
                    stock_previo = float(df_inv.at[idx, "Stock_Actual"])
                    
                    if tipo_limpio == "Entrada":
                        df_inv.at[idx, "Stock_Actual"] = stock_previo + cantidad
                    else:
                        df_inv.at[idx, "Stock_Actual"] = stock_previo - cantidad
                        
                    df_inv.at[idx, "Ultima_Actualizacion"] = fecha_actual
                else:
                    stock_inicial = cantidad if tipo_limpio == "Entrada" else -cantidad
                    nuevo_inv = pd.DataFrame([{
                        "Código": f"PROD-{len(df_inv)+1:04d}",
                        "Producto": prod_upper,
                        "Categoría": categoria,
                        "Stock_Actual": stock_inicial,
                        "Unidad_Medida": unidad,
                        "Ultima_Actualizacion": fecha_actual
                    }])
                    df_inv = pd.concat([df_inv, nuevo_inv], ignore_index=True)
                
                # Guardar cambios
                guardar_datos(df_inv, df_mov)
                st.success(f"✅ Movimiento guardado exitosamente. Se actualizó el stock de {prod_upper}.")
                st.rerun()

# --- PESTAÑA 2: INVENTARIO ACTUAL ---
with tab2:
    st.markdown("### Stock en Bodega")
    st.caption(f"📂 Conectado a: {EXCEL_FILE}")
    if df_inv.empty:
        st.info("El inventario está vacío. Registra un movimiento para comenzar.")
    else:
        df_inv_mostrar = df_inv.copy()
        st.dataframe(df_inv_mostrar, use_container_width=True, hide_index=True)

# --- PESTAÑA 3: HISTORIAL DE MOVIMIENTOS ---
with tab3:
    st.markdown("### Registro Histórico")
    if df_mov.empty:
        st.info("No hay movimientos registrados.")
    else:
        df_mov_mostrar = df_mov.sort_values(by="Fecha", ascending=False)
        st.dataframe(df_mov_mostrar, use_container_width=True, hide_index=True)