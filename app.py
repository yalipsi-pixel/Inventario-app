from datetime import datetime
import os
import pandas as pd
import streamlit as st

EXCEL_FILE = 'inventario_cumbre.xlsx'
LOGO_FILE = 'LOGO CUMBRE_2.jpg'


def cargar_datos():
  if not os.path.exists(EXCEL_FILE):
    data = {
        'SKU': ['URE-001', 'AZU-002', 'GLI-003', 'PON-004'],
        'Producto': [
            'Urea Granulada 46%',
            'Azufre Wettable',
            'Glifosato 48 SL',
            'Poncho 600 FS',
        ],
        'Categoría': ['Fertilizantes', 'Fungicidas', 'Herbicidas', 'Insecticidas'],
        'Stock': [150, 80, 45, 30],
        'Unidad': ['Kg', 'Kg', 'Lt', 'Lt'],
        'Bodega': ['El Huingan', 'Bucalemu', 'Rinconada', 'San Felipe'],
        'Proveedor': ['Copeval', 'Gmt', 'M&Valdivieso', 'Copeval'],
    }
    df = pd.DataFrame(data)
    df.to_excel(EXCEL_FILE, index=False)
  else:
    df = pd.read_excel(EXCEL_FILE)
  return df


def guardar_datos(df):
  df.to_excel(EXCEL_FILE, index=False)


st.set_page_config(
    page_title='Control de Inventario - Cumbre Ltda',
    page_icon='🌱',
    layout='wide',
)

if os.path.exists(LOGO_FILE):
  st.image(LOGO_FILE, width=160)

st.title('🌱 Servicios Agrícolas Cumbre Ltda — Panel de Terreno')
st.markdown(
    'Sistema sincronizado en tiempo real con el software de escritorio a través'
    ' de `inventario_cumbre.xlsx`.'
)
st.markdown('---')

df_inventario = cargar_datos()

# --- VENTANA CON 3 BOTONES PRINCIPALES HORIZONTALES ---
col_b1, col_b2, col_b3 = st.columns(3)

# Controlamos la vista activa mediante el session_state de Streamlit
if 'vista_activa' not in st.session_state:
  st.session_state.vista_activa = 'inventario'

with col_b1:
  if st.button('📦 Ver Inventario General', use_container_width=True):
    st.session_state.vista_activa = 'inventario'
with col_b2:
  if st.button('📥 Registrar Entrada', use_container_width=True):
    st.session_state.vista_activa = 'entrada'
with col_b3:
  if st.button('📤 Registrar Salida / Uso', use_container_width=True):
    st.session_state.vista_activa = 'salida'

st.markdown('---')

# --- CONTENIDO SEGÚN EL BOTÓN SELECCIONADO ---

if st.session_state.vista_activa == 'inventario':
  st.subheader('📊 Stock Actual en Bodegas')
  busqueda = st.text_input('🔍 Buscar producto o SKU:')
  if busqueda:
    df_filtrado = df_inventario[
        df_inventario['Producto'].str.contains(busqueda, case=False, na=False)
        | df_inventario['SKU'].str.contains(busqueda, case=False, na=False)
    ]
  else:
    df_filtrado = df_inventario

  st.dataframe(df_filtrado, use_container_width=True)

elif st.session_state.vista_activa == 'entrada':
  st.subheader('📥 Registrar Entrada de Insumos')
  with st.form('form_entrada'):
    producto_sel = st.selectbox(
        'Seleccionar Producto Existente', df_inventario['Producto'].tolist()
    )
    cantidad_ing = st.number_input('Cantidad a Agregar', min_value=1.0, step=1.0)
    submit_entrada = st.form_submit_button('Guardar Entrada')

    if submit_entrada:
      idx = df_inventario[df_inventario['Producto'] == producto_sel].index[0]
      df_inventario.loc[idx, 'Stock'] += cantidad_ing
      guardar_datos(df_inventario)
      st.success(
          f'¡Entrada registrada con éxito! Stock actualizado para'
          f' {producto_sel}.'
      )

elif st.session_state.vista_activa == 'salida':
  st.subheader('📤 Registrar Salida o Aplicación en Terreno')
  with st.form('form_salida'):
    producto_sel = st.selectbox(
        'Seleccionar Producto', df_inventario['Producto'].tolist()
    )
    cantidad_sal = st.number_input(
        'Cantidad Retirada / Usada', min_value=1.0, step=1.0
    )
    cuartel = st.text_input('Cuartel / Campo de Destino')
    responsable = st.text_input('Responsable')
    submit_salida = st.form_submit_button('Registrar Salida')

    if submit_salida:
      idx = df_inventario[df_inventario['Producto'] == producto_sel].index[0]
      stock_actual = df_inventario.loc[idx, 'Stock']
      if cantidad_sal > stock_actual:
        st.error(
            '¡Error! La cantidad a retirar supera el stock disponible en'
            ' bodega.'
        )
      else:
        df_inventario.loc[idx, 'Stock'] -= cantidad_sal
        guardar_datos(df_inventario)
        st.success(
            f'¡Salida registrada correctamente! Nuevo stock para {producto_sel}:'
            f' {df_inventario.loc[idx, "Stock"]}'
        )