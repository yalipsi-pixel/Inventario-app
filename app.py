import os
import shutil
import sys
from datetime import datetime, time
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

try:
    import win32com.client as win32
except ImportError:
    win32 = None

st.set_page_config(page_title="Gestión de Inventario y Trazabilidad", page_icon="📦", layout="wide")

# -------- Apariencia --------
st.markdown("""
<style>
:root { color-scheme: dark; }
[data-testid="stAppViewContainer"], [data-testid="stHeader"] { background:#101010; color:#f5f5f5; }
[data-testid="stHeader"] { background:#101010; }
[data-testid="stSidebar"] { background:#181818; }
h1,h2,h3,p,label,[data-testid="stCaptionContainer"] { color:#f2f2f2 !important; }
[data-testid="stForm"],[data-testid="stDataFrame"] { background:#181818; border:1px solid #333; border-radius:10px; padding:10px; }
[data-baseweb="input"]>div,[data-baseweb="textarea"]>div,[data-baseweb="select"]>div,[data-testid="stNumberInput"] div { background:#222 !important; color:#f5f5f5 !important; border-color:#444 !important; }
input,textarea { color:#f5f5f5 !important; }
[data-testid="stTabs"] [role="tablist"] { gap:8px; border-bottom:1px solid #444; }
[data-testid="stTabs"] button[role="tab"] { color:#ddd; background:#1b1b1b; border-radius:8px 8px 0 0; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color:#fff; background:#333; }
[data-testid="stFormSubmitButton"] button,[data-testid="stButton"] button,[data-testid="stDownloadButton"] button { background:#303030; color:#fff; border:1px solid #555; border-radius:8px; }
[data-testid="stFormSubmitButton"] button:hover,[data-testid="stButton"] button:hover,[data-testid="stDownloadButton"] button:hover { background:#444; border-color:#888; color:#fff; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = Path.home() / "OneDrive - Personal" / "CYO gestion traz e inventario"
EXCEL_FILE = Path(os.environ.get("INVENTARIO_EXCEL", str(BASE_DIR / "Inventario_Servicios_Cumbre (version 1).xlsx"))).expanduser()
LOGO_FILE = Path(os.environ.get("INVENTARIO_LOGO", str(BASE_DIR / "Gestion de inventario y trazabilidad.png"))).expanduser()

SHEET_IN = "Historial ingreso"
SHEET_STOCK = "Stock"
SHEET_OUT = "Movimientos-Salida"
IN_HEADERS = ["Codigo Barras", "Fecha y hora", "Factura / Guia", "Nombre Producto", "Tipo", "Cantidad", "Unidad Medida", "Bodega", "Proveedor"]
STOCK_HEADERS = ["Codigo Barras", "Factura / Guia", "Nombre Producto", "Tipo", "Stock Actual", "Unidad Medida", "Bodega", "Proveedor"]
OUT_HEADERS = ["Codigo Barras", "Fecha", "Nombre del Producto", "Cantidad", "Unidad de Medida", "Campo", "Variedad", "Cuartel", "Responsable"]
REQUIRED = {SHEET_IN: IN_HEADERS, SHEET_STOCK: STOCK_HEADERS, SHEET_OUT: OUT_HEADERS}

if LOGO_FILE.is_file():
    st.image(str(LOGO_FILE), width=700)
st.title("Gestión de Inventario y Trazabilidad")
st.caption("Una herramienta sencilla para registrar movimientos y mantener tus existencias bajo control.")


def norm(value):
    return " ".join(str(value or "").strip().casefold().split())


def headers_for(ws):
    result = {}
    for i, cell in enumerate(ws[1], start=1):
        if cell.value is not None:
            result[norm(cell.value)] = i
    return result


def validate_book(wb):
    missing = [s for s in REQUIRED if s not in wb.sheetnames]
    if missing:
        raise ValueError("Faltan pestañas: " + ", ".join(missing))
    for sheet, expected in REQUIRED.items():
        found = headers_for(wb[sheet])
        absent = [name for name in expected if norm(name) not in found]
        if absent:
            raise ValueError(f"En '{sheet}' faltan columnas: {', '.join(absent)}")


def read_book():
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"No se encontró el Excel: {EXCEL_FILE}")
    # Solo lectura con openpyxl; las escrituras las realiza Excel para preservar tablas y gráficos dinámicos.
    wb = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    try:
        validate_book(wb)
        result = {sheet: pd.DataFrame(list(wb[sheet].values)[1:], columns=list(wb[sheet].values)[0]) for sheet in REQUIRED}
    finally:
        wb.close()
    return result


def check_write_ready():
    if sys.platform != "win32" or win32 is None:
        raise RuntimeError("Para escribir en el Excel preservando Tab-Graf, ejecuta esta app en Windows con Microsoft Excel instalado y pywin32 disponible.")
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"No se encontró el Excel: {EXCEL_FILE}")


def excel_write(action):
    """Edita el libro mediante Excel COM para no eliminar los gráficos/tablas dinámicas."""
    check_write_ready()
    backup = EXCEL_FILE.with_name(f"{EXCEL_FILE.stem}.respaldo_{datetime.now():%Y%m%d_%H%M%S}{EXCEL_FILE.suffix}")
    shutil.copy2(EXCEL_FILE, backup)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        book = excel.Workbooks.Open(str(EXCEL_FILE.resolve()), UpdateLinks=0, ReadOnly=False)
        if book.ReadOnly:
            raise RuntimeError("Excel abrió el libro como solo lectura. Ciérralo en Excel y revisa que OneDrive haya terminado de sincronizar.")
        result = action(book)
        book.Save()
        return result, backup
    except Exception:
        if book is not None:
            book.Close(SaveChanges=False)
            book = None
        # Si una acción falla a mitad de camino, deja el original como estaba.
        shutil.copy2(backup, EXCEL_FILE)
        raise
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()


def com_headers(ws):
    last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  # xlToLeft
    return {norm(ws.Cells(1, col).Value): col for col in range(1, last_col + 1) if ws.Cells(1, col).Value is not None}


def com_last_data_row(ws, key_col):
    return max(1, ws.Cells(ws.Rows.Count, key_col).End(-4162).Row)  # xlUp


def append_com(ws, headers, columns, record):
    row = com_last_data_row(ws, headers[norm(columns[0])]) + 1
    # Cuando hay tabla de Excel, insertar mediante ListRows mantiene su rango y formato.
    if ws.ListObjects.Count:
        table = ws.ListObjects(1)
        try:
            added = table.ListRows.Add()
            row = added.Range.Row
        except Exception:
            pass
    for name in columns:
        ws.Cells(row, headers[norm(name)]).Value = record.get(name)
    return row


def row_for_product(ws, headers, code, name):
    code_key, name_key = norm(code), norm(name)
    key_col = headers[norm("Codigo Barras")]
    last = com_last_data_row(ws, key_col)
    for row in range(2, last + 1):
        row_code = norm(ws.Cells(row, key_col).Value)
        row_name = norm(ws.Cells(row, headers[norm("Nombre Producto")]).Value)
        if code_key and row_code == code_key:
            return row
        if not code_key and name_key and row_name == name_key:
            return row
    return None


def save_record(book, sheet_name, headers_expected, record):
    ws = book.Worksheets(sheet_name)
    headers = com_headers(ws)
    return append_com(ws, headers, headers_expected, record)


def read_records(df, columns):
    if df.empty:
        return pd.DataFrame(columns=columns)
    # Normaliza nombres para que encabezados de Excel no dependan de mayúsculas/minúsculas.
    rename = {col: next((expected for expected in columns if norm(col) == norm(expected)), col) for col in df.columns}
    df = df.rename(columns=rename)
    return df.reindex(columns=columns)


def filter_df(df, query):
    if not query.strip() or df.empty:
        return df
    mask = df.astype(str).apply(lambda col: col.str.contains(query.strip(), case=False, na=False, regex=False)).any(axis=1)
    return df.loc[mask]


def download_csv(label, df, filename, key):
    st.download_button(label, data=df.to_csv(index=False).encode("utf-8-sig"), file_name=filename, mime="text/csv", key=key)


def show_error(error):
    st.error(str(error))
    st.info("El Excel original no debería quedar parcialmente modificado. Asegúrate de que esté cerrado en Excel y que OneDrive esté sincronizado antes de volver a guardar.")


if not EXCEL_FILE.is_file():
    st.error(f"No se encontró el archivo: `{EXCEL_FILE}`")
    st.info("Confirma la carpeta sincronizada o configura INVENTARIO_EXCEL con la ruta completa del archivo.")
    st.stop()
try:
    data = read_book()
except Exception as error:
    show_error(f"No se pudo leer o validar el archivo: {error}")
    st.stop()

st.caption(f"📂 Archivo conectado: {EXCEL_FILE}")
st.caption("Tab-Graf se conserva. Los cambios al libro se hacen mediante Excel en Windows para preservar sus gráficos y tablas dinámicas.")

# Métricas superiores
stock_df = read_records(data[SHEET_STOCK], STOCK_HEADERS)
stock_df["Stock Actual"] = pd.to_numeric(stock_df["Stock Actual"], errors="coerce").fillna(0)
threshold = st.sidebar.number_input("Umbral de alerta de bajo stock", min_value=0.0, value=10.0, step=1.0, help="Se usa solo para destacar productos; no cambia el Excel.")
low_stock = stock_df[(stock_df["Stock Actual"] > 0) & (stock_df["Stock Actual"] <= threshold)]
empty_stock = stock_df[stock_df["Stock Actual"] <= 0]

m1, m2, m3, m4 = st.columns(4)
m1.metric("Productos", len(stock_df))
m2.metric("Unidades en stock", f"{stock_df['Stock Actual'].sum():,.2f}")
m3.metric("Stock bajo", len(low_stock))
m4.metric("Sin existencias", len(empty_stock))

if len(low_stock):
    with st.expander(f"⚠️ Revisar {len(low_stock)} producto(s) con stock bajo"):
        st.dataframe(low_stock, use_container_width=True, hide_index=True)

# Una copia descargable exacta del Excel, sin editar el original.
try:
    st.download_button("⬇️ Descargar copia del Excel", data=EXCEL_FILE.read_bytes(), file_name=f"{EXCEL_FILE.stem}_copia_{datetime.now():%Y%m%d_%H%M}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", help="Descarga una copia local. No modifica el archivo de OneDrive.")
except OSError as error:
    st.warning(f"No se pudo preparar la descarga: {error}")

if sys.platform != "win32" or win32 is None:
    st.warning("Modo de solo lectura: para registrar o editar y conservar los gráficos de Tab-Graf, usa Windows con Excel instalado y ejecuta `pip install pywin32`.")

tab_entry, tab_exit, tab_inventory, tab_history = st.tabs(["➕ Registrar entrada", "➖ Registrar salida", "📦 Stock y edición", "🔎 Historial y reportes"])

with tab_entry:
    st.subheader("Registrar ingreso")
    with st.form("form_entry", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            code = st.text_input("Código de barras")
            invoice = st.text_input("Factura / Guía")
            name = st.text_input("Nombre del producto *")
        with c2:
            kind = st.text_input("Tipo", placeholder="Fertilizante, herbicida...")
            qty = st.number_input("Cantidad ingresada *", min_value=0.01, step=1.0)
            unit = st.text_input("Unidad de medida *", placeholder="Kg, Lt, gr...")
        with c3:
            warehouse = st.text_input("Bodega")
            supplier = st.text_input("Proveedor")
        submit_in = st.form_submit_button("Guardar entrada", type="primary")
    if submit_in:
        if not name.strip() or not unit.strip():
            st.error("Nombre del producto y unidad son obligatorios.")
        else:
            try:
                def action(book):
                    ws_stock = book.Worksheets(SHEET_STOCK)
                    hs = com_headers(ws_stock)
                    row = row_for_product(ws_stock, hs, code, name)
                    if row is not None:
                        old_unit = ws_stock.Cells(row, hs[norm("Unidad Medida")]).Value
                        if norm(old_unit) != norm(unit):
                            raise ValueError(f"El producto ya está registrado en {old_unit}. No se pueden mezclar unidades.")
                        current = pd.to_numeric(ws_stock.Cells(row, hs[norm("Stock Actual")]).Value, errors="coerce")
                        current = 0.0 if pd.isna(current) else float(current)
                        new_stock = current + float(qty)
                        ws_stock.Cells(row, hs[norm("Stock Actual")]).Value = new_stock
                        for col_name, val in [("Factura / Guia", invoice), ("Tipo", kind), ("Bodega", warehouse), ("Proveedor", supplier)]:
                            if val.strip():
                                ws_stock.Cells(row, hs[norm(col_name)]).Value = val.strip()
                    else:
                        new_stock = float(qty)
                        save_record(book, SHEET_STOCK, STOCK_HEADERS, {"Codigo Barras": code.strip(), "Factura / Guia": invoice.strip(), "Nombre Producto": name.strip(), "Tipo": kind.strip(), "Stock Actual": new_stock, "Unidad Medida": unit.strip(), "Bodega": warehouse.strip(), "Proveedor": supplier.strip()})
                    save_record(book, SHEET_IN, IN_HEADERS, {"Codigo Barras": code.strip(), "Fecha y hora": datetime.now(), "Factura / Guia": invoice.strip(), "Nombre Producto": name.strip(), "Tipo": kind.strip(), "Cantidad": float(qty), "Unidad Medida": unit.strip(), "Bodega": warehouse.strip(), "Proveedor": supplier.strip()})
                    return new_stock
                result, backup = excel_write(action)
                st.success(f"Entrada registrada. Stock de {name}: {result:g} {unit}. Respaldo: {backup.name}")
                st.rerun()
            except Exception as error:
                show_error(error)

with tab_exit:
    st.subheader("Registrar salida")
    products = stock_df[stock_df["Nombre Producto"].notna()].copy()
    if products.empty:
        st.info("No hay productos registrados en Stock.")
    else:
        labels = [f"{r['Nombre Producto']} | {r['Codigo Barras'] or 'sin código'} | stock {r['Stock Actual']:g} {r['Unidad Medida']}" for _, r in products.iterrows()]
        with st.form("form_exit", clear_on_submit=True):
            selected_label = st.selectbox("Producto *", labels)
            selected = products.iloc[labels.index(selected_label)].to_dict()
            e1, e2 = st.columns(2)
            with e1:
                out_qty = st.number_input("Cantidad que sale *", min_value=0.01, step=1.0)
                field = st.text_input("Campo")
                variety = st.text_input("Variedad")
            with e2:
                block = st.text_input("Cuartel")
                responsible = st.text_input("Responsable")
            submit_out = st.form_submit_button("Guardar salida", type="primary")
        if submit_out:
            try:
                def action(book):
                    ws_stock = book.Worksheets(SHEET_STOCK)
                    hs = com_headers(ws_stock)
                    row = row_for_product(ws_stock, hs, selected.get("Codigo Barras"), selected.get("Nombre Producto"))
                    if row is None:
                        raise ValueError("Producto no encontrado. Recarga la página y vuelve a intentarlo.")
                    current = pd.to_numeric(ws_stock.Cells(row, hs[norm("Stock Actual")]).Value, errors="coerce")
                    current = 0.0 if pd.isna(current) else float(current)
                    if out_qty > current:
                        raise ValueError(f"Stock insuficiente; disponible: {current:g} {selected.get('Unidad Medida')}.")
                    left = current - float(out_qty)
                    ws_stock.Cells(row, hs[norm("Stock Actual")]).Value = left
                    save_record(book, SHEET_OUT, OUT_HEADERS, {"Codigo Barras": selected.get("Codigo Barras"), "Fecha": datetime.now(), "Nombre del Producto": selected.get("Nombre Producto"), "Cantidad": float(out_qty), "Unidad de Medida": selected.get("Unidad Medida"), "Campo": field.strip(), "Variedad": variety.strip(), "Cuartel": block.strip(), "Responsable": responsible.strip()})
                    return left
                result, backup = excel_write(action)
                st.success(f"Salida guardada. Stock restante: {result:g} {selected.get('Unidad Medida')}. Respaldo: {backup.name}")
                st.rerun()
            except Exception as error:
                show_error(error)

with tab_inventory:
    st.subheader("Stock actual")
    q = st.text_input("🔎 Buscar producto", placeholder="Nombre, código, tipo, bodega o proveedor", key="q_stock")
    visible_stock = filter_df(stock_df, q)
    st.caption(f"{len(visible_stock)} de {len(stock_df)} productos")
    st.dataframe(visible_stock, use_container_width=True, hide_index=True)
    download_csv("⬇️ Descargar resultados del stock (CSV)", visible_stock, "stock_filtrado.csv", "csv_stock")

    st.divider()
    st.subheader("Editar datos del producto")
    st.caption("La edición actualiza los datos descriptivos en Stock. Para ajustar cantidad, registra una entrada o salida: así queda la trazabilidad.")
    if len(products := stock_df[stock_df["Nombre Producto"].notna()]):
        edit_labels = [f"{r['Nombre Producto']} | {r['Codigo Barras'] or 'sin código'}" for _, r in products.iterrows()]
        edit_label = st.selectbox("Selecciona un producto", edit_labels, key="edit_product_choice")
        selected_edit = products.iloc[edit_labels.index(edit_label)].to_dict()
        with st.form("edit_product_form"):
            a, b = st.columns(2)
            with a:
                new_code = st.text_input("Código de barras", value=str(selected_edit.get("Codigo Barras") or ""))
                new_invoice = st.text_input("Factura / Guía", value=str(selected_edit.get("Factura / Guia") or ""))
                new_name = st.text_input("Nombre", value=str(selected_edit.get("Nombre Producto") or ""))
                new_kind = st.text_input("Tipo", value=str(selected_edit.get("Tipo") or ""))
            with b:
                new_unit = st.text_input("Unidad", value=str(selected_edit.get("Unidad Medida") or ""))
                new_warehouse = st.text_input("Bodega", value=str(selected_edit.get("Bodega") or ""))
                new_supplier = st.text_input("Proveedor", value=str(selected_edit.get("Proveedor") or ""))
            submit_edit = st.form_submit_button("Guardar cambios del producto", type="primary")
        if submit_edit:
            if not new_name.strip() or not new_unit.strip():
                st.error("Nombre y unidad son obligatorios.")
            else:
                try:
                    def action(book):
                        ws = book.Worksheets(SHEET_STOCK)
                        hs = com_headers(ws)
                        row = row_for_product(ws, hs, selected_edit.get("Codigo Barras"), selected_edit.get("Nombre Producto"))
                        if row is None:
                            raise ValueError("Producto no encontrado; recarga la página e inténtalo otra vez.")
                        values = {"Codigo Barras": new_code.strip(), "Factura / Guia": new_invoice.strip(), "Nombre Producto": new_name.strip(), "Tipo": new_kind.strip(), "Unidad Medida": new_unit.strip(), "Bodega": new_warehouse.strip(), "Proveedor": new_supplier.strip()}
                        for col_name, value in values.items():
                            ws.Cells(row, hs[norm(col_name)]).Value = value
                    _, backup = excel_write(action)
                    st.success(f"Producto actualizado. Respaldo: {backup.name}")
                    st.rerun()
                except Exception as error:
                    show_error(error)
    else:
        st.info("No hay productos para editar.")

with tab_history:
    st.subheader("Historial de ingresos y salidas")
    h1, h2 = st.columns(2)
    with h1:
        start_date = st.date_input("Desde", value=None, key="history_start")
    with h2:
        end_date = st.date_input("Hasta", value=None, key="history_end")

    incoming = read_records(data[SHEET_IN], IN_HEADERS)
    outgoing = read_records(data[SHEET_OUT], OUT_HEADERS)
    q_in = st.text_input("🔎 Buscar ingresos", placeholder="Producto, código, factura, bodega o proveedor", key="q_in")
    incoming_filtered = filter_df(incoming, q_in)
    if "Fecha y hora" in incoming_filtered:
        dates = pd.to_datetime(incoming_filtered["Fecha y hora"], errors="coerce")
        if start_date:
            incoming_filtered = incoming_filtered[dates.dt.date >= start_date]
            dates = pd.to_datetime(incoming_filtered["Fecha y hora"], errors="coerce")
        if end_date:
            incoming_filtered = incoming_filtered[dates.dt.date <= end_date]
    st.markdown("**Ingresos**")
    st.dataframe(incoming_filtered, use_container_width=True, hide_index=True)
    download_csv("⬇️ Descargar ingresos filtrados (CSV)", incoming_filtered, "ingresos_filtrados.csv", "csv_in")

    q_out = st.text_input("🔎 Buscar salidas", placeholder="Producto, código, campo, variedad o responsable", key="q_out")
    outgoing_filtered = filter_df(outgoing, q_out)
    if "Fecha" in outgoing_filtered:
        dates = pd.to_datetime(outgoing_filtered["Fecha"], errors="coerce")
        if start_date:
            outgoing_filtered = outgoing_filtered[dates.dt.date >= start_date]
            dates = pd.to_datetime(outgoing_filtered["Fecha"], errors="coerce")
        if end_date:
            outgoing_filtered = outgoing_filtered[dates.dt.date <= end_date]
    st.markdown("**Salidas**")
    st.dataframe(outgoing_filtered, use_container_width=True, hide_index=True)
    download_csv("⬇️ Descargar salidas filtradas (CSV)", outgoing_filtered, "salidas_filtradas.csv", "csv_out")
