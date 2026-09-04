import os
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import pandas as pd

EXCEL_FILE = "inventario_cumbre.xlsx"


class AppInventarioPC:

  def __init__(self, root):
    self.root = root
    self.root.title(
        "Software de Escritorio — Servicios Agrícolas Cumbre Ltda"
    )
    self.root.geometry("1150x650")
    self.pedir_password()

  def pedir_password(self):
    pwd = simpledialog.askstring(
        "Seguridad",
        "Ingrese contraseña de administrador:",
        show="*",
        parent=self.root,
    )
    if pwd == "cumbre2026":
      self.cargar_interfaz()
    else:
      messagebox.showerror("Error", "Contraseña incorrecta.")
      self.root.destroy()

  def cargar_datos(self):
    if os.path.exists(EXCEL_FILE):
      try:
        self.df_inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventario_Actual")
        self.df_mov = pd.read_excel(EXCEL_FILE, sheet_name="Movimientos")
        self.df_inv["Factura_Guia"] = (
            self.df_inv["Factura_Guia"].fillna("").astype(str)
        )
      except Exception:
        self.df_inv = pd.DataFrame()
        self.df_mov = pd.DataFrame()
    else:
      self.df_inv = pd.DataFrame()
      self.df_mov = pd.DataFrame()

  def cargar_interfaz(self):
    self.cargar_datos()

    self.notebook = ttk.Notebook(self.root)
    self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

    # Pestaña 1: Inventario Actual
    self.frame_inv = ttk.Frame(self.notebook)
    self.notebook.add(self.frame_inv, text="📦 Stock Actual (Inventario)")
    self.setup_inventario_tab()

    # Pestaña 2: Movimientos
    self.frame_mov = ttk.Frame(self.notebook)
    self.notebook.add(self.frame_mov, text="📊 Historial de Movimientos")
    self.setup_movimientos_tab()

  def setup_inventario_tab(self):
    frame_top = ttk.Frame(self.frame_inv)
    frame_top.pack(fill="x", padx=5, pady=5)

    btn_sync = ttk.Button(
        frame_top,
        text="🔄 Sincronizar / Recargar Datos",
        command=self.actualizar_tablas,
    )
    btn_sync.pack(side="left", padx=5)

    self.tree_inv = ttk.Treeview(self.frame_inv, selectmode="extended")
    self.tree_inv.pack(fill="both", expand=True, padx=5, pady=5)
    self.poblar_tree_inventario()

  def poblar_tree_inventario(self):
    for item in self.tree_inv.get_children():
      self.tree_inv.delete(item)

    if not self.df_inv.empty:
      self.tree_inv["columns"] = list(self.df_inv.columns)
      self.tree_inv["show"] = "headings"

      for col in self.df_inv.columns:
        self.tree_inv.heading(col, text=col)
        self.tree_inv.column(col, width=130, anchor="w")

      for _, row in self.df_inv.iterrows():
        self.tree_inv.insert("", "end", values=list(row))

  def setup_movimientos_tab(self):
    self.tree_mov = ttk.Treeview(self.frame_mov, selectmode="extended")
    self.tree_mov.pack(fill="both", expand=True, padx=5, pady=5)
    self.poblar_tree_movimientos()

  def poblar_tree_movimientos(self):
    for item in self.tree_mov.get_children():
      self.tree_mov.delete(item)

    if not self.df_mov.empty:
      self.tree_mov["columns"] = list(self.df_mov.columns)
      self.tree_mov["show"] = "headings"

      for col in self.df_mov.columns:
        self.tree_mov.heading(col, text=col)
        self.tree_mov.column(col, width=130, anchor="w")

      for _, row in self.df_mov.iterrows():
        self.tree_mov.insert("", "end", values=list(row))

  def actualizar_tablas(self):
    self.cargar_datos()
    self.poblar_tree_inventario()
    self.poblar_tree_movimientos()
    messagebox.showinfo(
        "Sincronización", "Datos actualizados correctamente desde el Excel."
    )


if __name__ == "__main__":
  root = tk.Tk()
  app = AppInventarioPC(root)
  root.mainloop()