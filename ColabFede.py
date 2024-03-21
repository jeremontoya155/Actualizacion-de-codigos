import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pandastable import Table, TableModel

def leer_archivo(archivo):
    # Leer el archivo Excel y devolver un DataFrame
    return pd.read_excel(archivo)

def encontrar_codigos_faltantes(archivo1, archivo2):
    # Leer los archivos Excel
    df1 = leer_archivo(archivo1)
    df2 = leer_archivo(archivo2)

    # Convertir todas las celdas a cadenas para que puedan ser comparadas
    df1 = df1.applymap(str)
    df2 = df2.applymap(str)

    # Combinar todas las columnas en una sola serie y obtener los códigos únicos
    codigos_archivo1 = set(df1.stack().unique())
    codigos_archivo2 = set(df2.stack().unique())

    # Encontrar los códigos únicos en cada archivo
    codigos_unicos_archivo1 = codigos_archivo1 - codigos_archivo2
    codigos_unicos_archivo2 = codigos_archivo2 - codigos_archivo1

    # Encontrar los códigos faltantes y sus IDs de producto asociados
    resultados = []

    for idx, row in df1.iterrows():
        id_producto = row['idproducto']
        codigos_faltantes = [codigo if codigo in codigos_unicos_archivo1 else '0' for col, codigo in row.items() if col != 'idproducto']
        if any(codigos_faltantes):
            resultados.append([id_producto, *codigos_faltantes, archivo1])  # Agregar el nombre del archivo

    for idx, row in df2.iterrows():
        id_producto = row['idproducto']
        codigos_faltantes = [codigo if codigo in codigos_unicos_archivo2 else '0' for col, codigo in row.items() if col != 'idproducto']
        if any(codigos_faltantes):
            resultados.append([id_producto, *codigos_faltantes, archivo2])  # Agregar el nombre del archivo

    return resultados

def cargar_archivo(entrada):
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    entrada.delete(0, tk.END)
    entrada.insert(0, archivo)

def mostrar_resultados():
    archivo1 = entry_archivo1.get()
    archivo2 = entry_archivo2.get()
    if archivo1 == '' or archivo2 == '':
        messagebox.showwarning("Advertencia", "Por favor, selecciona ambos archivos.")
        return

    resultados = encontrar_codigos_faltantes(archivo1, archivo2)

    # Mostrar los resultados en una tabla
    df_resultados = pd.DataFrame(resultados)
    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True)
    pt = Table(frame, dataframe=df_resultados, showtoolbar=True, showstatusbar=True)
    pt.show()

    messagebox.showinfo("Información", "Los resultados se han mostrado en la tabla.")

def descargar_resultados():
    archivo_salida = 'resultados.xlsx'
    resultados = encontrar_codigos_faltantes(entry_archivo1.get(), entry_archivo2.get())
    df_resultados = pd.DataFrame(resultados)
    df_resultados.to_excel(archivo_salida, index=False)
    messagebox.showinfo("Información", f"Los resultados se han descargado como '{archivo_salida}'.")

# Configurar la interfaz gráfica
root = tk.Tk()
root.title("Comparador de Archivos Excel")

# Entrada para el primer archivo
frame_archivo1 = tk.Frame(root)
frame_archivo1.pack(pady=10)
label_archivo1 = tk.Label(frame_archivo1, text="Archivo 1:")
label_archivo1.pack(side="left", padx=(10,5))
entry_archivo1 = tk.Entry(frame_archivo1, width=50)
entry_archivo1.pack(side="left", padx=5)
button_archivo1 = tk.Button(frame_archivo1, text="Seleccionar archivo", command=lambda: cargar_archivo(entry_archivo1))
button_archivo1.pack(side="left", padx=(5,10))

# Entrada para el segundo archivo
frame_archivo2 = tk.Frame(root)
frame_archivo2.pack(pady=10)
label_archivo2 = tk.Label(frame_archivo2, text="Archivo 2:")
label_archivo2.pack(side="left", padx=(10,5))
entry_archivo2 = tk.Entry(frame_archivo2, width=50)
entry_archivo2.pack(side="left", padx=5)
button_archivo2 = tk.Button(frame_archivo2, text="Seleccionar archivo", command=lambda: cargar_archivo(entry_archivo2))
button_archivo2.pack(side="left", padx=(5,10))

# Botón para mostrar resultados
button_mostrar_resultados = tk.Button(root, text="Mostrar resultados", command=mostrar_resultados)
button_mostrar_resultados.pack(pady=10)

# Botón para descargar resultados
button_descargar = tk.Button(root, text="Descargar resultados como Excel", command=descargar_resultados)
button_descargar.pack(pady=10)

root.mainloop()
