"""
------------------------------------------------------------
Nombre de la Aplicación: Copiador de Bloques – Origen por Carpeta y 4 Salidas Fijas
Descripción: 
  Esta aplicación permite extraer bloques de datos específicos de archivos
  Excel (XLSX) que contienen certificados y copiarlos a archivos de salida 
  en formato XLSM (archivos de Excel con macros). Los datos se extraen según 
  bloques predefinidos (10 filas x 3 columnas) y se ubican en la hoja "Registro"
  de cada archivo de salida, basándose en la coincidencia de claves encontradas 
  en la fila 10 de cada archivo de origen.
------------------------------------------------------------
"""

import openpyxl      # Biblioteca para manipular archivos Excel sin macros.
import xlwings as xw # Biblioteca para interactuar con Excel preservando macros y formatos.
import tkinter as tk # Biblioteca para crear interfaces gráficas.
from tkinter import filedialog, messagebox, font  # Módulos específicos de Tkinter.
import os            # Biblioteca para interactuar con el sistema operativo (archivos y directorios).

# --- Funciones de selección de archivos y carpetas ---

def seleccionar_carpeta_origen(entry):
    """
    Permite al usuario seleccionar la carpeta que contiene los archivos de origen (XLSX).
    
    Parámetros:
      entry: widget Entry donde se mostrará la ruta seleccionada.
    """
    folder_path = filedialog.askdirectory(title="Seleccionar Carpeta de Origen (XLSX)")
    if folder_path:
        entry.delete(0, tk.END)
        entry.insert(0, folder_path)

def seleccionar_archivo_salida(entry):
    """
    Permite al usuario seleccionar un archivo de salida en formato XLSM.
    
    Parámetros:
      entry: widget Entry donde se mostrará la ruta del archivo seleccionado.
    """
    filepath = filedialog.askopenfilename(
        title='Seleccionar archivo de Salida (XLSM)',
        filetypes=[('Archivos Excel con macros', '*.xlsm')]
    )
    if filepath:
        entry.delete(0, tk.END)
        entry.insert(0, filepath)

# --- Función principal para procesar los archivos ---
def procesar_archivos():
    """
    Función principal que realiza lo siguiente:
      1. Valida que se haya seleccionado la carpeta de origen y extrae los archivos XLSX.
      2. Valida e interpreta las filas de inicio ingresadas.
      3. Organiza los archivos de origen basándose en un prefijo presente en sus nombres.
      4. Extrae bloques de datos y las claves de cada archivo de origen.
      5. Abre los archivos de salida (XLSM) y pega los datos extraídos en la posición adecuada,
         buscando la coincidencia de la clave en la columna D de la hoja "Registro".
      6. Guarda y cierra cada archivo de salida.
    """
    # 1. Obtener la carpeta de origen y validar que no esté vacía.
    carpeta_origen = entry_origen.get().strip()
    if not carpeta_origen:
        messagebox.showerror("Error", "Debe seleccionar la carpeta de origen.")
        return

    # 2. Listar todos los archivos XLSX en la carpeta de origen.
    archivos_origen = [os.path.join(carpeta_origen, f)
                       for f in os.listdir(carpeta_origen)
                       if f.lower().endswith(".xlsx")]
    if not archivos_origen:
        messagebox.showerror("Error", "No se encontraron archivos XLSX en la carpeta de origen.")
        return

    # 3. Obtener y validar los valores de las filas de inicio ingresadas.
    skiprows_str = entry_skiprows.get().strip()
    try:
        # Convertir la cadena de valores separados por comas a una lista de enteros.
        skiprows_list = list(map(int, skiprows_str.split(',')))
    except ValueError:
        messagebox.showerror("Error", "Los valores de fila de inicio deben ser enteros separados por comas.")
        return

    num_origen = len(skiprows_list)
    
    # Actualizar estado: Inicio del procesamiento.
    status_label.config(text="Procesando archivos...", fg="blue")
    root.update_idletasks()
    
    # 4. Organizar los archivos de origen en un diccionario usando como clave el prefijo (dos primeros caracteres del nombre).
    origen_por_prefijo = {}
    for f in archivos_origen:
        nombre = os.path.basename(f)
        pref = nombre[:2]  # Se asume que el prefijo es formado por los dos primeros caracteres.
        if pref not in origen_por_prefijo:
            origen_por_prefijo[pref] = f

    # 5. Recoger las rutas de los 4 archivos de salida ingresados por el usuario.
    salida_files = [
        entry_salida1.get().strip(),
        entry_salida2.get().strip(),
        entry_salida3.get().strip(),
        entry_salida4.get().strip()
    ]
    if any(s == "" for s in salida_files):
        messagebox.showerror("Error", "Debe seleccionar los 4 archivos de salida.")
        status_label.config(text="Error: faltan archivos de salida.", fg="red")
        return

    # 6. Definir los rangos de los 4 bloques a extraer (cada bloque: 10 filas x 3 columnas).
    bloques_rangos = [
        "G11:I20",  # Bloque 1.
        "G21:I30",  # Bloque 2.
        "G31:I40",  # Bloque 3.
        "G41:I50"   # Bloque 4.
    ]

    # 7. Procesar cada archivo de origen basándose en el prefijo esperado ("01", "02", "03", etc.).
    for i in range(num_origen):
        prefijo_esperado = str(i+1).zfill(2)
        if prefijo_esperado not in origen_por_prefijo:
            messagebox.showwarning("Advertencia", f"No se encontró un archivo de origen con prefijo '{prefijo_esperado}'.")
            continue
        archivo_origen = origen_por_prefijo[prefijo_esperado]
        fila_inicio = skiprows_list[i]

        # 8. Abrir el archivo de origen con openpyxl (se leen los valores y se ignoran las fórmulas).
        try:
            wb_origen = openpyxl.load_workbook(archivo_origen, data_only=True)
            ws_origen = wb_origen["sr ysR(Método) ISO 5725"]
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir o encontrar la hoja en el archivo de origen {os.path.basename(archivo_origen)}:\n{e}")
            continue

        # 9. Leer las claves ubicadas en la fila 10 (celdas G10, H10 e I10).
        clave_G = ws_origen["G10"].value
        clave_H = ws_origen["H10"].value
        clave_I = ws_origen["I10"].value
        claves = [clave_G, clave_H, clave_I]

        # 10. Extraer los 4 bloques de datos usando los rangos definidos.
        bloques_datos = []
        for rango in bloques_rangos:
            celdas = ws_origen[rango]
            datos = []
            for fila in celdas:
                datos.append([cel.value for cel in fila])
            bloques_datos.append(datos)
        wb_origen.close()  # Cerrar el archivo de origen.

        # 11. Para cada bloque, abrir el archivo de salida correspondiente y pegar los datos.
        for bloque_idx, bloque in enumerate(bloques_datos):
            salida_path = salida_files[bloque_idx]
            try:
                app = xw.App(visible=False)
                wb_salida = app.books.open(salida_path)
                ws_salida = wb_salida.sheets["Registro"]
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo de salida {os.path.basename(salida_path)}:\n{e}")
                continue

            try:
                # 12. Por cada una de las 3 columnas del bloque, buscar la clave en la columna W del archivo de salida. SE CAMBIO PARA AJUSTRA A CASO ESPECIAL RTD MEDICIONES
                for c in range(3):
                    clave = claves[c]
                    fila_encontrada = None
                    for fila in range(fila_inicio, fila_inicio + 20):
                        valor_celda = ws_salida.range("W" + str(fila)).value
                        if valor_celda == clave:
                            fila_encontrada = fila
                            break
                    if fila_encontrada is None:
                        messagebox.showwarning("Advertencia", f"No se encontró la clave '{clave}' en {os.path.basename(salida_path)} a partir de la fila {fila_inicio}.")
                        continue

                    # 13. Pegar los 10 valores del bloque en la fila encontrada, a partir de la columna E.
                    fila_destino = fila_encontrada
                    col_inicial = 5  # La columna E corresponde al índice 5.
                    for r in range(10):
                        try:
                            valor = bloque[r][c]
                        except IndexError:
                            valor = None
                        ws_salida.range((fila_destino, col_inicial + r)).value = valor

                wb_salida.save()
                wb_salida.close()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar {os.path.basename(salida_path)}:\n{e}")
            finally:
                app.quit()
    
    # 14. Actualizar estado, abrir la carpeta de origen y notificar finalización.
    status_label.config(text="Proceso completado.", fg="green")
    os.startfile(carpeta_origen)
    messagebox.showinfo("Proceso Finalizado", f"Se han procesado los archivos de origen correctamente.")

# --- Funciones para manejar la base de datos de skiprows usando SQLite ---
import sqlite3

def crear_tabla_skiprows():
    """
    Crea la tabla 'skiprows' en la base de datos si no existe.
    """
    conn = sqlite3.connect('skiprows.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS skiprows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descripcion TEXT NOT NULL,
            valores TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

def guardar_skiprows(descripcion, valores):
    """
    Guarda una nueva entrada de skiprows en la base de datos.
    
    Parámetros:
      descripcion: Descripción de los valores de skiprows.
      valores: Valores de skiprows separados por comas.
    """
    conn = sqlite3.connect('skiprows.db')
    cursor = conn.cursor()
    cursor.execute('INSERT INTO skiprows (descripcion, valores) VALUES (?, ?)', (descripcion, valores))
    conn.commit()
    conn.close()

def obtener_skiprows():
    """
    Obtiene todas las entradas de skiprows de la base de datos.
    
    Retorna:
      Una lista de tuplas con (id, descripcion, valores).
    """
    conn = sqlite3.connect('skiprows.db')
    cursor = conn.cursor()
    cursor.execute('SELECT id, descripcion, valores FROM skiprows')
    rows = cursor.fetchall()
    conn.close()
    return rows

# Crear la tabla 'skiprows' al iniciar la aplicación.
crear_tabla_skiprows()

# --- Interfaz Gráfica con Tkinter ---

root = tk.Tk()
root.title("Copiador de Bloques – Origen por Carpeta y 4 Salidas Fijas")
root.geometry("800x650")
root.configure(bg="#f0f8ff")

# Fuentes personalizadas.
titulo_font = font.Font(family="Helvetica", size=16, weight="bold")
subtitulo_font = font.Font(family="Helvetica", size=12, weight="bold")
instrucciones_font = font.Font(family="Helvetica", size=10)
label_font = font.Font(family="Helvetica", size=10, weight="bold")

# --- Marco superior para el título, propósito e instrucciones ---
frame_top = tk.Frame(root, bg="#f0f8ff", pady=10)
frame_top.pack(fill="x")

lbl_titulo = tk.Label(frame_top, text="Copiador de Bloques – Origen por Carpeta y 4 Salidas Fijas", 
                      bg="#f0f8ff", fg="#003366", font=titulo_font)
lbl_titulo.pack()

lbl_proposito = tk.Label(frame_top, 
                         text="Propósito: Extraer bloques de datos de certificados en XLSX y copiarlos en archivos de salida XLSM manteniendo macros y formatos.",
                         bg="#f0f8ff", fg="#003366", font=subtitulo_font, wraplength=750, justify="center")
lbl_proposito.pack(pady=5)

instrucciones = (
    "Instrucciones de Uso:\n"
    "1. Selecciona la carpeta de origen (XLSX).\n"
    "2. Selecciona los 4 archivos de salida (XLSM) correspondientes:\n"
    "   - Salida 1: Bloque 1 (prefijo '01')\n"
    "   - Salida 2: Bloque 2 (prefijo '04')\n"
    "   - Salida 3: Bloque 3 (prefijo '07')\n"
    "   - Salida 4: Bloque 4 (prefijo '010')\n"
    "3. Ingresa las filas de inicio separadas por comas para cada archivo de origen.\n"
    "4. Haz clic en 'Procesar Archivos' para iniciar la extracción y copia de datos."
)
lbl_instrucciones = tk.Label(frame_top, text=instrucciones, bg="#f0f8ff", fg="#333333", 
                             font=instrucciones_font, justify="left")
lbl_instrucciones.pack(pady=5, padx=10)

# --- Marco principal para los controles de entrada ---
frame_main = tk.Frame(root, bg="#f0f8ff", padx=20, pady=10)
frame_main.pack(fill="both", expand=True)

# Carpeta de origen.
lbl_origen = tk.Label(frame_main, text="Carpeta de Origen (XLSX):", bg="#f0f8ff", font=label_font)
lbl_origen.grid(row=0, column=0, sticky="e", padx=10, pady=5)
entry_origen = tk.Entry(frame_main, width=60)
entry_origen.grid(row=0, column=1, padx=10, pady=5)
btn_origen = tk.Button(frame_main, text="Seleccionar Carpeta", command=lambda: seleccionar_carpeta_origen(entry_origen))
btn_origen.grid(row=0, column=2, padx=10, pady=5)

# Archivos de salida.
lbl_salida1 = tk.Label(frame_main, text="Salida 1 (XLSM, prefijo 01):", bg="#f0f8ff", font=label_font)
lbl_salida1.grid(row=1, column=0, sticky="e", padx=10, pady=5)
entry_salida1 = tk.Entry(frame_main, width=60)
entry_salida1.grid(row=1, column=1, padx=10, pady=5)
btn_salida1 = tk.Button(frame_main, text="Seleccionar", command=lambda: seleccionar_archivo_salida(entry_salida1))
btn_salida1.grid(row=1, column=2, padx=10, pady=5)

lbl_salida2 = tk.Label(frame_main, text="Salida 2 (XLSM, prefijo 04):", bg="#f0f8ff", font=label_font)
lbl_salida2.grid(row=2, column=0, sticky="e", padx=10, pady=5)
entry_salida2 = tk.Entry(frame_main, width=60)
entry_salida2.grid(row=2, column=1, padx=10, pady=5)
btn_salida2 = tk.Button(frame_main, text="Seleccionar", command=lambda: seleccionar_archivo_salida(entry_salida2))
btn_salida2.grid(row=2, column=2, padx=10, pady=5)

lbl_salida3 = tk.Label(frame_main, text="Salida 3 (XLSM, prefijo 07):", bg="#f0f8ff", font=label_font)
lbl_salida3.grid(row=3, column=0, sticky="e", padx=10, pady=5)
entry_salida3 = tk.Entry(frame_main, width=60)
entry_salida3.grid(row=3, column=1, padx=10, pady=5)
btn_salida3 = tk.Button(frame_main, text="Seleccionar", command=lambda: seleccionar_archivo_salida(entry_salida3))
btn_salida3.grid(row=3, column=2, padx=10, pady=5)

lbl_salida4 = tk.Label(frame_main, text="Salida 4 (XLSM, prefijo 010):", bg="#f0f8ff", font=label_font)
lbl_salida4.grid(row=4, column=0, sticky="e", padx=10, pady=5)
entry_salida4 = tk.Entry(frame_main, width=60)
entry_salida4.grid(row=4, column=1, padx=10, pady=5)
btn_salida4 = tk.Button(frame_main, text="Seleccionar", command=lambda: seleccionar_archivo_salida(entry_salida4))
btn_salida4.grid(row=4, column=2, padx=10, pady=5)

# Skiprows: Ingreso de filas de inicio.
lbl_skiprows = tk.Label(frame_main, text="Filas de inicio (separadas por coma):", bg="#f0f8ff", font=label_font)
lbl_skiprows.grid(row=5, column=0, sticky="e", padx=10, pady=5)
entry_skiprows = tk.Entry(frame_main, width=60)
entry_skiprows.grid(row=5, column=1, padx=10, pady=5)
entry_skiprows.insert(0, "48,62,76,90,104,119,133,147,161")  # Valor de ejemplo.

# Descripción para skiprows.
lbl_descripcion = tk.Label(frame_main, text="Descripción:", bg="#f0f8ff", font=label_font)
lbl_descripcion.grid(row=6, column=0, sticky="e", padx=10, pady=5)
entry_descripcion = tk.Entry(frame_main, width=60)
entry_descripcion.grid(row=6, column=1, padx=10, pady=5)

btn_guardar_skiprows = tk.Button(frame_main, text="Guardar Skiprows", 
                                 command=lambda: guardar_skiprows(entry_descripcion.get(), entry_skiprows.get()))
btn_guardar_skiprows.grid(row=6, column=2, padx=10, pady=5)

# Selección de skiprows guardados.
lbl_skiprows_guardados = tk.Label(frame_main, text="Skiprows Guardados:", bg="#f0f8ff", font=label_font)
lbl_skiprows_guardados.grid(row=7, column=0, sticky="e", padx=10, pady=5)
skiprows_var = tk.StringVar()
skiprows_opciones = ["{}: {}".format(desc, val) for _, desc, val in obtener_skiprows()]
if skiprows_opciones:
    skiprows_var.set(skiprows_opciones[0])
    skiprows_guardados = tk.OptionMenu(frame_main, skiprows_var, *skiprows_opciones)
else:
    skiprows_guardados = tk.OptionMenu(frame_main, skiprows_var, "")
skiprows_guardados.grid(row=7, column=1, padx=10, pady=5)

def cargar_skiprows_seleccionados(*args):
    """
    Al cambiar la selección en el OptionMenu de skiprows,
    carga el valor seleccionado en el campo entry_skiprows.
    """
    seleccion = skiprows_var.get()
    if seleccion:
        try:
            descripcion, valores = seleccion.split(": ", 1)
            entry_skiprows.delete(0, tk.END)
            entry_skiprows.insert(0, valores)
        except ValueError:
            pass

skiprows_var.trace("w", cargar_skiprows_seleccionados)

# --- Estado de Procesamiento ---
status_label = tk.Label(root, text="", font=label_font, bg="#f0f8ff", fg="blue")
status_label.pack(pady=10)

# --- Botón para iniciar el proceso de extracción y copia de datos ---
btn_procesar = tk.Button(root, text="Procesar Archivos", bg="#003366", fg="white", font=label_font, command=procesar_archivos)
btn_procesar.pack(pady=20)

# --- Botón de salida ---
btn_salir = tk.Button(root, text="Salir", bg="#cc0000", fg="white", font=label_font, command=root.destroy)
btn_salir.pack(pady=10)

# Iniciar el loop principal de la aplicación.
root.mainloop()
