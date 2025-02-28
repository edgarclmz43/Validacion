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
from tkinter import filedialog, messagebox, font  # Módulos específicos de Tkinter para diálogos, mensajes y configuración de fuentes.
import os            # Biblioteca para interactuar con el sistema operativo (trabajo con archivos y directorios).

# --- Funciones de selección de archivos y carpetas ---

def seleccionar_carpeta_origen(entry):
    """
    Permite al usuario seleccionar la carpeta que contiene los archivos de origen (XLSX).
    
    Parámetros:
      entry: widget Entry donde se mostrará la ruta seleccionada.
    """
    folder_path = filedialog.askdirectory(title="Seleccionar Carpeta de Origen (XLSX)")
    if folder_path:
        entry.delete(0, tk.END)      # Limpia cualquier contenido previo en el widget.
        entry.insert(0, folder_path) # Inserta la ruta seleccionada.

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
        entry.delete(0, tk.END)      # Limpia cualquier contenido previo en el widget.
        entry.insert(0, filepath)    # Inserta la ruta del archivo seleccionado.

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
            # Se asume que la hoja de interés se llama "sr ysR(Método) ISO 5725".
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
                # Se guarda cada valor de la fila en una lista.
                datos.append([cel.value for cel in fila])
            bloques_datos.append(datos)
        wb_origen.close()  # Cerrar el archivo de origen.

        # 11. Para cada bloque, abrir el archivo de salida correspondiente y pegar los datos.
        # Asignación:
        #   Bloque 1 -> salida_files[0] (prefijo "01")
        #   Bloque 2 -> salida_files[1] (prefijo "04")
        #   Bloque 3 -> salida_files[2] (prefijo "07")
        #   Bloque 4 -> salida_files[3] (prefijo "010")
        for bloque_idx, bloque in enumerate(bloques_datos):
            salida_path = salida_files[bloque_idx]
            # Abrir el archivo de salida con xlwings (para preservar macros y formatos).
            try:
                app = xw.App(visible=False)  # Se inicia Excel en modo no visible.
                wb_salida = app.books.open(salida_path)
                # Se asume que la hoja de destino se llama "Registro".
                ws_salida = wb_salida.sheets["Registro"]
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo de salida {os.path.basename(salida_path)}:\n{e}")
                continue

            try:
                # 12. Por cada una de las 3 columnas del bloque, buscar la clave en la columna D del archivo de salida.
                for c in range(3):
                    clave = claves[c]
                    fila_encontrada = None
                    # Se recorre un rango de 20 filas a partir de 'fila_inicio' para buscar la coincidencia.
                    for fila in range(fila_inicio, fila_inicio + 20):
                        valor_celda = ws_salida.range("D" + str(fila)).value
                        if valor_celda == clave:
                            fila_encontrada = fila
                            break
                    if fila_encontrada is None:
                        messagebox.showwarning("Advertencia", f"No se encontró la clave '{clave}' en {os.path.basename(salida_path)} a partir de la fila {fila_inicio}.")
                        continue

                    # 13. Una vez encontrada la fila, pegar los 10 valores del bloque (de la columna c) a partir de la columna E.
                    fila_destino = fila_encontrada
                    col_inicial = 5  # La columna E corresponde al índice 5.
                    for r in range(10):
                        try:
                            valor = bloque[r][c]
                        except IndexError:
                            valor = None  # Si no existe el valor, se asigna None.
                        ws_salida.range((fila_destino, col_inicial + r)).value = valor

                # 14. Guardar y cerrar el archivo de salida.
                wb_salida.save()
                wb_salida.close()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar {os.path.basename(salida_path)}:\n{e}")
            finally:
                app.quit()  # Cerrar la instancia de Excel para liberar recursos.

    # 15. Notificar al usuario que el proceso finalizó correctamente.
    messagebox.showinfo("Proceso Finalizado", "Se han procesado los archivos de origen correctamente.")


# --- Funciones para manejar la base de datos de skiprows usando SQLite ---
import sqlite3  # Biblioteca para interactuar con bases de datos SQLite.

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

# Configuración básica de la ventana principal.
root = tk.Tk()
root.title("Copiador de Bloques – Origen por Carpeta y 4 Salidas Fijas")
root.geometry("800x650")         # Tamaño de la ventana.
root.configure(bg="#f0f8ff")      # Color de fondo claro.

# Definición de fuentes personalizadas para el título, subtítulo, instrucciones y etiquetas.
titulo_font = font.Font(family="Helvetica", size=16, weight="bold")
subtitulo_font = font.Font(family="Helvetica", size=12, weight="bold")
instrucciones_font = font.Font(family="Helvetica", size=10)
label_font = font.Font(family="Helvetica", size=10, weight="bold")

# --- Marco superior para el título, propósito e instrucciones ---
frame_top = tk.Frame(root, bg="#f0f8ff", pady=10)
frame_top.pack(fill="x")

# Título principal de la aplicación.
lbl_titulo = tk.Label(frame_top, text="Copiador de Bloques – Origen por Carpeta y 4 Salidas Fijas", 
                      bg="#f0f8ff", fg="#003366", font=titulo_font)
lbl_titulo.pack()

# Sección que indica el propósito de la aplicación.
lbl_proposito = tk.Label(frame_top, 
                         text="Propósito: Extraer bloques de datos de certificados en XLSX y copiarlos en archivos de salida XLSM manteniendo macros y formatos.",
                         bg="#f0f8ff", fg="#003366", font=subtitulo_font, wraplength=750, justify="center")
lbl_proposito.pack(pady=5)

# Instrucciones de uso detalladas para el usuario.
instrucciones = (
    "Instrucciones de Uso:\n"
    "1. Seleccione la carpeta que contiene los archivos de origen (XLSX).\n"
    "2. Seleccione los 4 archivos de salida (XLSM) correspondientes a cada bloque:\n"
    "   - Salida 1: Bloque 1 (prefijo '01')\n"
    "   - Salida 2: Bloque 2 (prefijo '04')\n"
    "   - Salida 3: Bloque 3 (prefijo '07')\n"
    "   - Salida 4: Bloque 4 (prefijo '010')\n"
    "3. Ingrese las filas de inicio separadas por comas para cada archivo de origen.\n"
    "4. Haga clic en 'Procesar Archivos' para iniciar la extracción y copia de datos."
)
lbl_instrucciones = tk.Label(frame_top, text=instrucciones, bg="#f0f8ff", fg="#333333", 
                             font=instrucciones_font, justify="left")
lbl_instrucciones.pack(pady=5, padx=10)

# --- Marco principal para los controles de entrada ---
frame_main = tk.Frame(root, bg="#f0f8ff", padx=20, pady=10)
frame_main.pack(fill="both", expand=True)

# Campo para seleccionar la carpeta de origen (archivos XLSX).
lbl_origen = tk.Label(frame_main, text="Carpeta de Origen (XLSX):", bg="#f0f8ff", font=label_font)
lbl_origen.grid(row=0, column=0, sticky="e", padx=10, pady=5)
entry_origen = tk.Entry(frame_main, width=60)
entry_origen.grid(row=0, column=1, padx=10, pady=5)
btn_origen = tk.Button(frame_main, text="Seleccionar Carpeta", command=lambda: seleccionar_carpeta_origen(entry_origen))
btn_origen.grid(row=0, column=2, padx=10, pady=5)

# Campos para seleccionar los archivos de salida.
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

# --- Sección para manejar los skiprows mediante una base de datos SQLite ---
# Campo para ingresar las filas de inicio (skiprows) para la búsqueda.
lbl_skiprows = tk.Label(frame_main, text="Filas de inicio (separadas por coma):", bg="#f0f8ff", font=label_font)
lbl_skiprows.grid(row=5, column=0, sticky="e", padx=10, pady=5)
entry_skiprows = tk.Entry(frame_main, width=60)
entry_skiprows.grid(row=5, column=1, padx=10, pady=5)
entry_skiprows.insert(0, "48,62,76,90,104,119,133,147,161")  # Valor de ejemplo.

# Campo para ingresar la descripción de los skiprows.
lbl_descripcion = tk.Label(frame_main, text="Descripción:", bg="#f0f8ff", font=label_font)
lbl_descripcion.grid(row=6, column=0, sticky="e", padx=10, pady=5)
entry_descripcion = tk.Entry(frame_main, width=60)
entry_descripcion.grid(row=6, column=1, padx=10, pady=5)

# Botón para guardar los skiprows en la base de datos.
btn_guardar_skiprows = tk.Button(frame_main, text="Guardar Skiprows", 
                                 command=lambda: guardar_skiprows(entry_descripcion.get(), entry_skiprows.get()))
btn_guardar_skiprows.grid(row=6, column=2, padx=10, pady=5)

# Campo para seleccionar entre los skiprows guardados.
lbl_skiprows_guardados = tk.Label(frame_main, text="Skiprows Guardados:", bg="#f0f8ff", font=label_font)
lbl_skiprows_guardados.grid(row=7, column=0, sticky="e", padx=10, pady=5)
skiprows_var = tk.StringVar()
# Se crea una lista de opciones con el formato "descripción: valores".
skiprows_opciones = ["{}: {}".format(desc, val) for _, desc, val in obtener_skiprows()]
if skiprows_opciones:
    skiprows_var.set(skiprows_opciones[0])  # Establece la opción inicial.
    skiprows_guardados = tk.OptionMenu(frame_main, skiprows_var, skiprows_opciones[0], *skiprows_opciones[1:])
else:
    skiprows_guardados = tk.OptionMenu(frame_main, skiprows_var, "")
skiprows_guardados.grid(row=7, column=1, padx=10, pady=5)

def cargar_skiprows_seleccionados(*args):
    """
    Función que se dispara cuando se cambia la selección en el OptionMenu de skiprows.
    Carga el valor seleccionado en el campo entry_skiprows.
    """
    seleccion = skiprows_var.get()
    if seleccion:
        # Se asume que el formato es "descripción: valores".
        try:
            descripcion, valores = seleccion.split(": ", 1)
            entry_skiprows.delete(0, tk.END)
            entry_skiprows.insert(0, valores)
        except ValueError:
            pass

# Se establece una traza para que, al cambiar la selección, se actualice el campo entry_skiprows.
skiprows_var.trace("w", cargar_skiprows_seleccionados)

# --- Botón para iniciar el proceso de extracción y copia de datos ---
btn_procesar = tk.Button(root, text="Procesar Archivos", bg="#003366", fg="white", font=label_font, command=procesar_archivos)
btn_procesar.pack(pady=20)

# Iniciar el bucle principal de la aplicación Tkinter.
root.mainloop()
