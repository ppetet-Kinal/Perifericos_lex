import os
import platform
import psutil
import socket
import subprocess
import tkinter as tk
from tkinter import messagebox
import win32com.client


def obtener_nombre_procesador():
    try:
        comando = 'powershell "Get-CimInstance Win32_Processor | Select-Object -ExpandProperty Name"'
        resultado = subprocess.run(comando, capture_output=True, text=True, shell=True)
        procesador = resultado.stdout.strip()
        return procesador if procesador else "No disponible"
    except Exception as e:
        return f"No disponible ({e})"


def obtener_tipo_disco():
    try:
        wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        service = wmi.ConnectServer(".", "root\\cimv2")
        query = "SELECT MediaType, Model, Size FROM Win32_DiskDrive"
        discos = service.ExecQuery(query)
        tipos_discos = []
        for disco in discos:
            tipo = disco.MediaType or "Desconocido"
            modelo = disco.Model or "Desconocido"
            tamaño_gb = round(int(disco.Size) / (1024 ** 3), 2) if disco.Size else "No disponible"
            tipos_discos.append(f"Modelo: {modelo}, Tipo: {tipo}, Tamaño: {tamaño_gb} GB")
        return "\n".join(tipos_discos)
    except Exception as e:
        return f"No se pudo obtener la información del disco ({e})"


def obtener_informacion_sistema():
    try:
        nombre_equipo = platform.node()
        procesador = obtener_nombre_procesador()
        ram_instalada = round(psutil.virtual_memory().total / (1024 ** 3), 2)
        tipo_disco = obtener_tipo_disco()
        version_windows = platform.version()
        try:
            # La función win32_edition puede no estar disponible en algunas plataformas.
            edicion_windows = platform.win32_edition()
        except AttributeError:
            edicion_windows = "No disponible"
        try:
            nombre_host = socket.gethostname()
            direccion_ip = socket.gethostbyname(nombre_host)
        except socket.error:
            direccion_ip = "No disponible"
        informacion = (
            f"Nombre del equipo: {nombre_equipo}\n"
            f"Procesador: {procesador}\n"
            f"RAM instalada: {ram_instalada} GB\n"
            f"Versión de Windows: {version_windows} ({edicion_windows})\n"
            f"Dirección IP: {direccion_ip}\n"
            f"Discos Duros:\n{tipo_disco}"
        )
        return informacion
    except Exception as e:
        return f"Error al obtener la información del sistema:\n{e}"


def mostrar_informacion():
    informacion = obtener_informacion_sistema()
    text_widget.config(state="normal")
    text_widget.delete(1.0, tk.END)
    text_widget.insert(tk.END, informacion)
    text_widget.config(state="disabled")


def enviar_datos():
    nombre = entrada_nombre.get()
    codigo = entrada_codigo.get()

    if not nombre or not codigo:
        messagebox.showwarning("Campos vacíos", "Por favor, ingrese su nombre y código de empleado.")
        return

    informacion_sistema = obtener_informacion_sistema()
    datos_usuario = f"Nombre: {nombre}\nCódigo de empleado: {codigo}\n"
    datos_completos = f"{datos_usuario}\n{informacion_sistema}"

    messagebox.showinfo("Datos enviados", "Los datos se han enviado correctamente.\n\n" + datos_completos)

    try:
        with open("datos_sistema.txt", "w", encoding="utf-8") as archivo:
            archivo.write(datos_completos)
    except Exception as e:
        messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo:\n{e}")


# Configuración de la ventana principal
root = tk.Tk()
root.title("Lexcom TI")
root.geometry("500x500")

# Ruta del icono
ruta_icono = r"C:\Users\Soporte 2\Documents\Phyton\Lexcom 1.ico"
if os.path.exists(ruta_icono):
    root.iconbitmap(ruta_icono)

# Título
label_titulo = tk.Label(root, text="LEXCOM SUPPORT", font=("Adobe Garamond Pro", 18))
label_titulo.pack(pady=10)

# Entrada de datos del usuario
label_nombre = tk.Label(root, text="Nombre del empleado:")
label_nombre.pack()
entrada_nombre = tk.Entry(root, width=40)
entrada_nombre.pack(pady=5)

label_codigo = tk.Label(root, text="Código de empleado:")
label_codigo.pack()
entrada_codigo = tk.Entry(root, width=40)
entrada_codigo.pack(pady=5)

# Botón para mostrar la información del sistema
boton_ver_info = tk.Button(root, text="Ver información del sistema", command=mostrar_informacion)
boton_ver_info.pack(pady=10)

# Botón para enviar la información
boton_enviar = tk.Button(root, text="Enviar información", command=enviar_datos)
boton_enviar.pack(pady=15)

# Text widget para mostrar información del sistema (con barra de desplazamiento)
frame_texto = tk.Frame(root)
frame_texto.pack(padx=10, pady=10, fill="both", expand=True)
text_widget = tk.Text(frame_texto, wrap="word", state="disabled", font=("Courier", 10), height=10)
scroll_bar = tk.Scrollbar(frame_texto, command=text_widget.yview)
text_widget.config(yscrollcommand=scroll_bar.set)
text_widget.pack(side=tk.LEFT, fill="both", expand=True)
scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)

root.mainloop()
