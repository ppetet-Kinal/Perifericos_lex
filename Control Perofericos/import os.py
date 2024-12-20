import os
import platform
import psutil
import socket
import subprocess
import tkinter as tk
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
        return f"No se pudo obtener la información general del disco"

def obtener_informacion_sistema():
    try:
        nombre_equipo = platform.node()
        procesador = obtener_nombre_procesador()
        ram_instalada = round(psutil.virtual_memory().total / (1024 ** 3), 2)
        tipo_disco = obtener_tipo_disco()
        version_windows = platform.version()
        try:
            edicion_windows = platform.win32_edition()
        except:
            edicion_windows = "No disponible"
        try:
            nombre_host = socket.gethostname()
            direccion_ip = socket.gethostbyname(nombre_host)
        except:
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

# Crear la ventana principal
root = tk.Tk()
root.title("Lexcom TI")
root.geometry("500x300")

# Ruta del icono
ruta_icono = r"C:\Phyton\1-d57dee55.ico"

# Asignar el icono
try:
    if os.path.exists(ruta_icono):
        root.iconbitmap(ruta_icono)
    else:
        print(f"El icono no existe en la ruta {ruta_icono}.")
except Exception as e:
    print(f"No se pudo cargar el icono: {e}")

# Configurar la interfaz
label = tk.Label(root, text="LEXCOM SUPPORT", font=("Adobe Garamond Pro", 18))
label.pack(pady=10)

boton = tk.Button(root, text="Ver información", command=mostrar_informacion)
boton.pack(pady=5)

text_widget = tk.Text(root, wrap="word", state="disabled", font=("Courier", 10))
text_widget.pack(padx=10, pady=10, fill="both", expand=True)

root.mainloop()