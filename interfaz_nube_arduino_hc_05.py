import tkinter as tk
import serial
import serial.tools.list_ports
import threading
import openpyxl
import os
from datetime import datetime
import subprocess
import gspread
from oauth2client.service_account import ServiceAccountCredentials


# ─── CONFIGURACIÓN ─────────────────────
EXCEL_PATH = "datos_sensores.xlsx"  # Dentro del repositorio clonado

# ─── DETECCIÓN AUTOMÁTICA DEL PUERTO ───
def detectar_puerto_bluetooth():
    puertos = serial.tools.list_ports.comports()
    for puerto in puertos:
        descripcion = puerto.description.lower()
        if "bluetooth" in descripcion or "hc-05" in descripcion or "usb serial" in descripcion:
            return puerto.device
    return None

PUERTO_BLUETOOTH = detectar_puerto_bluetooth()
BAUDIOS = 9600

if PUERTO_BLUETOOTH is None:
    print("No se detectó el módulo Bluetooth HC-05.")
    exit()

try:
    ser = serial.Serial(PUERTO_BLUETOOTH, BAUDIOS, timeout=1)
    print(f"Conectado al puerto {PUERTO_BLUETOOTH}")
except serial.SerialException:
    print(f"Error al abrir el puerto {PUERTO_BLUETOOTH}")
    exit()

# ─── NOMBRES HUMANOS DE LOS SENSORES ───
nombres_sensores = {
    'T': "Temp DHT11 (°C)",
    'D': "Distancia (cm)",
    'G': "Gas MQ-2 (ppm aprox.)",
    'H': "Humedad (%)"
}

# ─── INTERFAZ ───────────────────────────
ventana = tk.Tk()
ventana.title("Lectura Automática de Sensores (Bluetooth)")

etiquetas = {}
for clave, nombre in nombres_sensores.items():
    lbl = tk.Label(ventana, text=f"{nombre}: ---", font=("Arial", 14))
    lbl.pack(pady=5)
    etiquetas[clave] = lbl

# ─── VARIABLES GLOBALES ────────────────
datos_actuales = {clave: '---' for clave in nombres_sensores}

# ─── LECTURA SERIAL ────────────────────
def leer_datos():
    while True:
        if ser.in_waiting:
            try:
                linea = ser.readline().decode().strip()
                partes = linea.split(',')
                for par in partes:
                    if ':' in par:
                        clave, valor = par.split(':')
                        clave = clave.strip()
                        valor = valor.strip()
                        if clave in nombres_sensores:
                            datos_actuales[clave] = valor
                            etiquetas[clave]['text'] = f"{nombres_sensores[clave]}: {valor}"
            except:
                continue

# ─── CREAR ARCHIVO EXCEL SI NO EXISTE ───────
def crear_excel_si_no_existe():
    if not os.path.exists(EXCEL_PATH):
        libro = openpyxl.Workbook()
        hoja = libro.active
        hoja.title = "DatosSensores"
        hoja.append(["Fecha", "Hora", "Temp DHT11", "Distancia", "Gas MQ-2", "Humedad"])
        libro.save(EXCEL_PATH)
        print("Archivo Excel creado con encabezados.")

# ─── GUARDAR EN EXCEL + GIT PUSH ───────
def grabar_datos():
    crear_excel_si_no_existe()

    fecha = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")

    fila = [
        fecha,
        hora,
        datos_actuales['T'],
        datos_actuales['D'],
        datos_actuales['G'],
        datos_actuales['H']
    ]

    # Cargar y guardar antes de git
    libro = openpyxl.load_workbook(EXCEL_PATH)
    hoja = libro.active
    hoja.append(fila)
    libro.save(EXCEL_PATH)  #Guardar antes del git add

    # Confirmación por consola
    print("Datos grabados en el archivo Excel correctamente.")
    
    # Guardar también en Google Sheets
    guardar_en_google_sheets(fila)

    # Construcción del mensaje
    mensaje = f"✔ Datos guardados:\nFecha: {fecha}\nHora: {hora}\n"
    mensaje += f"{nombres_sensores['T']}: {datos_actuales['T']} °C\n"
    mensaje += f"{nombres_sensores['D']}: {datos_actuales['D']} cm\n"
    mensaje += f"{nombres_sensores['G']}: {datos_actuales['G']} ppm aprox.\n"
    mensaje += f"{nombres_sensores['H']}: {datos_actuales['H']} %"

    # Mostrar en la interfaz
    notificacion.config(text=mensaje, fg="green")

    # Comandos Git en orden
    try:
        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", f"Actualización {datetime.now().isoformat()}"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("Cambios subidos a GitHub correctamente.")
    except subprocess.CalledProcessError as e:
        notificacion.config(text="Error al subir a GitHub", fg="red")
        print(f"Git error: {e}")

# ─── GUARDAR EN GOOGLE SHEETS ───────
def guardar_en_google_sheets(fila):
    try:
        # Autenticación
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credenciales = ServiceAccountCredentials.from_json_keyfile_name("credenciales_google.json", scope)
        cliente = gspread.authorize(credenciales)

        # Abrir hoja de cálculo por nombre (colocar nombre de hoja de cálculo)
        hoja = cliente.open("Datos_Sensores_Arduino").sheet1
        encabezados = ["Fecha", "Hora", "Temp DHT11", "Distancia", "Gas MQ-2", "Humedad"]
        primera_fila = hoja.row_values(1)

        # Verificar si ya existen encabezados
        if primera_fila != encabezados:
            hoja.insert_row(encabezados, 1)
            print("Encabezados escritos en Google Sheets.")

        # Agregar la fila de datos
        hoja.append_row(fila)
        print("Datos también guardados en Google Sheets.")
    except Exception as e:
        print(f"Error al guardar en Google Sheets: {e}")
        notificacion.config(text="Error en Google Sheets", fg="orange")

# ─── BOTÓN DE GRABAR ───────────────────
boton = tk.Button(ventana, text="Grabar", font=("Arial", 14), command=grabar_datos)
boton.pack(pady=10)

# Área de notificación visual
notificacion = tk.Label(ventana, text="", font=("Arial", 12), fg="green", justify="left")
notificacion.pack(pady=10)

# ─── INICIAR HILO DE LECTURA ───────────
hilo = threading.Thread(target=leer_datos, daemon=True)
hilo.start()

# ─── MOSTRAR INTERFAZ ──────────────────
ventana.mainloop()
