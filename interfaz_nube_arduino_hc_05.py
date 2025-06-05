import tkinter as tk
import serial
import serial.tools.list_ports
import threading
import openpyxl
import os
from datetime import datetime
import subprocess

# ─── CONFIGURACIÓN ─────────────────────
EXCEL_PATH = "datos_sensores.xlsx"  # Debe estar dentro del repositorio clonado

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
    print("❌ No se detectó el módulo Bluetooth HC-05.")
    exit()

try:
    ser = serial.Serial(PUERTO_BLUETOOTH, BAUDIOS, timeout=1)
    print(f"✅ Conectado al puerto {PUERTO_BLUETOOTH}")
except serial.SerialException:
    print(f"❌ Error al abrir el puerto {PUERTO_BLUETOOTH}")
    exit()

# ─── NOMBRES HUMANOS DE LOS SENSORES ───
nombres_sensores = {
    'T1': "Temp DHT11 (°C)",
    'D':  "Distancia (cm)",
    'T2': "Temp LM35 (°C)",
    'P':  "Presión (hPa)"
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

# ─── GUARDAR EN EXCEL + GIT PUSH ───────
def grabar_datos():
    fila = [
        datetime.now().strftime("%Y-%m-%d"),
        datetime.now().strftime("%H:%M:%S"),
        datos_actuales['T1'],
        datos_actuales['D'],
        datos_actuales['T2'],
        datos_actuales['P']
    ]

    if os.path.exists(EXCEL_PATH):
        libro = openpyxl.load_workbook(EXCEL_PATH)
        hoja = libro.active
    else:
        libro = openpyxl.Workbook()
        hoja = libro.active
        hoja.append(["Fecha", "Hora", "Temp DHT11", "Distancia", "Temp LM35", "Presión"])

    hoja.append(fila)
    libro.save(EXCEL_PATH)

    try:
        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", "Datos sensores actualizados"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("✅ Datos subidos a GitHub.")
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Error al subir a GitHub: {e}")

# ─── BOTÓN DE GRABAR ───────────────────
boton = tk.Button(ventana, text="Grabar", font=("Arial", 14), command=grabar_datos)
boton.pack(pady=10)

# ─── INICIAR HILO DE LECTURA ───────────
hilo = threading.Thread(target=leer_datos, daemon=True)
hilo.start()

# ─── MOSTRAR INTERFAZ ──────────────────
ventana.mainloop()
