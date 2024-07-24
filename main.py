import imaplib
import email
import time
import datetime
import os
import logging
from dotenv import load_dotenv
import tkinter as tk
from tkinter import ttk
import pygame  # Importar pygame para reproducir sonidos
from threading import Thread

load_dotenv()

# Obtener la ruta actual del script
ruta_script = os.path.dirname(os.path.abspath(__file__))

# Crear la carpeta "logs" si no existe
ruta_logs = os.path.join(ruta_script, 'logs')
if not os.path.exists(ruta_logs):
    os.makedirs(ruta_logs)

# Función para crear un nuevo archivo de log con la fecha actual
def crear_nuevo_log():
    global ruta_archivo_log
    global logger
    if logger:
        logger.removeHandler(file_handler)
        file_handler.close()
    nombre_archivo_log = time.strftime('logs-%Y-%m-%d.txt')
    ruta_archivo_log = os.path.join(ruta_logs, nombre_archivo_log)
    file_handler = logging.FileHandler(ruta_archivo_log)
    logger.addHandler(file_handler)

# Configuración de logging para guardar en un archivo específico
nombre_archivo_log = time.strftime('logs-%Y-%m-%d.txt')
ruta_archivo_log = os.path.join(ruta_logs, nombre_archivo_log)
logging.basicConfig(filename=ruta_archivo_log, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

username = os.getenv("OPERADOR_EMAIL")
password = os.getenv("OPERADOR_PASSWORD")
imap_host = os.getenv("IMAP_SERVER")
imap_port = 993
label = os.getenv("GMAIL_LABEL")  # Nombre de la etiqueta específica
highlight_color = "#FF8E8E"  # Color de fondo específico para correos recientes

# Inicializar el logger
logger = logging.getLogger()
file_handler = logging.FileHandler(ruta_archivo_log)
logger.addHandler(file_handler)

dia_actual = datetime.datetime.now().day

# Crear la ventana principal de Tkinter
root = tk.Tk()
root.title("Correo Electrónico")
root.geometry("600x400")

# Crear el Treeview para mostrar los correos
tree = ttk.Treeview(root, columns=("Asunto", "Hora"), show="headings")
tree.heading("Asunto", text="Asunto")
tree.heading("Hora", text="Hora")
tree.pack(fill=tk.BOTH, expand=True)

# Inicializar pygame para reproducir sonidos
pygame.mixer.init()

# Tamaño de la fuente inicial
font_size = 10

# Función para cambiar el tamaño de la fuente
def cambiar_tamaño_fuente(aumento):
    global font_size
    if aumento:
        font_size += 1
    else:
        font_size -= 1
    tree.tag_configure("default", font=("Arial", font_size))
    tree.tag_configure("highlight", font=("Arial", font_size))

# Botones para aumentar y disminuir el tamaño de la fuente
btn_aumentar = tk.Button(root, text="+", command=lambda: cambiar_tamaño_fuente(True))
btn_aumentar.pack(side=tk.LEFT)
btn_disminuir = tk.Button(root, text="-", command=lambda: cambiar_tamaño_fuente(False))
btn_disminuir.pack(side=tk.LEFT)

# Función para reproducir el sonido
def reproducir_sonido():
    pygame.mixer.music.load('alerta.mp3')  # Ruta al archivo de sonido
    pygame.mixer.music.play()

# Función para actualizar la lista de correos en la interfaz de Tkinter
def actualizar_lista(asunto, hora):
    now = datetime.datetime.now()
    email_time = datetime.datetime.strptime(hora, '%Y-%m-%d %H:%M:%S')
    tag = "default"
    if (now - email_time).total_seconds() < 300:  # Menos de 5 minutos
        tag = "highlight"
    tree.insert("", "end", values=(asunto, hora), tags=(tag,))
    Thread(target=reproducir_sonido).start()

# Configurar tags para el Treeview
tree.tag_configure("default", font=("Arial", font_size))
tree.tag_configure("highlight", background=highlight_color, font=("Arial", font_size))

def revisar_correos():
    global dia_actual

    while True:
        # Verificar si ha cambiado el día
        if datetime.datetime.now().day != dia_actual:
            crear_nuevo_log()
            dia_actual = datetime.datetime.now().day

        try:
            # Conéctate a Gmail a través de IMAP
            mail = imaplib.IMAP4_SSL(imap_host, imap_port)
            mail.login(username, password)
            mail.select(f'"{label}"')  # Selecciona la etiqueta específica
            logging.info(f'Conexión exitosa a Gmail y etiqueta "{label}" seleccionada')

            # Busca correos no leídos en la etiqueta específica
            result, data = mail.search(None, "UNSEEN")
            if result == 'OK':
                for num in data[0].split():
                    result, data = mail.fetch(num, '(RFC822)')
                    raw_email = data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    logging.info('Correo recibido')

                    # Extrae el asunto y el cuerpo del correo
                    subject = msg['subject']
                    hora = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                    # Actualizar la lista de correos en la interfaz de Tkinter
                    actualizar_lista(subject, hora)

            # Actualizar colores de los correos existentes
            for item in tree.get_children():
                values = tree.item(item, "values")
                email_time = datetime.datetime.strptime(values[1], '%Y-%m-%d %H:%M:%S')
                if (datetime.datetime.now() - email_time).total_seconds() >= 300:
                    tree.item(item, tags=("default",))

            mail.logout()
        except imaplib.IMAP4.error as e:
            print(f"Error al procesar el correo: {e}")
            logging.error(f'Error al procesar el correo: {e}')
            time.sleep(60)  # Espera 60 segundos antes de intentar reconectar
        except Exception as e:
            logging.error(f'Error al procesar el correo: {e}')
        # Espera 1 segundo antes de revisar el correo nuevamente
        time.sleep(1)
        logging.info('Esperando 1 segundo antes de revisar el correo nuevamente')

# Ejecutar la función de revisar correos en un hilo separado
Thread(target=revisar_correos).start()

# Iniciar el bucle principal de Tkinter
root.mainloop()
