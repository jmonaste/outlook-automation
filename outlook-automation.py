import win32com.client
import csv
import os

# Crea un objeto para interactuar con la aplicación Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Obtén la carpeta de la bandeja de entrada
inbox = outlook.GetDefaultFolder(6)

# Obtén los correos electrónicos no leídos de la bandeja de entrada que tienen 'correo a tratar' en el asunto
unread_messages = inbox.Items.Restrict("[Unread]=True AND [Subject]='correo a tratar'")

# Crea el archivo CSV
ouput_file_path = os.path.abspath('correos.csv')

with open(ouput_file_path, mode="w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["ID", "Fecha"])

    # Imprime el contenido de los correos electrónicos no leídos y guarda la información en el archivo CSV
    for message in unread_messages:
        body = message.Body
        id_pos = body.index("id: ") + 4
        fecha_pos = body.index("fecha: ") + 7
        id = body[id_pos:fecha_pos-8].strip()
        fecha = body[fecha_pos:].strip()

        writer.writerow([id, fecha])

        print("Asunto:", message.Subject)
        print("Cuerpo:", message.Body)
        print("----------------------------------------")
