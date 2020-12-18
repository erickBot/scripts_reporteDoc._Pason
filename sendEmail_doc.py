from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import openpyxl
import time
import os
from os import remove
from datetime import datetime, date, time, timedelta
import calendar

mylist = []

def run():
    enviarEmail = False

    name_file = "Control_Documentos.xlsx"
    file = openpyxl.load_workbook(name_file)
    sheet1 = file.get_sheet_by_name('Erick Pasache')
    sheet2 = file.get_sheet_by_name('Johnnel Panduro')
    sheet3 = file.get_sheet_by_name('A1I827')
    sheet4 = file.get_sheet_by_name('C9L911')

    lookCell(sheet1, 'C21', 'C22')
    lookCell(sheet2, 'C21', 'C22')
    lookCell(sheet3, 'C21', 'C22')
    lookCell(sheet4, 'C21', 'C22')

    print(mylist)
    #mira en la lista algun valor mayor que cero
    for i in mylist:
        if i > 0:
            enviarEmail = True

    print(enviarEmail)

    if (enviarEmail):
        procesaDestinatarios()

    file.close()


def lookCell(sheet, cellInit, cellEnd):
    i = 0
    cell_list = sheet[cellInit : cellEnd]
    for row in cell_list:
        for cell in row:
            mylist.append(cell.value)
            i += 1

def procesaDestinatarios():

    destinatarios = {
        "user1": "erickpasache0@gmail.com",
        "user2": "epasache_28@hotmail.com"
    }

    for i in destinatarios:
        sendEmail(destinatarios[i])

def sendEmail(destinatarios):
    #crea la instancia del objeto de mensaje
    msg = MIMEMultipart()
    message = "Hola buen dia \n\nSe envia lista actualizada de documentos de personal operativo. Hay\n\n" + str(mylist[0]) + " documentos de Erick por vencer,\n" + str(mylist[1]) + " documentos de Erick vencidos,\n" + str(mylist[2]) + " documentos de Johnnel por vencer,\n" + str(mylist[3]) + " documentos de Johnnel vencidos.\n\n" + str(mylist[4])+ " documentos de unidad A1I827 por vencer\n" + str(mylist[5])+ " documentos de unidad A1I827 vencidos\n" + str(mylist[6])+ " documentos de unidad C9L911 por vencer\n" + str(mylist[7]) + " documentos de unidad C9L911 vencidos\n\n" + "Soy un Bot."
    ruta_adjunto = "lista_Documentos_PersonalOperativo.xlsx"
    nombre_adjunto = "lista_Documentos_PersonalOperativo.xlsx"
    #configura los parametros del mensaje
    password = "99e12438cf"
    msg['From'] ="soyunbot2817@gmail.com"
    msg['To'] = destinatarios
    msg['Subject'] = "Lista de Documentos de Personal Operativo"
    #agrega el cuerpo del mensaje
    msg.attach(MIMEText(message, 'plain'))
    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(ruta_adjunto, 'rb')
    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    msg.attach(adjunto_MIME)
    #crear servidor
    server = smtplib.SMTP('smtp.gmail.com: 587')
    server.starttls()
    #Ingresa credenciales para enviar email
    server.login(msg['From'], password)
    #envia el mensaje al servidor
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    print ("Envio de email exitoso a %s:" % (msg['To']))

if __name__ == "__main__":
    run()
