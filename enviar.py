import pyautogui, webbrowser
from time import sleep
import pandas as pd #para leer nuemros de un archivo excel
import time


# FUNCION PARA CONVERTIR DE HORA MILITAR A HORA NORMAL
def convertir_hora(hora_militar):
    partes = hora_militar.split(":")
    horas = int(partes[0])
    minutos = int(partes[1])
    segundos = int(partes[2])

    if horas < 0 or horas > 23 or minutos < 0 or minutos > 59 or segundos < 0 or segundos > 59:
        return "Hora inválida"

    if horas == 0:
        hora_normal = "12"
        sufijo = "AM"
    elif horas < 12:
        hora_normal = str(horas)
        sufijo = "AM"
    elif horas == 12:
        hora_normal = str(horas)
        sufijo = "PM"
    else:
        hora_normal = str(horas - 12)
        sufijo = "PM"

    hora_normal += ":" + str(minutos).zfill(2) + ":" + str(segundos).zfill(2)

    return hora_normal + " " + sufijo

def enviarCita():
    df = pd.read_excel('citas.xlsx')
    nombre = df['Usuario_Asigna'].tolist() # Nombre del usuario
    cita = df['Especialidad'].tolist() # Especialidad de la cita
    profesional = df['Profesional'].tolist() # Profesional
    fecha = df['Fecha_Cita'].tolist() # Fecha de la cita
    hora = df['Hora'].tolist() # Hora de la cita
    numero = df['Tel'].tolist() # Telefono del usuario
    sede = df['Lugar'].tolist() # Lugar de cita
    for nom,cit,prof,fech,hor,num,sed in zip(nombre,cita,profesional,fecha,hora,numero,sede):

        # eliminar caracteres extras en fecha
        fech = str(fech)
        fech = fech.replace("00:00:00","")
        # convertir en hora normal
        hor = str(hor)
        hor = convertir_hora(hor)
        texto = f"Usuario {nom}. La E.S.E Hospital Divino Salvador de Sopo le recuerda que su cita de {cit} con el profesional {prof} esta programada para el dia {fech} a las {hor} en la sede {sed}. En caso de no poder asistir por favor comunicarse mínimo 12 horas antes para la respectiva cancelacion. Recuerde llegar 20 minutos antes de la cita para realizar el proceso de facturacion. Hospital Divino Salvador de Sopo por un servicio mas humano"
        telefono=int(num)
        telefono = str(telefono)
        webbrowser.open('https://web.whatsapp.com/send?phone=+57'+telefono)

        sleep(15)

        # PEGA LA IMAGEN COPIADA EN EL PORTAPAPELES
        pyautogui.hotkey("ctrl", "v")
        sleep(5)
        pyautogui.typewrite(texto)
        sleep(5)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey("ctrl", "w")


def enviar(texto,opcion = 1):
    df = pd.read_excel('numeros.xlsx')
    num = df['Tel'].tolist()

    for i in num:
        telefono=int(i)
        telefono = str(telefono)
        webbrowser.open('https://web.whatsapp.com/send?phone=+57'+telefono)

        sleep(15)

        if(opcion == "1"): # envia el mensaje accompañado de un texto
            pyautogui.hotkey("ctrl", "v")
            sleep(5)
        pyautogui.typewrite(texto)
        sleep(5)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey("ctrl", "w")

def TomarTexto():
    print("--------------------------------------------------------------------------------------------")
    print("ESCRIBE EL MENSAJE A ENVIAR")
    texto = input()
    print("EL TEXTO A ENVIAR ES EL SIGUIENTE:")
    print(texto)
    print("--------------------------------------------------------------------------------------------")
    return texto
valor = 0

while valor == 0:

    # MENU INICIAL
    print("--------------------------------------------------------------------------------------------")
    print("Bienvenido al programa de envio de whatsapp masivos")
    print("Para enviar los recordatorios de citas, mete los datos en el archivo 'citas.xlsx'")
    print("Para envio de mensajes personalizados con y sin imagenes ingresa los numeros en 'numeros.xlsx'")
    print("--------------------------------------------------------------------------------------------")
    print("-----------------------------RECUERDA TENER EL WHATSAPP ABIERTO-----------------------------")
    print("--------------------------------------------------------------------------------------------")
    print("SELECCIONA UNA OPCION POR FAVOR")
    print("1. PARA ENVIAR MENSAJE CON IMAGEN")
    print("2. PARA ENVIAR MENSAJE CON SOLO TEXTO")
    print("3. PARA ENVIAR RECORDATORIO DE CITAS AUTOMATICO")
    print("5. SI PRESENTAS ERROR CON EL WHATSAPP WEB")

    opcion = input()
    print(opcion)
    texto = ""

    if(opcion=="1"): # Enviar mensaje con imagen
        validar = "2"
        # Ingresar texto a enviar por whatsapp
        while validar != "1":
            print("COPIA LA IMAGEN CON 'CRLT + C' O DESDE EL NAVEGADOR CLICK DERECHO Y COPIAR")
            print("CUANDO LA TENGAS EN SU LUGAR PRESIONA ENTER....")
            x = input()
            sleep(5)
            texto = TomarTexto()
            print("SI ES CORRECTO PRESIONA '1' DE LO CONTRARIO PRESIONA '2'")
            validar = input()
        enviar(texto,opcion)
        valor = 1
    elif(opcion=="2"): # Enviar mensaje sin imagen
        validar = "2"
        # Ingresar texto a enviar por whatsapp
        while validar != "1":
            texto = TomarTexto()
            print("SI ES CORRECTO PRESIONA '1' DE LO CONTRARIO PRESIONA '2'")
            validar = input()
        enviar(texto,opcion)
        valor = 1
    elif(opcion=="3"): # enviar recordatorios de citas
        print("NECESITAS UNA IMAGEN")
        print("COPIA LA IMAGEN CON 'CRLT + C' O DESDE EL NAVEGADOR CLICK DERECHO Y COPIAR")
        print("CUANDO LA TENGAS EN SU LUGAR PRESIONA ENTER....")
        x = input()
        enviarCita()
        valor = 1
    elif(opcion=="5"): # Arreglo de problema con el whatsapp
        print("SE ABRIRA UNA VENTANA WEB, EJECUTA EL PROGRAMA NUEVAMENTE CUANDO SE CARGUE LA PAGINA")
        time.sleep(5)
        webbrowser.open('https://web.whatsapp.com/send?phone=+573168308100')
        valor = 1
    else:
        print("OPCION INCORRECTA VUELVA A INTENTARLO")
    time.sleep(5)

