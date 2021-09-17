from time import sleep
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.remote.webdriver import WebDriver as RemoteWebDriver
import re
from unicodedata import normalize
from excel_conexion import *
from send_mail2 import *

filepath = './resource/whatsapp_session.txt'
driver = webdriver

# Generales
HISTORIAL = {}
H = {}
control = True
SEDES = {"s1":"Calle 220", "s2":"CBI Soacha", "s3":"CBI Zona Franca", "s4":"Chía", "s5":"San Roque", "s6":"Suba", "s7":"Autopista Sur", "s8":"Cajicá"}

def crear_driver_session():
    with open(filepath) as fp:
        for cnt, line in enumerate(fp):
            if cnt == 0:
                executor_url = line
            if cnt == 1:
                session_id = line

    def new_command_execute(self, command, params=None):
        if command == "newSession":
            # Mock the response
            return {'success': 0, 'value': None, 'sessionId': session_id}
        else:
            return org_command_execute(self, command, params)

    org_command_execute = RemoteWebDriver.execute
    RemoteWebDriver.execute = new_command_execute
    new_driver = webdriver.Remote(command_executor=executor_url, desired_capabilities={})
    new_driver.session_id = session_id
    RemoteWebDriver.execute = org_command_execute

    return new_driver


def buscar_chats():
    #print("BUSCANDO CHATS")

    if len(driver.find_elements_by_class_name("_33LGR")) == 0:
        #print("CHAT ABIERTO")
        message = identificar_mensaje()

        if message is not None:
            return True

    chats = driver.find_elements_by_class_name("_3m_Xw")

    for chat in chats:
        #print("DETECTANDO MENSAJES SIN LEER")
        chats_mensajes = chat.find_elements_by_class_name("_23LrM")

        if len(chats_mensajes) == 0:
            #print("CHATS ATENDIDOS")
            continue

        element_name = chat.find_elements_by_class_name('zoWT4')
        name = element_name[0].text.upper().strip()

        #print("IDENTIFICANDO CONTACTO")

        with open("./resource/contactos_autorizados.txt", mode='r', encoding='utf-8') as archivo:
            contactos = [linea.rstrip() for linea in archivo]
            if name not in contactos:
                #print("CONTACTO NO AUTORIZADO : ", name)
                continue

        #print(name, "AUTORIZADO PARA SER ATENDIDO POR BOT")

        chat.click()
        return True
    return False


def normalizar(message: str):
    # -> NFD y eliminar diacríticos
    message = re.sub(
        r"([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+", r"\1",
        normalize("NFD", message), 0, re.I
    )

    # -> NFC
    return normalize('NFC', message)

def autorizacion(id):
    with open("./resource/contactos_autorizados.txt", mode='r', encoding='utf-8') as archivo:
        contactos = [linea.rstrip() for linea in archivo]
        if id not in contactos:
            return False
        else:
            return True



def eliminar_numero(numero):
    text1 = open("./resource/contactos_autorizados.txt", mode='r', encoding='utf-8').read().splitlines()
    contador = 0
    for i in text1:
        if i == numero:
            text1.pop(contador)
            f = open("./resource/contactos_autorizados.txt", mode='w')
            f.writelines("\n".join(text1))
            f.close()
        else:
            contador += 1

def indexar_phone(numero):
    aux = numero[4:]
    numero_ = ""
    for i in aux:
        if i != " ":
            numero_ += i
    return numero_


def identificar_usuario():
    # Numero Contacto
    id_usuario = driver.find_element_by_xpath('//*[@id="main"]/header/div[2]/div/div/span')
    id_usuario = id_usuario.text

    return id_usuario


def guardar_historial(message_: str):
    if identificar_usuario() not in HISTORIAL.keys():
        HISTORIAL[identificar_usuario()] = []

    #  Guardar Historial de Chats
    HISTORIAL[identificar_usuario()].append(message_)



def dict_strig(datos):
    string = ""
    for i in datos:
        string += "*"+i+"*" + ": " + datos[i][0] + ""
    return string


def identificar_mensaje():
    element_box_message = driver.find_elements_by_class_name("_22Msk")
    posicion = len(element_box_message) - 1
    posision_mensaje = element_box_message[posicion].location["x"]

    if posision_mensaje != 523:
        return None

    element_message = element_box_message[posicion].find_elements_by_class_name("_1Gy50")
    message = element_message[0].text.upper().strip()

    return normalizar(message)

def preparar_respuesta(message: str):

    #print("PREPARANDO RESPUESTA")
    #try:
    if message.__contains__("GRACIAS"):
        response = "¡Con gusto!\n"
    elif autorizacion(identificar_usuario()) == False:
        response = "En el momento el sistema indica que ya registraste una solicitud. En un momento un asesor lo atenderá \n"
        return response
    elif message.lower() in ["hola", "buenos dias", "buenas tardes"]:
        response = "Hola, buen día. \n"
    elif message == "PRESENCIAL":
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = text1.readlines()[12:13]
        text1.close()
        if identificar_usuario() not in HISTORIAL.keys():
            HISTORIAL[identificar_usuario()] = []
    elif message == "AV68":
        guardar_historial("vacio")
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = text1.readlines()[18:19]
        text1.close()
        identificar_usuario() in HISTORIAL.keys()
    elif message == "OTRAS":
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = text1.readlines()[16:17]
        text1.close()
    elif message.lower() in SEDES.keys():
        guardar_historial(SEDES[message.lower()])
        add_excel_base(indexar_phone(identificar_usuario()), HISTORIAL[identificar_usuario()][0])
        HISTORIAL.pop(identificar_usuario())
        eliminar_numero(identificar_usuario())
        response = "Sus datos han sido guardados, tan pronto haya reapertura en la dese seleccionada, nos cominicamos con usted. Muchas gracias por su atención, le deseamos un feliz día. \n"
    elif message == "DEVOLUCION":
        guardar_historial(message)
        response = "Escriba su correo electrónico, aqui se le enviará toda la información \n"
    elif message.__contains__("@") and HISTORIAL[identificar_usuario()][0] == "DEVOLUCION":
        guardar_historial(message)
        response = "¿El email *" + HISTORIAL[identificar_usuario()][1] + "* es correcto?*1.* Para Sí, *DIGITE 1**2.* Para No, *DIGITE 2* \n"
    elif "DEVOLUCION" in HISTORIAL[identificar_usuario()] and message in ["1", "2"]:
        if message == "1":
            add_excel_base(indexar_phone(identificar_usuario()), "Devolución a: "+HISTORIAL[identificar_usuario()][1])
            HISTORIAL.pop(identificar_usuario())
            eliminar_numero(identificar_usuario())
            response = "Pronto recibirá un correo electrónico donde se indica el proceso para devolución solicitada.Te deseamos un feliz día. \n"
        else:
            HISTORIAL[identificar_usuario()].pop()
            response = "Ingrese nuevamente su email \n"
    elif message.lower() in SEDES.keys():
        guardar_historial(SEDES[message.lower()])
        add_excel_base(indexar_phone(identificar_usuario()), HISTORIAL[identificar_usuario()][0])
        HISTORIAL.pop(identificar_usuario())
        eliminar_numero(identificar_usuario())
        response = "Sus datos han sido guardados, tan pronto haya reapertura en la dese seleccionada, nos cominicamos con usted. Muchas gracias por su atención, le deseamos un feliz día. \n"
    elif message.__contains__("GRACIAS"):
        response = "¡Con gusto!\n"
    elif identificar_usuario() in HISTORIAL.keys() and message in ["1", "2"] and "vacio" in HISTORIAL[identificar_usuario()]:
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = text1.readlines()[2:3]
        text1.close()
        if message == "1":
            HISTORIAL[identificar_usuario()][0] = "CC"
        else:
            HISTORIAL[identificar_usuario()][0] = "TI"
    elif message in cedulas:
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = "Por favor escriba el *género de la persona inscrita en el curso*, sin comillas: *MASCULINO* o *FEMENINO* \n"
        text1.close()
        guardar_historial(base_excel[int(message)][1])
        guardar_historial([base_excel[int(message)][0], base_excel[int(message)][2], horas_pendientes()[message][3]])
        if horas_pendientes()[message][3] == 0:
            response = "En el sistema no registran horas pendientes por reposición. \n"
            HISTORIAL.pop(identificar_usuario())
    elif message.lower() in ["femenino", "masculino"]:
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = text1.readlines()[4:5][0].format(HISTORIAL[identificar_usuario()][2][0],
                                                    HISTORIAL[identificar_usuario()][2][1],
                                                    HISTORIAL[identificar_usuario()][2][2])
        text1.close()
        guardar_historial(message)
    elif message.__contains__("@") and len(HISTORIAL[identificar_usuario()]) > 0:
        if len(HISTORIAL[identificar_usuario()]) == 4:
            HISTORIAL[identificar_usuario()].append("i")
            HISTORIAL[identificar_usuario()].append(identificar_usuario())
        HISTORIAL[identificar_usuario()][4] = message
        response = "¿El email *" + HISTORIAL[identificar_usuario()][4].lower() + "* es correcto?*1.* Para Sí, *DIGITE 1**2.* Para No, *DIGITE 2* \n"
    elif message == "1" and len(HISTORIAL[identificar_usuario()]) == 6:
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = text1.readlines()[6:7][0] + dict_strig(consulta()) + " \n"
        text1.close()
    elif message == "2" and len(HISTORIAL[identificar_usuario()]) == 6:
        try:
            HISTORIAL[identificar_usuario()][4] = " "
            response = "Escriba nuevamente su Email \n"
        except:
            response ="El dato ingresado no es válido \n"
    elif message in list(consulta().keys()):
        if len(HISTORIAL[identificar_usuario()]) == 6:
            HISTORIAL[identificar_usuario()].append(" ")
            HISTORIAL[identificar_usuario()].append([])
        if message in HISTORIAL[identificar_usuario()][7]:
            response = "Este horario ya está agendado, *seleccione uno diferente* \n"
        else:
            HISTORIAL[identificar_usuario()][7].append(message)
            response = "Elije la siguiente clase: \n"
        if len(HISTORIAL[identificar_usuario()][7]) == HISTORIAL[identificar_usuario()][2][2]:
            response = "Usted ha elegido los siquientes horarios:" + "*"+str(HISTORIAL[identificar_usuario()][7])+"*" + "Si sus horaios son correctos, escriba la palabra *GUARDAR*, de lo contrario escriba *CANCELAR* \n"
    elif message in ["1", "2"]:
        response = "El dato ingresado no es válido, por favor inténtelo nuevamentee --. \n"
    elif message.__contains__("GUARDAR"):
        response = "Ya tienes tus clases agendadas pendientes, por favor escribe la palabra *FIN*, para finalizar y agendar esta inscripción. En contados minutos recibirás un correo con la información del agendamiento y sus políticas. \n"
    elif message.__contains__("CANCELAR"):
        HISTORIAL[identificar_usuario()][7].clear()
        response = "Elije nuevamente tus horarios, recuerda escribir solo uno a la vez \n"
    elif message.__contains__("FIN"):
        HISTORIAL[identificar_usuario()][6] = "ok"
        H[identificar_usuario()] = HISTORIAL[identificar_usuario()]
        EXCEL.save_excel(H, disponibilidad(), [3,4,5,6,8,9,1])
        add_cero_to_base(H)
        enviar_mail(get_data(datos_correo(H)), H)
        HISTORIAL.pop(identificar_usuario())
        H.pop(identificar_usuario())
        eliminar_numero(identificar_usuario())
        response = "Ha sido un place ayudarte. *¡Nos vemos pronto!* \n"
    else:
        text1 = open("./resource/respuesta1.txt", mode='r', encoding='utf-8')
        response = text1.readlines()[8:9][0]
        text1.close()
    return response
    #except:
    #    response="El dato ingresado no es válido, por favor inténtelo nuevamente. \n"
    #    return response

def procesar_mensaje(message: str):
    chatbox = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[1]/div/div[2]')
    response = preparar_respuesta(message)
    chatbox.send_keys(response)

def whatsapp_boot_init():
    global driver
    driver = crear_driver_session()
    d = True

    while d:

        if not buscar_chats():
            sleep(2)
            continue
        d = False

    while True:
        sleep(1)
        message = identificar_mensaje()

        if message is None:
            buscar_chats()
            continue

        procesar_mensaje(message)


whatsapp_boot_init()


