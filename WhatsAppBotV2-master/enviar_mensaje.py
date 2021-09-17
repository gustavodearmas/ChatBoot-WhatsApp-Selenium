"""import pyautogui, webbrowser
from time import sleep
from excel_conexion import prueba

for i in prueba:
    name = prueba[i][0]
    ben = prueba[i][2]
    clase = prueba[i][3]
    phone = prueba[i][4]
    if len(str(phone)) == 10:
        mensaje = ""
        webbrowser.open(f'https://web.whatsapp.com/send?phone=+57{phone}')
        sleep(10)
        pyautogui.typewrite(mensaje)

        pyautogui.press("enter")"""

# open whatsapp web
"""for i in phone:
    webbrowser.open(f'https://web.whatsapp.com/send?phone={i}')
    sleep(10)
    pyautogui.typewrite(mensaje)
    pyautogui.press("enter")"""

from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver as RemoteWebDriver
from time import sleep, time
from excel_conexion import prueba
from whatsappBot2 import whatsapp_boot_init


filepath = './resource/whatsapp_session.txt'
driver_ = webdriver


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

driver = crear_driver_session()
text1 = open("resource/respuesta1.txt", mode='r', encoding='utf-8')
mensaje = text1.readlines()[10:11][0]
text1.close()

def send():
    for i in prueba:
        name = prueba[i][0]
        ben = prueba[i][2]
        clase = prueba[i][3]
        phone = prueba[i][4]
        driver.get(f"https://web.whatsapp.com/send?phone=+57{phone}")
        sleep(6)
        try:
            driver.find_element_by_xpath('//*[@id="app"]/div[1]/span[2]/div[1]/span/div[1]/div/div/div/div/div[2]/div').click()
        except:
            pass
        try:
            chatbox = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[1]/div/div[2]')
            response = mensaje.format(name.strip(), ben.strip(), clase) + "\n"
            chatbox.send_keys(response)
            sleep(3)
        except:
            pass
        print("MENSAJE PRINCIPAL ENVIADO ENVIADO: ", phone)
    #whatsapp_boot_init()

#//*[@id="app"]/div[1]/span[2]/div[1]/span/div[1]/div/div/div/div/div[2]/div
#_20C5O _2Zdgs
send()


"""driver2 = webdriver.Chrome(executable_path='driver/chromedriver')
for i in range(0,2):
    driver2.get("https://web.whatsapp.com/")"""

