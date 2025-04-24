import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time

# Solicitar al usuario el SSID y la contraseña
ssid = input("Ingrese el SSID: ")
password = input("Ingrese la contraseña: ")

# Determinar la ruta base
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

# Obtiene la ruta del driver en base a la ruta determinada
driver_path = os.path.join(base_path, "msedgedriver.exe")
service = Service(driver_path)
options = webdriver.EdgeOptions()
driver = webdriver.Edge(service=service, options=options)

try:
    driver.get("http://192.168.0.1")

    # Esperar a que los campos de usuario y contraseña estén presentes
    wait = WebDriverWait(driver, 20)
    username_field = wait.until(EC.presence_of_element_located((By.NAME, "username")))
    password_field = wait.until(EC.presence_of_element_located((By.NAME, "password")))

    # Ingresar las credenciales
    username_field.send_keys("user")
    password_field.send_keys("user")

    # Enviar el formulario
    password_field.send_keys(Keys.RETURN)

    # Hacer clic en el botón WAN
    wan_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@rel='4' and text()='WAN']")))
    wan_button.click()

    # Cambiar al contenido especificado
    content_iframe = wait.until(EC.presence_of_element_located((By.NAME, "contentIframe")))
    driver.switch_to.frame(content_iframe)

    # Hacer clic en la caja de verificación "Enable VLAN"
    vlan_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='vlan' and @type='checkbox']")))
    vlan_checkbox.click()

    # Ingresar el número 500 en el campo "VLAN ID"
    vlan_id_field = wait.until(EC.presence_of_element_located((By.NAME, "vid")))
    vlan_id_field.clear()
    vlan_id_field.send_keys("500")

    # Seleccionar "IPoE" en el menú desplegable "Channel Mode"
    channel_mode_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "adslConnectionMode")))
    channel_mode_dropdown.click()
    ipoe_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@name='adslConnectionMode']/option[@value='1']")))
    ipoe_option.click()

    # Seleccionar "INTERNET" en el menú desplegable "Connection Type"
    connection_type_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "ctype")))
    connection_type_dropdown.click()
    internet_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@name='ctype']/option[@value='2']")))
    internet_option.click()

    # Seleccionar "DHCP" en la sección "WAN IP Settings"
    dhcp_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='radio' and @value='1' and @name='ipMode']")))
    dhcp_option.click()

    # Hacer clic en la casilla "ALL" en la sección "Port Mapping"
    all_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='chkpt_all' and @type='checkbox']")))
    all_checkbox.click()

    # Hacer clic en el botón "Apply Changes"
    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Apply Changes' and @name='apply']")))
    apply_button.click()
    time.sleep(5)

    # Cambiar al contenido especificado
    driver.switch_to.default_content()
    wrapper = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='wrapper']")))

    # Hacer clic en el botón "WLAN"
    wlan_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='nav']/li[3]/a")))
    wlan_button.click()

    # Cambiar al contenido especificado
    content_iframe = wait.until(EC.presence_of_element_located((By.NAME, "contentIframe")))
    driver.switch_to.frame(content_iframe)

    # Ingresar el SSID en el campo correspondiente
    ssid_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text' and @name='ssid']")))
    ssid_field.clear()
    ssid_field.send_keys(ssid)

    # Seleccionar "Auto" en el menú desplegable "Channel Number"
    channel_number_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@name='chan']")))
    channel_number_dropdown.click()
    auto_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@name='chan']/option[@value='0']")))
    auto_option.click()

    # Seleccionar "100%" en el menú desplegable "Radio Power (%)"
    radio_power_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@name='txpower']")))
    radio_power_dropdown.click()
    full_power_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@name='txpower']/option[@value='0']")))
    full_power_option.click()

    # Hacer clic en el botón "Apply Changes"
    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Apply Changes' and @name='save']")))
    apply_button.click()
    time.sleep(5)

    # Cambiar al contenido especificado
    driver.switch_to.default_content()
    wrapper = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='wrapper']")))

    # Hacer clic en "Security"
    security_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@target='contentIframe' and @href='boaform/formWlanRedirect?redirect-url=/wlwpa.asp&wlan_idx=0']")))
    security_link.click()

    # Cambiar al contenido especificado
    content_iframe = wait.until(EC.presence_of_element_located((By.NAME, "contentIframe")))
    driver.switch_to.frame(content_iframe)

    # Seleccionar la opción "WPA2 Mixed" en el menú desplegable de "Encryption"
    encryption_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='security_method' and @name='security_method']")))
    encryption_dropdown.click()
    wpa2_mixed_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@id='security_method']/option[@value='6']")))
    wpa2_mixed_option.click()

    # Ingresar la contraseña en el campo "Pre-Shared Key"
    psk_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password' and @name='pskValue' and @id='wpapsk']")))
    psk_field.clear()
    psk_field.send_keys(password)

    # Hacer clic en el botón "Apply Changes"
    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Apply Changes' and @name='save']")))
    apply_button.click()
    time.sleep(5)

    # Cambiar al contenido especificado
    driver.switch_to.default_content()
    wrapper = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='wrapper']")))

    # Hacer clic en "wlan1 (2.4GHz)"
    wlan1_link = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div[1]/div[1]/div/ul/li[2]/h3/a")))
    wlan1_link.click()

    # Cambiar al contenido especificado
    content_iframe = wait.until(EC.presence_of_element_located((By.NAME, "contentIframe")))
    driver.switch_to.frame(content_iframe)

    # Ingresar el SSID en el campo correspondiente
    ssid_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text' and @name='ssid']")))
    ssid_field.clear()
    ssid_field.send_keys(ssid)

    # Seleccionar "Auto" en el menú desplegable "Channel Number"
    channel_number_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@name='chan']")))
    channel_number_dropdown.click()
    auto_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@name='chan']/option[@value='0']")))
    auto_option.click()

    # Seleccionar "100%" en el menú desplegable "Radio Power (%)"
    radio_power_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@name='txpower']")))
    radio_power_dropdown.click()
    full_power_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@name='txpower']/option[@value='0']")))
    full_power_option.click()

    # Hacer clic en el botón "Apply Changes"
    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Apply Changes' and @name='save']")))
    apply_button.click()
    time.sleep(5)

     # === Nuevo bloque: Hacer clic en "Security" para la red de 2.4GHz ===
    # Salir del iframe para interactuar con el menú lateral
    driver.switch_to.default_content()
    wrapper = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='wrapper']")))

    # Hacer clic en "Security" para la red de 2.4GHz
    security_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@target='contentIframe' and contains(@href, 'wlwpa.asp') and contains(@href, 'wlan_idx=1')]")))
    security_link.click()

    # Cambiar al contenido especificado
    content_iframe = wait.until(EC.presence_of_element_located((By.NAME, "contentIframe")))
    driver.switch_to.frame(content_iframe)

    # Seleccionar la opción "WPA2 Mixed" en el menú desplegable de "Encryption"
    encryption_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='security_method' and @name='security_method']")))
    encryption_dropdown.click()
    wpa2_mixed_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@id='security_method']/option[@value='6']")))
    wpa2_mixed_option.click()

    # Ingresar la contraseña en el campo "Pre-Shared Key"
    psk_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password' and @name='pskValue' and @id='wpapsk']")))
    psk_field.clear()
    psk_field.send_keys(password)

    # Hacer clic en el botón "Apply Changes"
    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Apply Changes' and @name='save']")))
    apply_button.click()
    time.sleep(5)

    # Cambiar al contenido especificado
    driver.switch_to.default_content()
    wrapper = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='wrapper']")))

    # Hacer clic en "Admin"
    admin_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='javascript:void(0)' and @rel='9']")))
    admin_link.click()

    # Hacer clic en el botón "Password"
    password_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@target='contentIframe' and @href='password.asp']")))
    password_link.click()

    # Cambiar al contenido especificado
    content_iframe = wait.until(EC.presence_of_element_located((By.NAME, "contentIframe")))
    driver.switch_to.frame(content_iframe)

    # Ingresar la palabra "user" en el campo "Old Password"
    old_password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password' and @name='oldpass']")))
    old_password_field.clear()
    old_password_field.send_keys("user")

    # Ingresar la nueva contraseña en el campo "New Password"
    new_password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password' and @name='newpass']")))
    new_password_field.clear()
    new_password_field.send_keys("Conectar21")

    # Ingresar la confirmación de la nueva contraseña en el campo "Confirmed Password"
    confirm_password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password' and @name='confpass']")))
    confirm_password_field.clear()
    confirm_password_field.send_keys("Conectar21")

    # Hacer clic en el botón "Apply Changes"
    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class='link_bg' and @type='submit' and @value='Apply Changes' and @name='save']")))
    apply_button.click()
    time.sleep(5)

    #Cambiar al contenido especificado y hacer clic en "Advance"
    driver.switch_to.default_content()
    wrapper = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='wrapper']")))

    # Hacer clic en "Advance"
    advance_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='javascript:void(0)' and @rel='7']")))
    advance_link.click()

    # Hacer clic en el botón "Remote Access"
    remote_access_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@target='contentIframe' and @href='rmtacc.asp']")))
    remote_access_link.click()

    # Cambiar al contenido especificado
    content_iframe = wait.until(EC.presence_of_element_located((By.NAME, "contentIframe")))
    driver.switch_to.frame(content_iframe)

    # Hacer clic en la casilla de verificación "HTTPS"
    https_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and @name='w_https' and @value='1']")))
    https_checkbox.click()

    # Hacer clic en el botón "Apply Changes"
    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class='link_bg' and @type='submit' and @value='Apply Changes' and @name='set']")))
    apply_button.click()
    time.sleep(5)

finally:

     # Mostrar mensaje de finalización mejorado con correo de contacto
    import ctypes
    ctypes.windll.user32.MessageBoxW(
        0, 
        "La configuración del módem y la red Wi-Fi se completó con éxito.\n\n"
        "Si necesitas soporte adicional, no dudes en contactarme:\n"
        "miraglioluis1@gmail.com\n\n"
        "Muchas gracias por utilizar mi Automatización.\n\n"
        "- Luis Miraglio -", 
        "Proceso de Configuración Finalizado", 
        0x40 | 0x1
    )
    time.sleep(15)
    driver.quit()
#Comando para compilar el script en un ejecutable con driver del navegador dentro de la carpeta del Script
#pyinstaller --onefile --add-binary "msedgedriver.exe;." "nombre del archivo.py"