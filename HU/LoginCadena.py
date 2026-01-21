from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from config import settings as CADENA_CONFIG


URL = CADENA_CONFIG.CADENA_CONFIG['CADENA_URL']
USUARIO = CADENA_CONFIG.CADENA_CONFIG['CADENA_USUARIO']
PASSWORD = CADENA_CONFIG.CADENA_CONFIG['CADENA_PASSWORD']


def login_colsubsidio():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    wait = WebDriverWait(driver, 25)

    driver.get(URL)
    driver.maximize_window()

    try:
        # Usuario
        user_input = wait.until(EC.presence_of_element_located((By.ID, "txtUser")))
        user_input.clear()
        user_input.send_keys(USUARIO)

        # Contraseña
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "txtPass")))
        pass_input.clear()
        pass_input.send_keys(PASSWORD)

        # Click Ingresar (es un DIV)
        btn_login = wait.until(EC.element_to_be_clickable((By.ID, "btnIngresarLogin")))
        driver.execute_script("arguments[0].click();", btn_login)

        # Esperar alguno de los resultados
        wait.until(
            lambda d: 
                "/Home/Index" in d.current_url or
                d.find_element(By.ID, "lblError").text != "" or
                d.find_element(By.ID, "dvCaptcha").is_displayed()
        )

        # Caso 1: Login exitoso
        if "/Home/Index" in driver.current_url:
            print("Login exitoso")
            return driver

        # Caso 2: Captcha
        if driver.find_element(By.ID, "dvCaptcha").is_displayed():
            raise Exception("Captcha detectado - requiere intervención")

        # Caso 3: Error credenciales
        error_msg = driver.find_element(By.ID, "lblError").text
        if error_msg:
            raise Exception(f"Error login: {error_msg}")

    except TimeoutException:
        driver.save_screenshot("timeout_login.png")
        raise Exception("Timeout durante login")

    except Exception as e:
        driver.save_screenshot("error_login.png")
        #driver.quit()
        raise e