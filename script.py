import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

class CrazyFormBot:
    def __init__(self, excel_path, url_form):
        self.excel_path = excel_path
        self.url_form = url_form
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 10)

    def load_data(self):
        # Cargar datos del excel
        df = pd.read_excel(self.excel_path)
        return df

    def open_form(self):
        #Abrir pagina formulario
        self.driver.get(self.url_form)
        print("Formulario abierto correctamente.")

    def start_challenge(self):
        # Iniciar reto correctamente
        try:
            start_button = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//a[contains(@class,"bg-lime-300") and contains(text(),"Iniciar Reto")]')
            ))
            start_button.click()
            print("Reto iniciado correctamente.")
            time.sleep(2)
        except TimeoutException:
            print("No se encontró el botón 'Iniciar Reto'. Verifica la página.")

    def fill_form(self, row):
        # llenar formulario con los campos del excel
        try:
            web_field = self.wait.until(EC.presence_of_element_located((By.XPATH, '//input[contains(@placeholder,"Web") or contains(@name,"web")]')))
            web_field.clear()
            web_field.send_keys(row['Web'])

            empresa_field = self.driver.find_element(By.XPATH, '//input[contains(@placeholder,"Empresa") or contains(@name,"empresa")]')
            empresa_field.clear()
            empresa_field.send_keys(row['Empresa'])

            nombre_field = self.driver.find_element(By.XPATH, '//input[contains(@placeholder,"Nombre") or contains(@name,"nombre")]')
            nombre_field.clear()
            nombre_field.send_keys(row['Nombres'])

            apellido_field = self.driver.find_element(By.XPATH, '//input[contains(@placeholder,"Apellido") or contains(@name,"apellido")]')
            apellido_field.clear()
            apellido_field.send_keys(row['Apellidos'])

            pais_field = self.driver.find_element(By.XPATH, '//input[contains(@placeholder,"País") or contains(@name,"pais")]')
            pais_field.clear()
            pais_field.send_keys(row['Pais'])

            email_field = self.driver.find_element(By.XPATH, '//input[contains(@placeholder,"Email") or contains(@name,"email")]')
            email_field.clear()
            email_field.send_keys(row['Email'])

            numero_field = self.driver.find_element(By.XPATH, '//input[contains(@placeholder,"Número") or contains(@name,"numero")]')
            numero_field.clear()
            numero_field.send_keys(str(row['Numero']))

            submit_button = self.driver.find_element(By.XPATH, '//button[contains(text(),"Enviar") or contains(@id,"submit")]')
            submit_button.click()
            print(f"Formulario para: {row['Nombres']} {row['Apellidos']}")
            time.sleep(2)

        except (TimeoutException, NoSuchElementException) as e:
            print(f"Error al llenar formualario: {e}")

    def run(self):
        data = self.load_data()
        self.open_form()
        self.start_challenge()

        for _, row in data.iterrows():
            self.fill_form(row)

        print("#### Formularios enviado correctamente ####")
        print("El navegador se cerrará automáticamente en 30 segundos")
        time.sleep(30)
        self.driver.quit()

if __name__ == "__main__":
    excel_file = "Arena RPA FormData.xlsx" #Nombre de archivo
    url = "https://arenarpa.com/crazy-form" #URL

    bot = CrazyFormBot(excel_file, url)
    bot.run()
