from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import requests
import time
import pyautogui
from time import sleep
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert

def pegar_débitos_sp():


    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--ignore-certificate-errors")

    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.get('https://senhawebsts.prefeitura.sp.gov.br/Account/Login.aspx?ReturnUrl=%2f%3fwa%3dwsignin1.0%26wtrealm%3dhttps%253a%252f%252fduc.prefeitura.sp.gov.br%252fportal%252f%26wctx%3drm%253d0%2526id%253dpassive%2526ru%253d%25252fportal%25252f%26wct%3d2024-12-16T11%253a48%253a11Z&wa=wsignin1.0&wtrealm=https%3a%2f%2fduc.prefeitura.sp.gov.br%2fportal%2f&wctx=rm%3d0%26id%3dpassive%26ru%3d%252fportal%252f&wct=2024-12-16T11%3a48%253a11Z')


    wb = openpyxl.load_workbook('dados_sp.xlsx')
    sheet_wb = wb['São Paulo']

    for indice, linha in enumerate(sheet_wb.iter_rows(min_row=10, max_row=10)):  # Ajuste o intervalo conforme necessário
        login = linha[4].value
        senha = linha[5].value
        ccm = linha[3].value

        # Preencher Login
        clicar_login = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='ctl00$ctl00$formBody$formBody$txtUser']"))
        )
        clicar_login.send_keys(str(login))

        sleep(2)

        # Preencher Senha
        clicar_senha = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='ctl00$ctl00$formBody$formBody$txtPassword']"))
        )
        clicar_senha.send_keys(str(senha))

        sleep(10)
        
        
        botao_mais = WebDriverWait(driver,5).until(
            EC.element_to_be_clickable((By.XPATH,"//img[@id='img_maisTwo']"))
        )
        botao_mais.click()

        sleep(2)

        

        botao_ccm = WebDriverWait(driver,5).until(
            EC.element_to_be_clickable((By.XPATH,"//input[@name='ctl00$ctl00$ConteudoPrincipal$ContentPlaceHolder1$txtCCM']"))
        )
        botao_ccm.click()
        botao_ccm.send_keys(ccm)

        sleep(2)

        botao_OK = WebDriverWait(driver,5).until(
            EC.element_to_be_clickable((By.XPATH,"//input[@name='ctl00$ctl00$ConteudoPrincipal$ContentPlaceHolder1$btnOKDAMSP']"))
        )

        botao_OK.click()

        time.sleep(15)

        #print(driver.page_source)

        driver.switch_to.default_content()

        try:
            wait = WebDriverWait(driver, 10)

            # Define as mensagens esperadas (certifique-se de que estão corretas)
            mensagens_esperadas = [
                "Não há informações registradas para sugestões de recolhimentos de ISS.",
                "Não há informações registradas para sugestões de recolhimentos de TFE.",
                "Não há informações registradas para sugestões de recolhimentos de TFA.",
                "Não há informações registradas para sugestões de recolhimentos de TRSS."
            ]

            # Lista para armazenar os resultados da verificação
            mensagens_encontradas = []

            # Procura pelos elementos no DOM
            for mensagem in mensagens_esperadas:
                try:
                    # Localiza o elemento `<span>` contendo o texto (com normalize-space)
                    elemento = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, f"//span[@class='rotuloCampo' and contains(normalize-space(), '{mensagem}')]"))
                    )
                    print(f"Encontrado: {elemento.text}")
                    mensagens_encontradas.append(True)
                except Exception:
                    print(f"Não encontrado: {mensagem}")
                    mensagens_encontradas.append(False)

            # Verificação final
            if all(mensagens_encontradas):
                print("A empresa NÃO possui débitos.")
            else:
                print("A empresa POSSUI débitos.")
                
                pagar = WebDriverWait(driver,5).until(
                    EC.element_to_be_clickable((By.XPATH,"//input[@name='ctl00$ctl00$ConteudoPrincipal$ContentPlaceHolder1$btnPagar']"))
                )

                pagar.click()
                try:
                    alert = Alert(driver)
                    alert.accept()  # Fecha o pop-up clicando em "OK"

                except:
                    print("Nenhum alerta foi encontrado.")

                    


                
        except Exception as e:
            print(f"Ocorreu um erro: {e}")

        time.sleep(1)
        
        
        driver.quit()

# Executa a função
pegar_débitos_sp()
