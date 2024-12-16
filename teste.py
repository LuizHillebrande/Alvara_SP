from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import requests
import time
import pyautogui
   
def pegar_débitos_sp():
    driver = webdriver.Chrome()
    driver.get('https://senhawebsts.prefeitura.sp.gov.br/Account/Login.aspx?ReturnUrl=%2f%3fwa%3dwsignin1.0%26wtrealm%3dhttps%253a%252f%252fduc.prefeitura.sp.gov.br%252fportal%252f%26wctx%3drm%253d0%2526id%253dpassive%2526ru%253d%25252fportal%25252f%26wct%3d2024-12-16T11%253a48%253a11Z&wa=wsignin1.0&wtrealm=https%3a%2f%2fduc.prefeitura.sp.gov.br%2fportal%2f&wctx=rm%3d0%26id%3dpassive%26ru%3d%252fportal%252f&wct=2024-12-16T11%3a48%3a11Z')

    wb = openpyxl.load_workbook('dados_sp.xlsx')
    sheet_wb = wb['São Paulo']

    for indice, linha in enumerate(sheet_wb.iter_rows(min_row=2, max_row=2)):  # Ajuste o intervalo conforme necessário
        login = linha[4].value
        senha = linha[5].value

        # Preencher Login
        clicar_login = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='ctl00$ctl00$formBody$formBody$txtUser']"))
        )
        clicar_login.send_keys(str(login))

        # Preencher Senha
        clicar_senha = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='ctl00$ctl00$formBody$formBody$txtPassword']"))
        )
        clicar_senha.send_keys(str(senha))

       
        # Preencher o formulário
        campo_captcha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,"//input[@name='ctl00$ctl00$formBody$formBody$wucRecaptcha1$txtValidacao']"))
        )
        campo_captcha.click()
        campo_captcha.send_keys('1234')  
        
        print("CAPTCHA preenchido com sucesso.")
        
        # (Opcional) Submeter o formulário
        botao_submit = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='ctl00$ctl00$formBody$formBody$btnLogin']"))
        )
        botao_submit.click()
        
        print("Formulário enviado.")
        time.sleep(5)
        
        driver.quit()

# Executa a função
pegar_débitos_sp()
