from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import requests
import time
import pyautogui
import glob

API_KEY = "62a06978600e98d3cbf430ffe18ea254"

def capturar_regiao_captcha(x1, y1, x2, y2, output_path):
    try:
        largura = x2 - x1  # Calcula a largura do retângulo
        altura = y2 - y1   # Calcula a altura do retângulo
        print(f"Capturando região com largura={largura} e altura={altura}")
        
        # Captura apenas a região especificada
        screenshot = pyautogui.screenshot(region=(x1, y1, largura, altura))
        screenshot.save(output_path)
        print(f"CAPTCHA salvo em: {output_path}")
        return output_path
    except Exception as e:
        print(f"Erro ao capturar a região do CAPTCHA: {e}")
        return None
    
def resolver_captcha_2captcha(image_path):
    try:
        # Passo 1: Enviar a imagem do CAPTCHA
        print("Enviando CAPTCHA para o 2Captcha...")
        with open(image_path, 'rb') as img_file:
            files = {'file': img_file}
            response = requests.post(
                f"http://2captcha.com/in.php?key={API_KEY}&method=post",
                files=files
            )
        
        if response.status_code != 200 or "OK" not in response.text:
            raise Exception("Erro ao enviar o CAPTCHA para o 2Captcha")
        
        captcha_id = response.text.split('|')[1]
        print(f"CAPTCHA enviado. ID: {captcha_id}")
        
        # Passo 2: Aguardar e consultar o resultado
        for _ in range(30):  # Tenta por até 30 segundos
            time.sleep(5)  # Espera 5 segundos entre consultas
            result_response = requests.get(f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}")
            if result_response.text == "CAPCHA_NOT_READY":
                print("CAPTCHA ainda não está pronto, aguardando...")
                continue
            
            if "OK" in result_response.text:
                captcha_text = result_response.text.split('|')[1]
                print(f"CAPTCHA resolvido: {captcha_text}")
                return captcha_text
            
        raise Exception("Tempo esgotado para resolver o CAPTCHA")
    except Exception as e:
        print(f"Erro ao resolver CAPTCHA com 2Captcha: {e}")
        return None
    
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

        try:
            # Captura a região do CAPTCHA
            captcha_path = "captcha_regiao.png"
            capturar_regiao_captcha(428, 455, 780, 600, captcha_path)  # regiao exata
            
            # Resolver o CAPTCHA via 2Captcha
            captcha_resposta = resolver_captcha_2captcha(captcha_path)
            print(captcha_resposta)
            if not captcha_resposta:
                print("Falha ao resolver o CAPTCHA.")
                return
            
            # Preencher o formulário
            campo_captcha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,"//input[@name='ctl00$ctl00$formBody$formBody$wucRecaptcha1$txtValidacao']"))
            )
            campo_captcha.click()
            campo_captcha.send_keys(captcha_resposta)  
            
            print("CAPTCHA preenchido com sucesso.")
            
            botao_submit = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@name='ctl00$ctl00$formBody$formBody$btnLogin']"))
            )
            botao_submit.click()
            
            print("Formulário enviado.")
            time.sleep(15)
        finally:
            driver.quit()
    

# Executa a função
pegar_débitos_sp()
