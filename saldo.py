import requests

API_KEY = '62a06978600e98d3cbf430ffe18ea254'  # Substitua com sua chave de API

def verificar_saldo(api_key):
    url = 'http://2captcha.com/res.php'
    params = {
        'key': api_key,
        'action': 'getbalance',  # Ação para obter o saldo
        'json': 1  # Resposta em JSON
    }
    response = requests.get(url, params=params)
    result = response.json()

    if result['status'] == 1:
        return result['request']  # O saldo será retornado aqui
    else:
        raise Exception("Erro ao obter o saldo:", result)

try:
    saldo = verificar_saldo(API_KEY)
    print(f"Seu saldo é: {saldo} créditos")
except Exception as e:
    print(f"Erro: {e}")
