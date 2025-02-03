from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import pyautogui
import logging
from dotenv import load_dotenv
import os

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

load_dotenv()
url = os.getenv("URL")

# Nome do arquivo de dados
nome_arquivo = 'cpf_codigo.xlsx'

def efetuar_login(driver, login, senha):
    try:
        driver.get(url)

        elemento_login = driver.find_element(By.XPATH, '//*[@id="pCodigoEmpresa"]')
        elemento_login.send_keys(login)
        logger.info(f"Campo de login preenchido: {login}")

        elemento_senha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="pSenha"]'))
        )
        elemento_senha.click()
        time.sleep(1)
        pyautogui.typewrite(senha)
        logger.info(f"Campo de senha preenchido: {senha}")
        time.sleep(1)

        driver.find_element(By.XPATH, '//*[@id="div1"]/div/input').click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="div2"]/h3'))
        )
        logger.info("Login realizado com sucesso!")
        return True
    except:
        logger.error(f"Erro ao efetuar login para: {login}" )
        driver.get(url)
        return False
        

 

def extraindo_dados_tabela(driver):
    logger.info("Buscando usuários ativos...")
    driver.execute_script("window.scrollBy(0, 200);")

    data = []  # Lista para armazenar os dados da tabela

    try:
        # Clica no link que leva à tabela
        elemento_link = driver.find_element(By.XPATH, '//*[@id="div2"]/table[2]/tbody/tr[5]/td[1]/strong/a[3]')
        elemento_link.click()

        # Aguarda o carregamento da tabela
        tabela_xpath = '//*[@id="div2"]/table/tbody/tr/td[2]/table[2]/tbody/tr[3]/td/table/tbody'
        linhas = driver.find_elements(By.XPATH, tabela_xpath + '/tr')  # Pega todas as linhas

        # Itera sobre as linhas, capturando CPF, código, titular_dependente e plano
        for i, linha in enumerate(linhas[2:], start=1):  # Ignora as duas primeiras linhas (cabeçalho)

            if len(data) >= 5:  # Interrompe após capturar 5 registros
                break

            try:
                element_titular_dependente = linha.find_elements(By.XPATH, './td[1]')  # Coluna titular_dependente
                element_cpf = linha.find_elements(By.XPATH, './td[2]')  # Coluna CPF
                element_codigo = linha.find_elements(By.XPATH, './td[3]')  # Coluna Código
                element_plano = linha.find_elements(By.XPATH, './td[4]')  # Coluna plano

                # Verifica se os elementos existem na linha
                if not element_titular_dependente or not element_cpf or not element_codigo or not element_plano:
                    logger.info(f"Linha {i} ignorada: Um ou mais dados não encontrados.")
                    continue  # Ignora a linha caso não tenha os dados necessários

                # Extrai os valores
                titular_dependente = element_titular_dependente[0].text if element_titular_dependente else ''
                cpf = element_cpf[0].text if element_cpf else ''
                codigo = element_codigo[0].text if element_codigo else ''
                plano = element_plano[0].text if element_plano else ''

                dados_linha = {
                    "TITULAR / DEPENDENTE": titular_dependente,
                    "CPF": cpf,
                    "CODIGO": codigo,
                    "PLANO": plano,
                }
                data.append(dados_linha)

                logger.info(f"dados_linha: {dados_linha}.")

            except Exception as e:
                logger.error(f"Erro ao processar linha {i}: {e}")

        # Salva os dados em um arquivo Excel
        df = pd.DataFrame(data)
        df.to_excel("credenciais_odonto.xlsx", index=False)
        logger.info("Dados salvos no arquivo 'credenciais_odonto.xlsx'.")

    except Exception as e:
        logger.error(f"Erro ao captar informações da tabela: {e}")

    driver.back()
    time.sleep(2)
    logger.info("Retornou para a página anterior após buscar usuários.")
    return data

def realizar_cancelamento(driver, codigos):
    try:
        # Clica no link que leva à página de cancelamento
        elemento_cancelamento = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="div2"]/table[2]/tbody/tr[2]/td[1]/strong/a[1]'))
        )
        elemento_cancelamento.click()
        logger.info("Navegando para a página de cancelamento...")

        # Itera sobre os códigos e realiza o cancelamento
        for codigo in codigos:
            try:
                # Localiza o campo de entrada e insere o código
                campo_usuario = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="pUsuario"]'))
                )
                campo_usuario.clear()  # Limpa o campo antes de inserir o próximo código
                campo_usuario.send_keys(codigo)
                logger.info(f"Código: {codigo}, inserido para cancelamento.")

                # Aqui você pode adicionar a lógica para confirmar o cancelamento, se necessário
                # Exemplo: clicar em um botão de confirmação

                # Limpa o campo para o próximo código
                time.sleep(1)
                campo_usuario.clear()
                # Espera um pouco antes de processar o próximo código

            except Exception as e:
                logger.error(f"Erro ao processar código {codigo}: {e}")

        logger.info("Todos os códigos foram processados para cancelamento.")

    except Exception as e:
        logger.error(f"Erro ao acessar a página de cancelamento: {e}")

def main():
    df = pd.read_excel(nome_arquivo)
    driver = webdriver.Chrome()
    
    for index, row in df.iterrows():
        login, senha = str(row['login']), str(row['senha'])
        if efetuar_login(driver, login, senha):
            dados_tabela = extraindo_dados_tabela(driver)
            codigos = [dado['CODIGO'] for dado in dados_tabela]  # Extrai os códigos da tabela
            realizar_cancelamento(driver, codigos)
    
    # Aqui, após o loop de logins
    logger.info("Todos os logins e senhas foram tentados")
    logger.info("Encerrando o algoritmo...")

    driver.quit()

if __name__ == "__main__":
    main()