import pandas as pd
import logging
from funcoes import abrir_navegador, efetuar_login, extraindo_dados_tabela, realizar_cancelamento

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# Nome do arquivo de dados
nome_arquivo = 'cpf_codigo.xlsx'

def main():
    df = pd.read_excel(nome_arquivo)
    driver = abrir_navegador()
    
    for index, row in df.iterrows():
        login, senha = str(row['login']), str(row['senha'])
        if efetuar_login(driver, login, senha):
            dados_tabela = extraindo_dados_tabela(driver)
            codigos = [dado['CODIGO'] for dado in dados_tabela]  # Extrai os códigos da tabela
            realizar_cancelamento(driver, codigos)
    
    logger.info("Todos os logins e senhas foram tentados")
    logger.info("Encerrando o algoritmo...")

    driver.quit()

if __name__ == "__main__":
    main()