import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Carregar a planilha
def carregar_planilha(caminho_planilha):
    df = pd.read_excel(caminho_planilha)
    print(df.columns)  # Verifica os nomes das colunas carregadas
    if 'Status' not in df.columns:
        df['Status'] = ''  # Inicializa a coluna 'Status' se não existir
    return df


def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conecta à porta de depuração
    driver = webdriver.Chrome(options=options)  # Inicia o driver
    # Espera a página carregar
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    return driver

# Função para verificar se o cliente pode ser cadastrado
def verificar_cliente(driver, cnpj):
    try:

        print("Preenchendo o CNPJ...")
        cnpj_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '708:0')]")))
        cnpj_input.clear()
        cnpj_input.send_keys(cnpj)

        # Submeter o formulário
        submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'slds-button') and contains(@class, 'uiButton--brand') and span[text()='Confirmar']]")))
        submit_button.click()
        

        
        # Ocultar a animação de carregamento
        driver.execute_script("document.querySelector('.loadingCon.global.siteforceLoadingBalls').style.display = 'none';")


        # Verificar erros específicos e continuar assim que a mensagem aparecer
        try:
            # Aguarda até que uma das mensagens de erro seja visível
            erro_telefone = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'O formato correto para o telefone é DDI + DDD + telefone')]"))
            )
            if erro_telefone:
                return "LIVRE"
        except TimeoutException:
            pass  # Se o tempo esgotar, significa que essa mensagem não apareceu, continuar para os próximos erros

        try:
            erro_cadastro = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Já existe um lead cadastrado com o CNPJ informado.')]"))
            )
            if erro_cadastro:
                return "CARIMBADO"
        except TimeoutException:
            pass

        try:
            erro_cliente = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Já existe um lead e um cliente cadastrado com o CNPJ informado.')]"))
            )
            if erro_cliente:
                return "JÁ É CLIENTE"
        except TimeoutException:
            pass

        return "CARIMBADO"  # Se nenhuma das mensagens aparecer, retornamos "CARIMBADO"
        
    except NoSuchElementException as e:
        print(f"Erro ao localizar elemento: {e}")
        return "ERRO: Elemento não encontrado"
    except TimeoutException as e:
        print(f"Erro de tempo de espera: {e}")
        return "ERRO: Tempo esgotado"

# Atualizar a planilha com o status de cada cliente
def criar_nova_planilha(df, caminho_planilha):
    novo_caminho = caminho_planilha.replace('.xlsx', '_resultado.xlsx')  # Cria novo nome para a planilha
    df.to_excel(novo_caminho, index=False)

# Função principal
def processar_verificacao(caminho_planilha):
    df = carregar_planilha(caminho_planilha)
    driver = iniciar_driver()

    for index, row in df.iterrows():
        
        cnpj = row['CNPJ']

        status = verificar_cliente(driver, cnpj)
        df.at[index, 'Status'] = status  # Atualizar a coluna "Status" com o resultado


    criar_nova_planilha(df, caminho_planilha)  # Criar uma nova planilha com os resultados
    driver.quit()

# Exemplo de uso com a planilha Leads.xlsx:
processar_verificacao('Leads.xlsx')
