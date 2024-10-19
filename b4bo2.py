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

def dividir_nome_completo(nome_completo):
    partes = nome_completo.split()
    nome = partes[0]  # Primeiro nome
    sobrenome = " ".join(partes[1:]) if len(partes) > 1 else ""  # O resto é o sobrenome
    return nome, sobrenome


def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conecta à porta de depuração
    driver = webdriver.Chrome(options=options)  # Inicia o driver
    # Espera a página carregar
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    return driver

# Função para verificar se o cliente pode ser cadastrado
def verificar_cliente(driver, nome, sobrenome, email, telefone, cnpj):
    try:
        print("Preenchendo o nome...")
        nome_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '610:0')]")))
        nome_input.clear()
        nome_input.send_keys(nome)

        print("Preenchendo o sobrenome...")
        sobrenome_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '620:0')]")))
        sobrenome_input.clear()
        sobrenome_input.send_keys(sobrenome)

        print("Preenchendo o email...")
        email_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '655:0')]")))
        email_input.clear()
        email_input.send_keys(email)

        print("Preenchendo o telefone...")
        telefone_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '671:0')]")))
        telefone_input.clear()
        telefone_input.send_keys(telefone)

        print("Preenchendo o CNPJ...")
        cnpj_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '708:0')]")))
        cnpj_input.clear()
        cnpj_input.send_keys(cnpj)

        # Submeter o formulário
        submit_button = driver.find_element(By.XPATH, "//button[contains(@class, 'slds-button') and contains(@class, 'uiButton--brand') and span[text()='Confirmar']]")
        submit_button.click()
        
        time.sleep(5)

        # Verificar se o erro "Telefone precisa começar com 55" aparece
        try:
            erro_telefone = driver.find_element(By.XPATH, "//*[contains(text(), 'O formato correto para o telefone é DDI + DDD + telefone')]")
            if erro_telefone:
                return "LIVRE"
        except NoSuchElementException:
            pass

        # Verificar se o erro "Cliente já cadastrado" aparece
        try:
            erro_cadastro = driver.find_element(By.XPATH, "//*[contains(text(), 'Já existe um lead cadastrado com o CNPJ informado.')]")
            if erro_cadastro:
                return "CARIMBADO"
        except NoSuchElementException:
            pass

        return "Pode cadastrar"
        
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
        # Dividir nome completo
        nome_completo = row['Nome']
        nome, sobrenome = dividir_nome_completo(nome_completo)
        
        email = row['Email']
        
        # Adicionar DDI ao telefone
        telefone = row['Telefone']
        
        cnpj = row['CNPJ']

        status = verificar_cliente(driver, nome, sobrenome, email, telefone, cnpj)
        df.at[index, 'Status'] = status  # Atualizar a coluna "Status" com o resultado


    criar_nova_planilha(df, caminho_planilha)  # Criar uma nova planilha com os resultados
    driver.quit()

# Exemplo de uso com a planilha Leads.xlsx:
processar_verificacao('Leads.xlsx')
