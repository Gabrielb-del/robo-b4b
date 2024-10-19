import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import tkinter as tk
from tkinter import filedialog, messagebox

# Funções existentes

def carregar_planilha(caminho_planilha):
    df = pd.read_excel(caminho_planilha)
    if 'Status' not in df.columns:
        df['Status'] = ''  # Inicializa a coluna 'Status' se não existir
    return df

def dividir_nome_completo(nome_completo):
    partes = nome_completo.split()
    nome = partes[0]
    sobrenome = " ".join(partes[1:]) if len(partes) > 1 else ""
    return nome, sobrenome

def formatar_telefone(telefone, ddi='55'):
    telefone = str(telefone)
    if not telefone.startswith('55'):
        telefone = ddi + telefone
    return telefone

def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9222")
    driver = webdriver.Chrome(options=options)
    return driver

def cadastrar_cliente(driver, nome, sobrenome, email, telefone, cnpj):
    try:
        # Esperar e preencher os campos do formulário usando XPath
        nome_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '610:0')]")))
        nome_input.send_keys(nome)
        sobrenome_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '620:0')]")))
        sobrenome_input.send_keys(sobrenome)
        email_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '655:0')]")))
        email_input.send_keys(email)
        telefone_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '671:0')]")))
        telefone_input.send_keys(telefone)
        cnpj_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '708:0')]")))
        cnpj_input.send_keys(cnpj)
        submit_button = driver.find_element(By.XPATH, "//button[contains(@class, 'slds-button') and contains(@class, 'uiButton--brand') and span[text()='Confirmar']]")
        submit_button.click()

        time.sleep(5)

        if driver.current_url != "https://c6bank.my.site.com/partners/s/createrecord/IndicacaoContaCorrente":
            return "CADASTRADO"
        else:
            return "CARIMBADO"

    except NoSuchElementException:
        return "CARIMBADO"
    except TimeoutException:
        return

def criar_nova_planilha(df, caminho_planilha):
    novo_caminho = caminho_planilha.replace('.xlsx', '_resultado.xlsx')
    df.to_excel(novo_caminho, index=False)

# Função principal adaptada para a interface
def processar_cadastro(caminho_planilha, janela):
    df = carregar_planilha(caminho_planilha)
    driver = iniciar_driver()

    for index, row in df.iterrows():
        nome, sobrenome = dividir_nome_completo(row['Nome'])
        email = row['Email']
        telefone = formatar_telefone(row['Telefone'])
        cnpj = row['CNPJ']
        status = cadastrar_cliente(driver, nome, sobrenome, email, telefone, cnpj)
        df.at[index, 'Status'] = status

        if status == "CADASTRADO":
            driver.back()
            time.sleep(5)
        else:
            driver.refresh()
            time.sleep(5)

    criar_nova_planilha(df, caminho_planilha)
    driver.quit()
    messagebox.showinfo("Processo concluído", "O cadastro foi processado e a planilha foi atualizada.")
    janela.quit()

# Interface Gráfica

def selecionar_arquivo():
    caminho_planilha = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if caminho_planilha:
        # Executa o processamento e a automação
        processar_cadastro(caminho_planilha, janela)

# Criar a janela principal
janela = tk.Tk()
janela.title("Cadastro Automático de Clientes")
janela.geometry("600x400")

# Adicionar um botão para carregar a planilha
texto = tk.Label(text = "Selecione a Planilha com os Leads")
texto.pack(padx=1, pady=80)
botao_selecionar = tk.Button(janela, text="Selecionar Planilha", command=selecionar_arquivo)
botao_selecionar.pack(pady=10)

# Iniciar o loop da interface gráfica
janela.mainloop()
