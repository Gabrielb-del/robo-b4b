import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk  # Para a barra de progresso
import threading  # Para rodar o processamento em segundo plano

# Função para carregar a planilha
def carregar_planilha(caminho_planilha):
    df = pd.read_excel(caminho_planilha)
    if 'Status' not in df.columns:
        df['Status'] = ''  # Inicializa a coluna 'Status' se não existir
    return df

# Função para iniciar o driver do Selenium
def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conecta à porta de depuração
    driver = webdriver.Chrome(options=options)  # Inicia o driver
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    return driver

# Função para verificar se o cliente pode ser cadastrado
def verificar_cliente(driver, cnpj, contador):
    try:
        print(f"{contador} Preenchendo o CNPJ...")
        cnpj_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, '708:0')]")))
        cnpj_input.clear()
        cnpj_input.send_keys(cnpj)

        submit_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'slds-button') and contains(@class, 'uiButton--brand') and span[text()='Confirmar']]")))
        submit_button.click()

        driver.execute_script("document.querySelector('.loadingCon.global.siteforceLoadingBalls').style.display = 'none';")

        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class, 'loadingCon')]")))

        try:
            erro_telefone = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'O formato correto para o telefone é DDI + DDD + telefone')]")))
            if erro_telefone:
                return "LIVRE"
        except TimeoutException:
            pass

        try:
            erro_cadastro = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Já existe um lead cadastrado com o CNPJ informado.')]")))
            if erro_cadastro:
                return "CARIMBADO"
        except TimeoutException:
            pass

        try:
            erro_cliente = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Já existe um lead e um cliente cadastrado com o CNPJ informado.')]")))
            if erro_cliente:
                return "JÁ É CLIENTE"
        except TimeoutException:
            pass

        return "CARIMBADO"
    except NoSuchElementException as e:
        print(f"Erro ao localizar elemento: {e}")
        return "ERRO: Elemento não encontrado"
    except TimeoutException as e:
        print(f"Erro de tempo de espera: {e}")
        return "ERRO: Tempo esgotado"

# Atualizar a planilha com o status de cada cliente
def criar_nova_planilha(df, caminho_planilha):
    workbook = load_workbook(caminho_planilha)
    sheet = workbook.active
    colunas = list(df.columns)
    coluna_status = len(colunas)

    for index, status in enumerate(df['Status'], start=2):
        sheet.cell(row=index, column=coluna_status).value = status

    novo_caminho = caminho_planilha.replace('.xlsx', '_resultado.xlsx')
    workbook.save(novo_caminho)

# Função principal que será executada em uma thread separada
def processar_verificacao(caminho_planilha, progress_bar, progress_label):
    df = carregar_planilha(caminho_planilha)
    driver = iniciar_driver()
    total = len(df)
    contador = 1

    for index, row in df.iterrows():
        cnpj = row['CNPJ']
        status = verificar_cliente(driver, cnpj, contador)
        df.at[index, 'Status'] = status
        
        # Atualizar progresso
        progress = (contador / total) * 100
        progress_bar['value'] = progress
        progress_label.config(text=f"{contador}/{total} ({int(progress)}%)")
        progress_bar.update_idletasks()

        contador += 1

    criar_nova_planilha(df, caminho_planilha)
    driver.quit()  # Garante que o navegador é completamente fechado
    messagebox.showinfo("Concluído", "Processamento da planilha finalizado!")

    # Reseta a barra de progresso e o rótulo após o término
    progress_bar['value'] = 0
    progress_label.config(text="0/0 (0%)")

# Função para abrir o explorador de arquivos e iniciar o processamento
def selecionar_planilha(progress_bar, progress_label):
    caminho_planilha = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if caminho_planilha:
        # Rodar o processamento em uma nova thread
        threading.Thread(target=processar_verificacao, args=(caminho_planilha, progress_bar, progress_label)).start()

# Criar interface gráfica com Tkinter
def criar_interface():
    root = tk.Tk()
    root.title("BauBau")
    root.geometry("400x200")

    fonte_padrao = ("Calibri", 12)

    label = tk.Label(root, text="Selecione a planilha para verificar:", font=fonte_padrao)
    label.pack(pady=20)

    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)

    progress_label = tk.Label(root, text="0/0 (0%)", font=fonte_padrao)
    progress_label.pack()

    botao_selecionar = tk.Button(root, text="Selecionar Planilha", command=lambda: selecionar_planilha(progress_bar, progress_label), font=fonte_padrao)
    botao_selecionar.pack(pady=10)

    root.mainloop()

# Executa a interface
if __name__ == "__main__":
    criar_interface()
