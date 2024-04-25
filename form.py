import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import messagebox
from time import sleep
import pyautogui
import re

rot = tk.Tk()
rot.withdraw()

df = "dados.xlsx"
tabela = pd.read_excel(df)

messagebox.showinfo("Alerta", "O programa vai começar.")
rot.destroy()


servico = Service(ChromeDriverManager().install())

navegador = webdriver.Chrome(service=servico)

def formata(numero):
    numero = numero.split('/')[0]
    numero = str(numero)
    numero_limpo = re.sub(r'\D', '', numero)

    if len(numero_limpo) == 9 or len(numero_limpo) == 8:
        ddd = "38"
        numero_tratado = ddd + numero_limpo
        return numero_tratado
    elif len(numero_limpo) >= 12:
        ddd = "38"
        numero_limpo = numero_limpo[:9]
        numero_tratado = ddd + numero_limpo
        return numero_tratado
    else:
        return numero

for i,nome in enumerate(tabela["NOME"]):
    nome = nome.strip()
    data_de_nacimento = (pd.to_datetime(tabela.loc[i,"DATA DE NASCIMENTO"], format='%d/%m/%y').date())
    data_de_nacimento = data_de_nacimento.strftime('%d/%m/%Y ')
    data_de_nacimento = re.sub(r'\D', '', data_de_nacimento)
    ano = tabela.loc[i, "ANO"]
    mae = tabela.loc[i, "MAE"]
    responsavel = "Ex N: (xx)xxxxx-xxxx"

    navegador.get("https://docs.google.com/forms/d/e/1FAIpQLSdmPUZtsX-ktyjgtMBIMI53E_-qvdTFXp0wT2qoaUKGiQW8WQ/viewform")
    sleep(3)

    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/textarea').click
    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys('Ex escola')
    sleep(1)
    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]').click()
    sleep(1)
    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(nome)

    day_picker = WebDriverWait(navegador,20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div/div[2]/div[1]/div/div[1]/input")))
    day_picker.clear()
    day_picker.send_keys(data_de_nacimento)
    pyautogui.press('enter')

    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[2]/textarea').click
    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(f"{int(ano)}° Ano")

    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div[2]/textarea').click()
    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(mae)

    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div[2]/textarea').click()
    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(responsavel)

    navegador.find_element(By.XPATH,'/html/body/div[1]/div[2]/form/div[2]/div/div[3]/div[1]/div[1]/div').click()
    sleep(2)
    navegador.quit()
    
