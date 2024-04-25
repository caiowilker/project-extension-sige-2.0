#Importaçoes de bibliotecas necessarias
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from time import sleep
import pyautogui
import re
from selenium.common.exceptions import NoSuchElementException

#Abre tela de invisível
rot = tk.Tk()
rot.withdraw()
#Exibe a tela, alertando o início do programa
messagebox.showinfo("Alerta", "O programa vai começar.")
rot.destroy()

#Instala chromedriver compatível com o navegador
servico = Service(ChromeDriverManager().install())
#Abre a tela do navegador controlada pelo programa
navegador = webdriver.Chrome(service=servico)
#Recebe a planilha contendo os dados a serem processados
df = "dados.xlsx"
tabela = pd.read_excel(df)
#Função que formata os números recebidos
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


#Função que corrige o erro de sobreposição impedindo o click
def clicar_elemento_por_id(xpath_elemento):

    try:
        element_locator = (By.XPATH,xpath_elemento)
        element = navegador.find_element(*element_locator)

        actions = ActionChains(navegador)
        actions.move_to_element(element).click().perform()


        print(f"Elemento com ID '{xpath_elemento}' clicado com sucesso!")

    except Exception as e:
        print(f"Ocorreu um erro ao clicar no elemento: {e}")


#Abre o site de cadastro dos dados
navegador.get("https://saofrancisco-mg.nobesistemas.com.br/educacao/users/sign_in")
sleep(3)
navegador.find_element(By.XPATH,'//*[@id="user_login"]').send_keys('usuário')
navegador.find_element(By.XPATH,'//*[@id="user_password"]').send_keys('senha')
navegador.find_element(By.XPATH,'//*[@id="user_submit"]').click()
sleep(5)
#Laço que percorre os dados individualmente de cada aluno e executa o cadastramento ou renovação de matrícula
for i,nome in enumerate(tabela["NOME"]):
    navegador.find_element(By.XPATH, '//*[@id="menu"]/ul/li[6]/a').click()
    sleep(3)
    navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/ul/li[6]/ul/li[1]/a').click()
    sleep(3)
    nome = nome.strip()
    data_de_nacimento = (pd.to_datetime(tabela.loc[i,"DATA DE NASCIMENTO"], format='%d/%m/%y').date())
    data_de_nacimento = data_de_nacimento.strftime('%d/%m/%Y ')
    telefone = formata(str(tabela.loc[i,"TELEFONE"]))
    turma = tabela.loc[i,"TURMA"]
    #Turmas a serem percorridas individualmente para o cadastramento
    match turma:
        case"A BELA E A FERA":
            turma = 2
        case"ALEGRIA":
            turma = 3
        case"AMIZADE":
            turma = 4
        case"AMOR":
            turma = 5
        case"A PEQUENA SEREIA":
            turma = 6
        case"AROEIRA":
            turma = 7
        case"BELEZA":
            turma = 8
        case"BONDADE":
            turma = 9
        case"BRANCA DE NEVE":
            turma = 10
        case"BRAVURA":
            turma = 11
        case"CACHINHOS DOURADOS":
            turma = 12
        case"CARIDADE":
            turma  = 13
        case"CAVALEIROS DO ZODIACO":
            turma = 14
        case"CHAPEUZINHO VERMELHO":
            turma = 15
        case"CINDERELA":
            turma = 16
        case"CONFIANÇA":
            turma = 17
        case"CORAGEM":
            turma = 18
        case"DRAGON BALL":
            turma = 19
        case"ENTUSIASMO":
            turma = 20
        case"FE":
            turma = 21
        case"FELICIDADE":
            turma = 22
        case"FORÇA":
            turma = 23
        case"FORTALEZA":
            turma = 24
        case"GENTILEZA":
            turma = 25
        case"GUARDIOES DA GALAXIA":
            turma = 26
        case"HARMONIA":
            turma = 27
        case"HONESTIDADE":
            turma = 28
        case"HUMILDADE":
            turma = 29
        case"IGUALDADE":
            turma = 30
        case"JACARE":
            turma = 31
        case"JOAO E MARIA":
            turma = 32
        case"JOVENS TITAS":
            turma = 33
        case"JUSTICA JOVEM":
            turma = 34
        case"LEALDADE":
            turma = 35
        case"LIBERDADE":
            turma = 36
        case"LIGA DA JUSTIÇA":
            turma = 37
        case"LIGA EXTRAORDINÁRA":
            turma = 38
        case"MARGARIDA":
            turma = 39
        case"O MAGICO DE OZ":
            turma = 40
        case"OS DEFENSORES ":
            turma = 41
        case"OS INCRIVEIS":
            turma = 42
        case"OS TRES PORQUINHOS":
            turma = 43
        case"OS VINGADORES":
            turma = 44
        case"OUSADIA":
            turma = 45
        case"PAZ":
            turma = 46
        case"PEIXE":
            turma = 47
        case"PEIXE":
            turma = 48
        case"POWER RANGERS":
            turma = 49
        case"QUARTETO FANTASTICO":
            turma = 50
        case"SABEDORIA":
            turma = 51
        case"SIMPATIA":
            turma = 52
        case"SOLDADINHO DE CHUMBO":
            turma  = 53
        case"SOLIDARIEDADE":
            turma  = 54
        case"SUPER AMIGOS":
            turma = 55
        case"TARTARUGAS NINJAS":
            turma = 56
        case"UNIAO":
            turma = 57
        case"URSINHO PUFF":
            turma = 58
        case"URSINHOS CARINHOSOS":
            turma = 59
        case"X- FORCE":
            turma = 60
        case"X-MEN":
            turma = 61
        case _:
            turma = 1


    sexo = tabela.loc[i,"SX"]

    navegador.find_element(By.XPATH, '//*[@id="content"]/div[3]/ul/li[3]/a').click()
    pyautogui.press('backspace')
    navegador.find_element(By.XPATH, '//*[@id="enrollment_date"]').send_keys('08/02/2024')
    sleep(3)
    navegador.find_element(By.XPATH, '//*[@id="s2id_enrollment_person"]/a').click()
    navegador.find_element(By.XPATH, '//*[@id="select2-drop"]/div/input').send_keys(nome)
    sleep(5)

    #Função que verifica se o elemento está visível para efetuar o click
    def is_element_dysplayed(xpath):

        try:
            element = navegador.find_element(By.XPATH,xpath)
            return element.is_displayed()
        except NoSuchElementException:
            return False
    xpath_element ='//*[@id="select2-drop"]/ul/li/a'

    #Condição que verifica se efetuará o cadastramento ou a renovação
    if is_element_dysplayed(xpath_element):
        #Área de cadastro de novas pessoas
        navegador.get(f'https://saofrancisco-mg.nobesistemas.com.br/educacao/people/new?term={nome}')
        dados = pd.DataFrame()
        dados = dados._append({'Aluno':nome}, ignore_index=True)
        sleep(5)
        navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_extended_individual_attributes_marital_status"]/option[6]').click()
        sleep(2)
        navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_birthdate"]').click()
        navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_birthdate"]').send_keys(data_de_nacimento)
        if sexo == 'F':
            navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_gender_female"]').click()
        else:
            navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_gender_male"]').click()
        sleep(2)

        navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_mother"]').send_keys('xxxxx')
        navegador.find_element(By.XPATH,'//*[@id="person_address_attributes_zip_code"]').click()
        navegador.find_element(By.XPATH,'//*[@id="person_address_attributes_zip_code"]').send_keys('Cep')
        navegador.find_element(By.XPATH,'//*[@id="s2id_person_address_attributes_street"]/a').click()
        navegador.find_element(By.XPATH,'//*[@id="select2-drop"]/div/input').send_keys('Endereço')
        sleep(3)
        pyautogui.press('enter')
        navegador.find_element(By.XPATH,'//*[@id="s2id_person_address_attributes_neighborhood"]/a').click()
        navegador.find_element(By.XPATH,'//*[@id="select2-drop"]/div/input').send_keys('Bairro')
        sleep(3)
        pyautogui.press('enter')
        navegador.find_element(By.XPATH,'//*[@id="s2id_person_address_attributes_district"]/a').click()
        navegador.find_element(By.XPATH,'//*[@id="select2-drop"]/div/input').send_keys('Município')
        sleep(3)
        pyautogui.press('down')
        pyautogui.press('enter')
        navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_extended_individual_attributes_differentiated_treatment"]').click()
        navegador.find_element(By.XPATH,'//*[@id="person_personable_attributes_extended_individual_attributes_differentiated_treatment"]/option[4]').click()
        navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[4]/input').click()
        sleep(5)

        #Função que verifica se o aluno foi cadastrado com sucesso ou se já existia o cadastro 
        xpath_erro ='/html/body/div[1]/div[3]/form/div[1]/div/ul/li'
        if is_element_dysplayed(xpath_erro):
            navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[3]/div[2]/a')
        else:
            pass
        
        #Área de renovação de matrícula 
        navegador.find_element(By.XPATH, '//*[@id="menu"]/ul/li[6]/a').click()
        sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="menu"]/ul/li[6]/ul/li[1]/a').click()
        sleep(3)
        navegador.find_element(By.XPATH, '//*[@id="content"]/div[3]/ul/li[3]/a').click()
        sleep(3)
        pyautogui.press('backspace')
        navegador.find_element(By.XPATH, '//*[@id="enrollment_date"]').send_keys('08/02/2024')
        navegador.find_element(By.XPATH, '//*[@id="enrollment_enrollment_type"]/option[7]').click()
        sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="s2id_enrollment_person"]/a').click()
        navegador.find_element(By.XPATH, '//*[@id="select2-drop"]/div/input').send_keys(nome)
        sleep(4)
        pyautogui.press('enter')
        navegador.find_element(By.XPATH,f'// *[ @ id = "enrollment_enrollment_academic_years_attributes_0_academic_classroom_id"] / option[{str(turma)}]').click()
        sleep(5)
        nvt = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[3]/div[1]/div/label/input"
        clicar_elemento_por_id(nvt)
        sleep(3)
        #navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[3]/div[1]/div/label/input').click()
        ensino = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[4]/div[1]/fieldset/div[1]/label/input"
        clicar_elemento_por_id(ensino)
        atividades = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[4]/div[1]/fieldset/div[2]/label/input"
        clicar_elemento_por_id(atividades)
        celular = "/html/body/div[1]/div[3]/form/div[1]/div[1]/fieldset[1]/div[2]/div[1]/div/span[1]/label/input"
        clicar_elemento_por_id(celular)
        tef = "/html/body/div[1]/div[3]/form/div[1]/div[1]/fieldset[1]/div[2]/div[2]/div/input"
        clicar_elemento_por_id(tef)
        navegador.find_element(By.XPATH,'//*[@id="enrollment_responsible_phone"]').send_keys(telefone)
        navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[2]/div[2]/input').click()
    else:
        #Caso o aluno já estiver cadastro é direcionado direto para a renovação de matrícula, sendo esta área 
        pyautogui.press('enter')
        navegador.find_element(By.XPATH,'//*[@id="enrollment_enrollment_type"]/option[10]').click()
        sleep(2)
        navegador.find_element(By.XPATH,f'// *[ @ id = "enrollment_enrollment_academic_years_attributes_0_academic_classroom_id"] / option[{str(turma)}]').click()
        sleep(2)
        nvt = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[3]/div[1]/div/label/input"
        clicar_elemento_por_id(nvt)
        sleep(3)
        #navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[3]/div[1]/div/label/input').click()
        ensino = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[4]/div[1]/fieldset/div[1]/label/input"
        clicar_elemento_por_id(ensino)
        atividades = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[4]/div[1]/fieldset/div[2]/label/input"
        clicar_elemento_por_id(atividades)
        celular = "/html/body/div[1]/div[3]/form/div[1]/div[1]/fieldset[1]/div[2]/div[1]/div/span[1]/label/input"
        clicar_elemento_por_id(celular)
        tef = "/html/body/div[1]/div[3]/form/div[1]/div[1]/fieldset[1]/div[2]/div[2]/div/input"
        clicar_elemento_por_id(tef)
        navegador.find_element(By.XPATH,'//*[@id="enrollment_responsible_phone"]').send_keys(telefone)
        navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[2]/div[2]/input').click()

#Abre uma tela invisível 
root = tk.Tk()
root.withdraw()

#Pede ao usuário um caminho para salvar os alunos novatos 
caminho_pasta = filedialog.askdirectory()

#Cria uma planilha no caminho informado contendo os alunos cadastrados
if caminho_pasta:
    caminho_arquivo = caminho_pasta + "/alunos.xlsx"

    dados.to_excel(caminho_arquivo, index=False)
    print(f"A planilha foi salva em '{caminho_arquivo}'.")
else:
    print("Nenhum diretório foi selecionado. A planilha não foi salva.")

#limpa os dados 
dados.drop(index=dados ,columns=dados.columns, inplace=True)
