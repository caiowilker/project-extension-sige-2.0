import time
import logging
import re
import random

import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


class SigeAutomator:
    def __init__(self, excel_path, navegador='chrome'):
        self.excel_path = excel_path
        self.navegador_nome = navegador
        self.driver = None
        self.alunos = []
        self.resultados = []
        self.logger = self._setup_logger()

    def _setup_logger(self):
        logging.basicConfig(
            filename='sige_automacao.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        return logging.getLogger()

    def iniciar_navegador(self):
        try:
            if self.navegador_nome.lower() == 'chrome':
                self.driver = webdriver.Chrome()
            elif self.navegador_nome.lower() == 'edge':
                self.driver = webdriver.Edge()
            else:
                raise ValueError(f"Navegador '{self.navegador_nome}' não suportado.")
            self.driver.maximize_window()
            self.logger.info(f'Navegador {self.navegador_nome} iniciado.')
        except Exception as e:
            self.logger.error(f'Erro ao iniciar navegador: {e}')
            raise

    def esperar_elemento(self, by, valor, timeout=15, condicao=EC.presence_of_element_located):
        try:
            return WebDriverWait(self.driver, timeout).until(condicao((by, valor)))
        except TimeoutException:
            raise TimeoutException(f'Elemento {valor} não encontrado dentro de {timeout}s.')

    def clicar_elemento(self, by, valor, timeout=15):
        elemento = self.esperar_elemento(by, valor, timeout, EC.element_to_be_clickable)
        elemento.click()

    def preencher_campo(self, by, valor, texto, timeout=15):
        campo = self.esperar_elemento(by, valor, timeout, EC.visibility_of_element_located)
        campo.clear()
        campo.send_keys(texto)

    def marcar_checkbox(self, for_id):
        try:
            label = self.driver.find_element(By.CSS_SELECTOR, f"label[for='{for_id}']")
            if label:
                label.click()
        except NoSuchElementException:
            self.logger.warning(f'Checkbox {for_id} não encontrado.')

    def _formatar_telefone(self, telefone_bruto):

        telefone_bruto = telefone_bruto.split('/')[0]
        telefone_bruto = telefone_bruto.replace(' ', '').replace('-', '').replace('(', '').replace(')', '')

        match = re.search(r'\d{8,11}', telefone_bruto)
        if match:
            numero = match.group(0)
            if len(numero) == 8 or len(numero) == 9:
                numero = '38' + numero
            return numero
        else:
            return None

    def preencher_campo_mascarado(self, by, valor, texto, timeout=15):
        campo = self.esperar_elemento(by, valor, timeout, EC.visibility_of_element_located)
        campo.clear()
        campo.click()
        for caractere in str(texto):
            campo.send_keys(caractere)
            time.sleep(random.uniform(0.03, 0.08))

    def fazer_login(self, usuario, senha):
        self.driver.get("https://saofrancisco-mg.nobesistemas.com.br/educacao/users/sign_in")
        self.preencher_campo(By.ID, 'user_login', usuario)
        self.preencher_campo(By.ID, 'user_password', senha)
        self.clicar_elemento(By.ID, 'user_submit')
        self.logger.info('Login realizado com sucesso.')

    def carregar_dados(self):
        try:
            self.alunos = pd.read_excel(self.excel_path).to_dict(orient='records')
            self.logger.info(f'{len(self.alunos)} alunos carregados.')
        except Exception as e:
            self.logger.error(f'Erro ao carregar Excel: {e}')
            raise

    def navegar_para_formulario(self):
        self.clicar_elemento(By.XPATH, '/html/body/div[1]/div[1]/div[2]/ul/li[6]/a')
        self.clicar_elemento(By.XPATH, '/html/body/div[1]/div[1]/div[2]/ul/li[6]/ul/li[1]/a')
        self.clicar_elemento(By.XPATH, '/html/body/div[1]/div[3]/div[4]/ul[2]/li[3]/a')

    def selecionar_aluno(self, nome):
        self.clicar_elemento(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/div[1]/div/div[2]/div/a')
        campo_nome = self.esperar_elemento(By.XPATH, '/html/body/div[5]/div/input', condicao=EC.visibility_of_element_located)
        campo_nome.send_keys(nome)
        WebDriverWait(self.driver, 5).until(lambda driver: campo_nome.get_attribute('value') == nome)
        time.sleep(4)
        pyautogui.press('enter')
        self.logger.info(f'Aluno {nome} selecionado.')

    def preencher_formulario(self, aluno):
        self.navegar_para_formulario()
        self.selecionar_aluno(aluno['Nome'])

        # Tipo de matrícula
        select_matricula = Select(self.esperar_elemento(By.ID, 'enrollment_enrollment_type'))
        select_matricula.select_by_value('renewal')

        # Turma
        select_turma = Select(self.esperar_elemento(
            By.ID, 'enrollment_enrollment_academic_years_attributes_0_academic_classroom_id'))
        select_turma.select_by_visible_text(aluno['Turma'])

        # Marcar checkboxes
        opcoes = [
            'enrollment_literate',
            'enrollment_religious_education',
            'enrollment_activities_proposed_by_school',
            'enrollment_studied_last_year'
        ]
        for opcao in opcoes:
            self.marcar_checkbox(opcao)

        # Preencher responsável e telefone
        self.preencher_campo(By.ID, 'enrollment_responsible_name', aluno['Responsavel'])
        self.marcar_checkbox('enrollment_phone_type_mobile')

        telefone_formatado = self._formatar_telefone(aluno['Fone'])
        if telefone_formatado is None:
            raise ValueError(f"Telefone inválido para o aluno {aluno.get('Nome')}: {aluno.get('Fone')}")

        self.preencher_campo_mascarado(By.ID, 'enrollment_responsible_phone', telefone_formatado)

        time.sleep(30)
        self.logger.info(f'Dados do aluno {aluno["Nome"]} preenchidos com sucesso.')

        # Salvar matrícula
        self.clicar_elemento(By.ID, 'enrollment_submit')

    def processar_alunos(self):
        for aluno in self.alunos:
            inicio = time.time()
            try:
                self.preencher_formulario(aluno)
                status = 'Sucesso'
                mensagem = ''
            except Exception as e:
                status = 'Erro'
                mensagem = str(e)
                self.logger.error(f'Erro ao processar {aluno.get("Nome", "Sem Nome")}: {mensagem}')
            tempo_exec = round(time.time() - inicio, 2)
            self.resultados.append({
                'Nome': aluno.get('Nome'),
                'Status': status,
                'Mensagem': mensagem,
                'Tempo de Execução (s)': tempo_exec
            })

    def gerar_relatorio(self, caminho_saida='relatorio_final.xlsx'):
        df = pd.DataFrame(self.resultados)
        df.to_excel(caminho_saida, index=False)
        self.logger.info(f'Relatório final salvo em {caminho_saida}')

    def executar(self, usuario, senha):
        try:
            self.carregar_dados()
            self.iniciar_navegador()
            self.fazer_login(usuario, senha)
            self.processar_alunos()
        finally:
            if self.driver:
                self.driver.quit()
                self.logger.info('Navegador fechado.')
            self.gerar_relatorio()
