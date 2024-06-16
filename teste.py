from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import sys
import json
from glob import iglob
from os.path import getmtime
from pathlib import Path
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class INEP_BOT:
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        self.url = "http://sistemasenem.inep.gov.br/EnemSolicitacao/login.seam"
        self.driver.get(self.url)
        
        if not self.wait_for_element(By.ID, 'Wrapper'):
            print("Erro ao carregar a página de login.")
            return

        self.config = self.load_json_file("bot_config")
        self.db = self.load_json_file("db")
        self.verif_all_files()

    def wait_for_element(self, by, value, timeout=10):
        for _ in range(timeout):
            if self.driver.find_elements(by, value):
                return True
            time.sleep(1)
        return False

    def load_json_file(self, filename):
        path = os.getcwd()
        with open(os.path.join(path, filename), "r", encoding='utf-8') as file:
            return json.load(file)

    def save_json_file(self, data, filename):
        path = os.getcwd()
        with open(os.path.join(path, filename), "w", encoding='utf-8') as file:
            json.dump(data, file)

    def verif_all_files(self):
        remessa = self.config['caminho_pasta_remessa']
        arquivos = iglob(os.path.join(remessa, "*.txt"))
        files = sorted(arquivos, key=getmtime, reverse=True)

        for f in files:
            self.arquivo_recente = f
            filename = os.path.basename(f)
            if filename not in self.db:
                self.ano = filename.split('-')[3]
                self.login()

        self.modo_verificacao()

    def modo_verificacao(self):
        os.system('cls')
        remessa = self.config['caminho_pasta_remessa']
        temp_verif = self.config['tempo_verificacao']

        while True:
            arquivos = iglob(os.path.join(remessa, "*.txt"))
            self.arquivo_recente = max(arquivos, key=getmtime)
            arquivo_recente = os.path.basename(self.arquivo_recente)

            if arquivo_recente not in self.db:
                self.ano = arquivo_recente.split('-')[3]
                self.login()
            else:
                self.wait_for_new_file(arquivo_recente, temp_verif)

    def wait_for_new_file(self, arquivo_recente, temp_verif):
        chars = "/—\\|"
        for char in chars:
            sys.stdout.write('\r' + 'Aguardando...' + char)
            time.sleep(0.1)
            sys.stdout.flush()
        time.sleep(temp_verif)
        self.verif_all_files()

    def login(self):
        self.driver.refresh()
        time.sleep(2)
        usuario = self.config['usuario']
        senha = self.config['senha']

        self.enter_text(By.XPATH, '/html/body/div[2]/div[3]/form/div/div[2]/div/table/tbody/tr[1]/td[2]/input', usuario)
        self.enter_text(By.XPATH, '/html/body/div[2]/div[3]/form/div/div[2]/div/table/tbody/tr[2]/td[2]/input', senha + Keys.ENTER)

        self.gerenciamento()

    def enter_text(self, by, value, text):
        element = self.driver.find_element(by, value)
        element.clear()
        element.click()
        element.send_keys(text)
        time.sleep(1)

    def convert_ano(self, ano):
        element_map = {
            '23': 'iconmenugroup_4_16', '22': 'iconmenugroup_4_15', '21': 'iconmenugroup_4_14',
            '20': 'iconmenugroup_4_13', '19': 'iconmenugroup_4_12', '18': 'iconmenugroup_4_11',
            '17': 'iconmenugroup_4_10', '16': 'iconmenugroup_4_9', '15': 'iconmenugroup_4_8',
            '14': 'iconmenugroup_4_7', '13': 'iconmenugroup_4_6', '12': 'iconmenugroup_4_5',
            '11': 'iconmenugroup_4_4', '10': 'iconmenugroup_4_3', '09': 'iconmenugroup_4_2'
        }
        return element_map.get(ano, 'iconmenugroup_4_1')

    def convert_inscricao(self, inscricao):
        element_map = {
            '23': 'iconj_id25', '22': 'iconj_id29', '21': 'iconj_id33', '20': 'iconj_id37',
            '19': 'iconj_id41', '18': 'iconj_id45', '17': 'iconj_id49', '16': 'iconj_id53',
            '15': 'iconj_id57', '14': 'iconj_id61', '13': 'iconj_id65', '12': 'iconj_id69',
            '11': 'iconj_id73', '2010': 'iconj_id77', '09': 'iconj_id81'
        }
        return element_map.get(inscricao, 'iconmenugroup_4_1')

    def gerenciamento(self):
        self.wait_for_element(By.CLASS_NAME, 'rich-pmenu')
        self.click_element_by_id(self.convert_ano(self.ano))
        self.click_element_by_id(self.convert_inscricao(self.ano))
        self.upload_file()

    def click_element_by_id(self, element_id):
        element = self.driver.find_element(By.ID, element_id)
        element.click()
        time.sleep(1)

    def upload_file(self):
        input_file = self.wait_for_element(By.XPATH, '//input[@name="uploadid:file"]')
        if input_file:
            input_file = self.driver.find_element(By.XPATH, '//input[@name="uploadid:file"]')
            input_file.send_keys(self.arquivo_recente)
            self.click_send_button()
            self.download_file()

    def click_send_button(self):
        btn_enviar = self.wait_for_element(By.XPATH, '/html/body/div[2]/div[3]/form/div[2]/table/tbody/tr/td/div[2]/div/div')
        if btn_enviar:
            btn_enviar = self.driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/form/div[2]/table/tbody/tr/td/div[2]/div/div')
            btn_enviar.click()
            time.sleep(5)

    def download_file(self):
        self.click_element_by_id("iconj_id87")
        processo_solicitacao = self.driver.find_element(By.XPATH, '//*[@id="listaSolicitacaoAtendidas:0:j_id170"]/div').text
        n_solicitacao = self.driver.find_element(By.XPATH, '//*[@id="listaSolicitacaoAtendidas:0:j_id163"]/div').text + ".txt"

        path_download = str(os.path.join(Path.home(), "Downloads"))
        arquivos = iglob(os.path.join(path_download, "*.txt"))
        files = max(arquivos, key=getmtime)
        file = os.path.basename(files)

        if file == n_solicitacao:
            shutil.move(files, os.path.join(self.config['caminho_pasta_retorno'], os.path.basename(self.arquivo_recente)))
            self.db.append(os.path.basename(self.arquivo_recente))
            self.save_json_file(self.db, "db")
            self.enviar_email(os.path.basename(self.arquivo_recente))
            self.modo_verificacao()

    def enviar_email(self, nome_arquivo):
        email = self.config['email']
        senha_email = self.config['senha_email']
        enviar_para = self.config['enviar_para']

        s = smtplib.SMTP('smtp.gmail.com:587')
        s.starttls()
        s.login(email, senha_email)
        message = f'[ + ] Arquivo Registrado na pasta com sucesso: {self.config["caminho_pasta_retorno"]}\\{nome_arquivo}\n\n Enviado por NDS Robot'
        email_msg = MIMEMultipart()
        email_msg['From'] = email
        email_msg['To'] = enviar_para
        email_msg['Subject'] = '[ Resgistro de Arquivo INEP ]'
        email_msg.attach(MIMEText(message, 'plain'))

        s.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
        print(f"[+] E-mail enviado com sucesso para {enviar_para}")
        time.sleep(2)


if __name__ == "__main__":
    try:
        bot = INEP_BOT()
    except KeyboardInterrupt:
        print("Processo interrompido pelo usuário.")
