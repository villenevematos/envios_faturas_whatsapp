# importar bibliotecas
from imap_tools import MailBox, AND
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from pathlib import Path
from datetime import datetime
import time

# login
username = "villeneve01@outlook.com"
password = "villeneve229125145225"
meu_email = MailBox('imap-mail.outlook.com').login(username, password)

def esperar_elemento(elemento):
    while len(navegador.find_elements(By.CLASS_NAME, elemento)) < 1:
        time.sleep(1)
    time.sleep(1)

# pegar emails com um anexo específico
lista_emails = meu_email.fetch(AND(from_="fatura.digital@faturamento.copel.com"))
hoje = datetime.now()
mes_ano = f"{hoje.strftime('%b')} {hoje.strftime('%Y')}"
for email in lista_emails:
    if len(email.attachments) > 0:
        for anexo in email.attachments:
            # Se o mês e ano estiver dentro da data do e-mail, baixe o arquivo.
            if mes_ano in email.date_str:
                informacoes_anexo = anexo.payload
                with open("fatura_copel.pdf", "wb") as arquivo:
                    arquivo.write(informacoes_anexo)
                
                # abrir navegador
                servico = Service(ChromeDriverManager().install())
                options = webdriver.ChromeOptions()
                options.add_argument(r'user-data-dir=C:\Users\VILLE\AppData\Local\Google\Chrome\User Data\Profile Selenium')
                navegador = webdriver.Chrome(service=servico, options=options)
                navegador.maximize_window()
                
                # enviar anexo
                caminho_arquivo = str(Path.cwd() / 'fatura_copel.pdf')
                lista_telefone = [('Bianca', '1796696623'), ('Vinicius', '4179555328'), ('Diego', '1188536369')]
                for nome, telefone in lista_telefone:
                    link = f'https://web.whatsapp.com/send?phone={telefone}'
                    navegador.get(link)
                                                                                
                    # esperando carregar o WhatsApp
                    esperar_elemento('ercejckq')
                    time.sleep(2)
                    
                    # verificar se o número é inválido
                    if len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
                        # anexar
                        navegador.find_element(By.CLASS_NAME, '_1OT67').click()
                        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input').send_keys(caminho_arquivo)
                        esperar_elemento('_3wFFT')

                        # enviar
                        navegador.find_element(By.CLASS_NAME, '_3wFFT').click()
                        time.sleep(5)
                        print(f'ENVIADO COM SUCESSO PARA O {nome.upper()}')