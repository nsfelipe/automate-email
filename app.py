import smtplib
import os
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from time import sleep
from dotenv import load_dotenv

load_dotenv()

# Configuração
HOST = os.getenv('HOST')
PORT = os.getenv('PORT')
USER = os.getenv('USER')
PASSWORD = os.getenv('PASSWORD')

# Importando a planilha

planilha = pd.read_excel('dados/beneficios.xlsx')
#planilha = planilha.astype({'VALOR ALIMENTAÇAO': 'float', 'VALOR TRANSPORTE': 'float', 'CULTURA': 'float', 'AUXILIO': 'float', 'SAUDE': 'float'})
planilha.head(50)

# Cria lista a partir do dataframe para envio

nome_list = planilha['NOME']
emails_list = planilha['EMAIL']
mesref_list = planilha['MÊS REF']
dias_uteis_list = planilha['DIAS UTEIS']
dias_aliment_list = planilha['DIAS ALIMENTAÇÃO']
valor_aliment_list = planilha['VALOR ALIMENTAÇAO']
dias_transp_list = planilha['DIAS TRANSPORTE']
valor_transp_list = planilha['VALOR TRANSPORTE']
cultura_list = planilha['CULTURA']
auxilio_list = planilha['AUXILIO']
saude_list = planilha['SAUDE']
obs_list = planilha['OBSERVAÇÕES']

# Criando objeto
print('Criando objeto servidor...')
server = smtplib.SMTP(HOST, PORT)

# Login com servidor
print('Login...')
server.ehlo()
server.starttls()
server.login(USER, PASSWORD)


# Contador de emails enviados

contador = 0 

for i in range(len(emails_list)):

    # print(f'nome: {nome_list[i]} email: {emails_list[i]}, mês ref: {mesref_list[i]}')
    
    message_html = f'''
    <html>
    <body>
        <p>Olá {nome_list[i]}!</p>
        <br>
        <h4>Segue abaixo detalhamento dos benefícios depositados referente a {mesref_list[i].strftime("%m/%Y")}:</h4>
        <br>
        <p>- Dias úteis no mês <strong>{mesref_list[i].strftime("%m/%Y")}</strong>: <strong>{dias_uteis_list[i]} dias</p></strong>
        <p>- Dias considerados para alimentação: <strong>{dias_aliment_list[i]} dias</strong></p>
        <p>- Valor alimentação: <strong>R$ {valor_aliment_list[i]}</strong></p>
        <p>- Dias considerados para transporte: <strong>{dias_transp_list[i]} dias</strong></p>
        <p>- Valor transporte: <strong>R$ {valor_transp_list[i]}</strong></p>
        <p>- Vale cultura: <strong>R$ {cultura_list[i]}</strong></p>
        <p>- Auxílio: <strong>R$ {auxilio_list[i]}</strong></p>
        <p>- Saúde: <strong>R$ {saude_list[i]}</strong></p>
        <p>- Observações: <strong>{obs_list[i]}</strong></p>
        <br>
        <p>Caso tenha alguma dúvida gentileza abrir uma soliciatação em nosso Jira: <a href='https://centraldosbeneficios.atlassian.net/servicedesk/customer/portal/15' target='_blank'>Clique aqui para acessar o portal!</a></p>
        <br>
        <br>
        <p>Atenciosamente,</p>
        <p><strong>Departamento Pessoal</strong></p>
        <p><strong>Central dos Benefícios</strong></p>
    </body>
    </html>
    '''

    email_msg = MIMEMultipart()
    email_msg['From'] = 'Gisiane Reis | Central dos Benefícios <gisianereis@centraldosbeneficios.com.br>'
    email_msg['To'] = emails_list[i]
    email_msg['Subject'] = f'Recarga de Benefícios - {mesref_list[i].strftime("%m/%Y")} - {nome_list[i]}'
    email_msg.attach(MIMEText(message_html, 'html'))

    # Enviando mensagem
    print('Enviando mensagem...')
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
    print(f'Mensagem enviada para: {nome_list[i]}')
    

    contador = contador + 1

    sleep(3)

print(f'E-mails enviados: {contador}')

# Fecha a conexão
server.quit()
