import pandas as pd
import smtplib
import email.message
from dotenv import load_dotenv
import os

load_dotenv()

def gerar_corpo_email():
    ## importando a base de dados
    tabela_vendas = pd.read_excel('Vendas.xlsx')
    pd.set_option('display.max_columns', None)

    ## Faturamento por loja
    faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby("ID Loja").sum()

    ## Quantidade de produtos vendidos por loja
    quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

    ## Ticket médio por produto em cada loja
    ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
    ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})

    corpo_email = f'''
    <p>Segue o relatório de vendas por loja:</p>

    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

    <p>Quantidade vendida:</p>
    {quantidade.to_html()}

    <p>Ticket médio:</p>
    {ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}
    '''

    return corpo_email

def enviar_email():  
    corpo_email = gerar_corpo_email()

    msg = email.message.Message()
    msg['Subject'] = "Relatório"
    msg['From'] = os.getenv('FROM')
    msg['To'] = os.getenv('TO')
    password = os.getenv('PASSWORD') 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()
