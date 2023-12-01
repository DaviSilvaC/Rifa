import time
import os
#Planilhas
import gspread
from oauth2client.service_account import ServiceAccountCredentials
#Whats
import pywhatkit as kit
#Email
import smtplib
import ssl
from email.message import EmailMessage



# Configure o escopo e as credenciais da planilha
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('./creds.json', scope)
client = gspread.authorize(creds)

# Continuar com o acesso à planilha e coleta de dados

# Abra a planilha pelo nome
planilha = client.open('Cópia de Rifa Minerva eRacing outubro 23')
# Acesse uma folha específica
folha = planilha.worksheet('Controle')

# Coleta os dados das colunas (A=1,B=2,...)
email_coluna = folha.col_values(1)
quantidade_rifas_coluna = folha.col_values(2)
nInicial_coluna = folha.col_values(3)
nFinal_coluna = folha.col_values(4)
whats_coluna = folha.col_values(5)
situacao_coluna = folha.col_values(9)
enviado_coluna = folha.col_values(11)


#Configuracao do email
email_sender = 'davicostasilva2@gmail.com'
email_password = os.environ.get("EMAILDAVI")


for i in range(len(quantidade_rifas_coluna)):
    situacao = situacao_coluna[i]
    enviado = enviado_coluna[i]
    quantidade_rifas = quantidade_rifas_coluna[i]
    nInicial = nInicial_coluna[i]
    nFinal = nFinal_coluna[i]
    whats = whats_coluna[i]
    email = email_coluna[i]
    
    if situacao == "OK" and enviado == "FALSE" and quantidade_rifas:
        message = (
                    f'Olá!! 🦉\nMuito obrigada por participar da Rifa da Minerva eRacing. 🏎🏁\n\n'
                    f'Quantidade de rifas compradas: {quantidade_rifas}\n'
                    f'Números: {nInicial}-{nFinal}\n\n'
                    f'O sorteio será realizado no Instagram da equipe, @minervaeracing .\n'
                    f'Boa sorte!!!🍀'
                   )

        if whats.isdigit():
            numero_telefone=f'+55{whats}'
            kit.sendwhatmsg_instantly(numero_telefone, message, 60, True)
            # Atualize a coluna "enviado_coluna" para "True" na linha atual
            folha.update_cell(i + 1, 11, "TRUE")  # +1 porque as linhas começam em 1

        else :
            context = ssl.create_default_context()
            em = EmailMessage()
            em['From'] = email_sender
            em['To'] = email
            em['Subject'] = 'Rifa da Minerva eRacing'
            em.set_content(message)
            
            # Log in and send the email
            try:
                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
               
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, em['To'], em.as_string())
                    folha.update_cell(i + 1, 11, "TRUE")  # +1 porque as linhas começam em 1
        
            except:
                print(f"Na linha {i + 1}: Situação: {situacao}, Mensagem a ser enviada por email não foi completada")

                        
            


# Aguarda a entrada do usuário antes de encerrar
input("Pressione Enter para sair...")


