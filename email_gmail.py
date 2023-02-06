##
import smtplib
import email.message
import pandas
##
dados = pandas.read_excel("C:/Users/User/Desktop/dados.xlsx")
##

for i in range(0, len(dados)):
    corpo_excel = dados["corpo do email"][i]

    def enviar_email():
        corpo_email = """
        escrita padrao para todos os emails
        
        envio de mensagem especifica para cada email
        {}
        \n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n
        \n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\nÇ""".format(corpo_excel)

        msg = email.message.Message()
        msg['Subject'] = "AUTOMAÇÃO DE EMAILS"
        msg['From'] = 'xxxxxxxxxxxxxxxxxxxxx@gmail.com'
        msg['To'] = dados["emails"][i]

        # para senha sera gerada pelo gmail
        """
        vai na inicial do seu nome
        gerenciar configuraçoes de segurança
        segurança
        como fazer login no google
        ativa verificação de 2 etapas
        senha de app
        """
        password = 'eypdwxulptyppmjt'

        msg.add_header('Content-Type', 'text')  # se quiser escrever com HTML "text/html"
        msg.set_payload(corpo_email)

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print(f'Email {i+1} enviado')
    enviar_email()

##
