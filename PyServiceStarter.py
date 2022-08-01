import wmi
import os
import time
import win32api
import smtplib
from time import sleep
from wmi import WMI
from os import system
from win32api import GetComputerName
from datetime import date
from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

os.system('cls')

servidor = win32api.GetComputerName()
services = wmi.WMI()

servico = f'\\\\srv-sp0254\\pasta-remota$\\Monitora-Service\\servico.txt'
with open(servico, 'r', encoding="utf8") as lendoServico:
    servico = lendoServico.readline().strip('\n')
    lendoServico = str(servico)

tempo = f'\\\\srv-sp0254\\pasta-remota$\\Monitora-Service\\intervalo.txt'
with open(tempo, 'r', encoding="utf8") as lendoLimite:
    tempo = lendoLimite.readline().strip('\n')
    tempo = int(tempo)

e_mails = []

arquivo = open('\\\\srv-sp0254\\pasta-remota$\\Monitora-Service\\E-mail.txt', 'r', encoding="utf8")

for linhas in arquivo:  # Vou percorrer o arquivo fonte onde estão os e-mails linha a linha e adicioná-los a minha lista vázia
    # Adicionando os a-emails um a um a lista de e-mails vázia, retirando o caracter de quebra de linha (\n)
    e_mails.append(linhas.strip('\n'))

arquivo.close()  # Fechando o arquivo de onde foram pegos os e-mails


def dipare_emails(servico, servidor, hoje, agora, conte, tempo):

    for e_mail in e_mails:  # Pegando os e-mails da lista e disparando a mensagem para cada um deles

        de = 'informativo@kitani.com.br'  # E-mail que será usado para o disparo dos e-mails
        para = e_mail  # E-mail do remetente pego na lista de e-mails

        corpo = MIMEMultipart('alternative')

        # Abaixo entre as três aspas """ esta a mensagem de texto simples que será enviada, em auternativa ao HTML
        # caso o cliente de e-mail do destinatário não de suporte a HTML
        # A primeira \ após as três aspas indica que a mensagem iniciará logo na linha de baixo
        # O '\n' quebra linha e o '\t' cria uma tabulação.

        if conte == 0:
            text = f"""\
            \n\Informativo:\n
            \n\n\tO serviço {nome} no servidor {servidor} parou hoje {hoje} às {agora}, foi enviado comando para iniciar o serviço.\n\tCaso o {nome} tenha iniciado com sucesso você não receberá mais e-mail, entendasse que está tudo ok.
            """

            html = f"""\
    <!DOCTYPE html>
    <html>
    <head>
    </head>
    <body>
    <p><strong><span style="font-family: 'trebuchet ms', geneva, sans-serif; font-size: 12pt;">Informativo</span></strong></p>
    <p><span style="font-family: 'trebuchet ms', geneva, sans-serif; font-size: 12pt;">O servi&ccedil;o <strong>{nome}</strong> no servidor <strong>{servidor}</strong> parou hoje <strong>{hoje}</strong> &agrave;s <strong>{agora}</strong>, foi enviado comando para iniciar o servi&ccedil;o.<br />Caso o <strong>{nome}</strong> tenha iniciado com sucesso voc&ecirc; n&atilde;o receber&aacute; mais e-mail, entendasse que est&aacute; tudo ok.</span></p>
    </body>
    </html>
            """
        else:
            text = f"""\
            \n\Informativo\n
            \n\n\tTentei iniciar o serviço {servico} no servidor {servidor} por {conte}x em intervalos de {tempo} segundos, mas o mesmo não iniciou.\n\tNeste caso é necessário a sua intervenção.
            """

            html = f"""\
    <!DOCTYPE html>
    <html>
    <head>
    </head>
    <body>
    <p><strong><span style="font-family: 'trebuchet ms', geneva, sans-serif; font-size: 12pt;">Informativo</span></strong></p>
    <p><span style="font-family: 'trebuchet ms', geneva, sans-serif; font-size: 12pt;">Tentei iniciar o servi&ccedil;o <strong>{servico}</strong> no servidor <strong>{servidor}</strong> por <strong>{conte}x</strong> em intervalos de <strong>{tempo}</strong> segundos, mas o mesmo n&atilde;o iniciou.<br /> Neste caso &eacute; necess&aacute;rio a sua interven&ccedil;&atilde;o.</span></p>
    </body>
    </html>
            """


        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')

        corpo.attach(part1)
        corpo.attach(part2)

        corpo['From'] = f'{servidor}@domínio.com.br'
        corpo['To'] = e_mail
        corpo['Reply-To'] = 'no-reply@domínio.com.br'
        corpo['Subject'] = f'Monitoramento do serviço {servico} no servidor {servidor}'

        usuario = 'usuário@domínio'
        senha = 'Senha@S#n@'

        # Conectando ao servidor de e-mail, autenticando e disparando a mensagem
        with smtplib.SMTP("192.168.20.254", 587) as server:

            server.login(usuario, senha)
            server.sendmail(de, para, corpo.as_string())
            print("E-mail enviado com sucesso para", e_mail, "!")
            server.quit()

    print('\n\nForam enviados %.4d e-mails.' % len(e_mails))


def teste_se_esta_on(nome):

    conte = 1

    while conte <= 3:

        def nao_iniciou():
            print(f'Tentei iniciar o {nome} em {servidor} por {conte} vezes em intervalos de {tempo} segundos, mas não iniciou! É necessária a sua intervanção.')
            dipare_emails(servico, servidor, hoje, agora, conte, tempo)
            exit()

        for s in services.Win32_Service(Name=servico, State="Stopped"):
            print(f'O {nome} em {servidor} não iniciou, tantando pela {conte}x.')
            # os.system(f'msg luciano /SERVER:srv-019 /TIME:1800 O {servico} parou às {agora} do dia {hoje} no servidor {servidor}.')
            conte += 1
            s.StartService()
            time.sleep(tempo)
            if conte == 4:
                nao_iniciou()

        for s in services.Win32_Service(Name=servico, State="Running"):
            print(f'O {nome} em {servidor} foi iniciado com sucesso!')
            conte = 4


while True:

    conte = 0
    data_atual = date.today()
    hora_atual = time.localtime()
    hoje = data_atual.strftime('%d/%m/%Y')
    agora = f'{hora_atual.tm_hour}:{hora_atual.tm_min}'

    for s in services.Win32_Service(Name=servico, State="Stopped"):
        nome = s.DisplayName
        s.StartService()
        print(f'O {nome} parou em {servidor}.')
        dipare_emails(servico, servidor, hoje, agora, conte, tempo)
        time.sleep(tempo)
        teste_se_esta_on(nome)

    for s in services.Win32_Service(Name=servico, State="Running"):
        nome = s.DisplayName
        print(f'O {nome} em {servidor} está iniciado!')
