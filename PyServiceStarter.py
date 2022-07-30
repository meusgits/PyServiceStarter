import wmi, os, time, win32api
from time import sleep
from wmi import WMI
from os import system
from win32api import GetComputerName

os.system('cls')

servidor = win32api.GetComputerName()
services = wmi.WMI()
servico = str('wuauserv')
tempo = 10


def teste_se_esta_on(nome):

    conte = 1

    while conte <= 3:


        def nao_iniciou():
            print(f'Luciano, tentei iniciar o {nome} em {servidor} por {conte} vezes em intervalos de {tempo} segundos, mas não iniciou! É necessária a sua intervanção.')
            exit()


        for s in services.Win32_Service(Name=servico, State="Stopped"):
            print(f'Luciano o {nome} em {servidor} não iniciou.')
            conte += 1
            s.StartService()
            time.sleep(tempo)
            if conte == 4:
                nao_iniciou()


        for s in services.Win32_Service(Name=servico, State="Running"):
            print(f'Luciano o {nome} em {servidor} foi iniciado com sucesso!')
            conte = 4


while True:
    
    
    for s in services.Win32_Service(Name=servico, State="Stopped"):
        nome = s.DisplayName
        print(f'Luciano o {nome} parou em {servidor}.')
        print(f'Enviando comando para iniciar o {nome} em {servidor}!')
        s.StartService()
        time.sleep(tempo)
        teste_se_esta_on(nome)


    for s in services.Win32_Service(Name=servico, State="Running"):
        nome = s.DisplayName
        print(f'Luciano o {nome} em {servidor} está iniciado!')
        

