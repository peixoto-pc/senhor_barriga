from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

class Whatsapp:
    
    def enviar_msg(self, nome, telefone, data_venc):
        print('chamou enviar_msg')
        PATH = os.path.join(os.getcwd(), "seta.png")

        # nome, telefone, vencimento
        nome = nome
        telefone = telefone
        vencimento = data_venc

        mensagem = (f'Olá {nome} seu boleto vence no dia {vencimento}. Para sua comodidade acesse o link https://www.link_do_pagamento.com')

        # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
        # com base nos dados da planilha
        try:
            link_mensagem_whatsapp = (f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}')
            webbrowser.open(link_mensagem_whatsapp)
            sleep(10)
            seta = pyautogui.locateCenterOnScreen(PATH)
            sleep(2)
            pyautogui.click(seta[0],seta[1])
            sleep(2)
            pyautogui.hotkey('ctrl','w')
        except Exception as exc:
            print(f'Não foi possível enviar mensagem para {nome}')
            print(exc)
            with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}{os.linesep}')