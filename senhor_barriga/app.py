# coding=UTF-8

import os
import pandas as pd
from datetime import datetime
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

import configparser

#Lê o arquivo .ini para recuperar as mensagens personalizadas
config_parser = configparser.ConfigParser()

config_parser.read('senhor_barriga.ini', encoding='utf-8')

planilha = config_parser.get('config', 'bd')
mensagem_dias_antes_vcto = config_parser.get('mensagens', 'dias_antes_vcto')
mensagem_atraso = config_parser.get('mensagens', 'pagamento_vencido')
mensagem_dia_vcto = config_parser.get('mensagens', 'dia_vcto')
nro_dias_antes_vcto = config_parser.get('config_regras', 'nro_dias_antes_vcto')

class Whatsapp:
    
    def enviar_msg(self, telefone, nome, mensagem):
        PATH = os.path.join(os.getcwd(), "seta.png")
        data_hoje = datetime.today().strftime('%d/%m/%y')

        # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
        # com base nos dados da planilha
        try:
            link_mensagem_whatsapp = (f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}')
            webbrowser.open(link_mensagem_whatsapp)
            sleep(15)
            seta = pyautogui.locateCenterOnScreen(PATH)
            sleep(2)
            pyautogui.click(seta[0],seta[1])
            sleep(2)
            pyautogui.hotkey('ctrl','w')
            sleep(1)
        except Exception as exc:
            print(f'Não foi possível enviar mensagem para {nome}')
            print(exc)
            with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{data_hoje} - {nome} - {telefone}{os.linesep}')

def enviar_mensagens_cobranca():
    #carrega a planilha, os dados e aplica as regras para enviar mensagem
    tabela = pd.read_excel("aluguel_senhor_barriga.xlsx")
    data_hoje = datetime.strptime(datetime.today().strftime('%d/%m/%y'), '%d/%m/%y')

    wapp = Whatsapp()

    #ler todas as linhas da tabela e aplicar regras
    for index, linha in tabela.iterrows():
        nome = linha['nome']
        data_venc = linha['data_venc'].to_pydatetime()
        #email = linha['email']
        valor = float(linha['valor'])
        valor_corrigido = float(linha['valor_corrigido'])
        telefone = linha['telefone']
        status = str(linha['status']).upper()
        
        diff = (data_venc - data_hoje).days
        
        if status != 'PAGO': 
            #clientes com data de vencimento para X dias
            if (diff) == int(nro_dias_antes_vcto):
                #mensagem = (f'Olá {nome} seu aluguel no valor de R$ {valor:.2f} vence daqui {diff} dias, em {data_venc.strftime('%d/%m/%Y')}. Para sua comodidade acesse o link https://www.link_do_pagamento.com')
                mensagem = mensagem_dias_antes_vcto.replace('{nome}', nome)
                mensagem = mensagem.replace('{valor}', "{:.2f}".format(valor))
                mensagem = mensagem.replace('{nro_dias_antes_vcto}', str(abs(diff)))
                mensagem = mensagem.replace('{data_venc}', data_venc.strftime('%d/%m/%Y'))
                print(mensagem)
                wapp.enviar_msg(telefone, nome, mensagem)
            elif (diff) == 0:
                #mensagem = (f'Olá {nome} seu aluguel no valor de R$ {valor:.2f} vence hoje, dia {data_venc.strftime('%d/%m/%Y')}. Para sua comodidade acesse o link https://www.link_do_pagamento.com')
                mensagem = mensagem_dia_vcto.replace('{nome}', nome)
                mensagem = mensagem.replace('{valor}', "{:.2f}".format(valor))
                mensagem = mensagem.replace('{data_venc}', data_venc.strftime('%d/%m/%Y'))
                print(mensagem)
                wapp.enviar_msg(telefone, nome, mensagem)
            elif (diff) < 0:
                #mensagem = (f'Olá {nome} seu aluguel no valor de R$ {valor:.2f} venceu no dia {data_venc.strftime('%d/%m/%Y')}. O pagamento está atrasado em {abs(diff)} dias. Para sua comodidade acesse o link https://www.link_do_pagamento.com')
                mensagem = mensagem_atraso.replace('{data_venc}', data_venc.strftime('%d/%m/%Y'))
                mensagem = mensagem.replace('{nome}', nome)
                mensagem = mensagem.replace('{valor}', "{:.2f}".format(valor))
                mensagem = mensagem.replace('{dias}', str(abs(diff)))
                mensagem = mensagem.replace('{valor_corrigido}', "{:.2f}".format(valor_corrigido))
                print(mensagem)
                wapp.enviar_msg(telefone, nome, mensagem)
        
def main():
    enviar_mensagens_cobranca()

main()