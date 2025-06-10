 

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(15)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimentort4rtgjm
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    # mensagem = f'Olá {nome}  seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://spe.mangaratiba.rj.gov.br/IPTU/guia.aspx'
    
    mensagem = f'Olá {nome} regularize sua situação  entre no link https://spe.mangaratiba.rj.gov.br/IPTU/guia.aspx'

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(12)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        pyautogui.hotkey('ctl','w')
           
    except:
       
     
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
    
