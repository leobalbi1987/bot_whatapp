import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
print("Aguardando carregamento do WhatsApp Web...")
sleep(25)  # Tempo para escanear o QR Code

# Abrir a planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    if nome and telefone and vencimento:
        mensagem = f'Olá Sr(a) {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.google.com'

        try:
            # Criar link do WhatsApp com a mensagem
            link = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link)
            print(f'Enviando mensagem para {nome} ({telefone})...')
            sleep(11)  # Tempo para carregar o chat

            # Pressionar ENTER para enviar a mensagem
            pyautogui.press('enter')
            print(f'Mensagem enviada para {nome}')

            sleep(3)  # Tempo para o envio antes de fechar a aba
            pyautogui.hotkey('ctrl', 'w')  # Fecha a aba
            sleep(2)

        except Exception as e:
            print(f'Erro ao enviar mensagem para {nome}: {e}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}{os.linesep}')
