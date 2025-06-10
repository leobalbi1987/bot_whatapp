import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
from datetime import datetime

def enviar_mensagens():
    # Abrir o WhatsApp Web
    webbrowser.open('https://web.whatsapp.com/')
    print("Aguardando carregamento do WhatsApp Web...")
    sleep(12)  # Tempo para escanear o QR Code

    # Abrir a planilha
    workbook = openpyxl.load_workbook('clientes.xlsx')
    pagina_clientes = workbook['Sheet1']

    for linha in pagina_clientes.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value
        vencimento = linha[2].value

        if nome and telefone and vencimento:
            # mensagem = f'Olá Sr(a) {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.google.com'

            try:
                link = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                webbrowser.open(link)
                print(f'Enviando mensagem para {nome} ({telefone})...')
                sleep(11)

                pyautogui.press('enter')
                print(f'Mensagem enviada para {nome}')

                sleep(5)
                pyautogui.hotkey('ctrl', 'w')
                sleep(3)

            except Exception as e:
                print(f'Erro ao enviar mensagem para {nome}: {e}')
                with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{nome},{telefone}{os.linesep}')


# Loop contínuo para verificar o horário
while True:
    agora = datetime.now()
    hora_atual = agora.strftime("%H:%M")

    if hora_atual == "09:08":
        print("⏰ Hora de enviar mensagens!")
        enviar_mensagens()
        print("✅ Mensagens enviadas. Aguardando o próximo dia...")
        sleep(5)  # Espera 5 segundos para não repetir no mesmo minuto

    sleep(1)  # Verifica a cada 1 segundo



