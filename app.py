'''
Fazer a automação das mensagens
'''

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


webbrowser.open('https://web.whatsapp.com/')
sleep(20)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_cliente = workbook['Sheet1']



for linha in pagina_cliente.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    data_compra = linha[2].value


    mensagem = f'Sr(a) {nome}, sua compra foi aprovada dia {data_compra.strftime('%d/%m/%y')}. O prazo de entrega é de 15 dias úteis a partir dessa data. '

    try:
        link_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'    
        webbrowser.open(link_whatsapp)
        sleep(20)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl' + 'w')  
    except:
            print(f'Não foi possível enviar mensagem para {nome}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{nome}, {telefone}')


