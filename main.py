# Importação da biblioteca openpyxl para permitir a leitura do contéudo das planilhas, biblioteca quote para permitir acesso a links e webbrowser para executar o navegador
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

def main():
    webbrowser.open('https://web.whatsapp.com/')
    sleep(15)

    # Carregamento dos dados das planilhas e especificação das páginas disponiveis
    workbook = openpyxl.load_workbook('contatos.xlsx')
    paginas_contatos = workbook['Planilha1']

    # Laço de repetição para iterar sobre as colunas da planilha, sendo específicada a coluna inicial começando em 2
    for linha in paginas_contatos.iter_rows(min_row = 2):
        # Leitura dos campos de nome e telefone da planilha do LibreOffice
        nome = linha[0].value
        telefone = linha[1].value

        # Criação da mensagem que será enviada para o usuário
        mensagem = f'Olá, {nome}. É um prazer revê-lo!'
        try:
            # Desenvolvimento do link personalizado para enviar mensagens dentro do WhatApp Web
            link_mensagem_whatsapp = f'https://api.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(15)
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(5)
            pyautogui.click(seta[0],seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl','w')
            sleep(5)
        except:
            print(f'Não foi possivel enviar mensagem para {nome}.')
            with open('erros.csv', 'a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}')

main()