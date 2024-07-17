"""
Preciso enviar minhas mensagens de confirmação de consulta para os meus pacientes. Preciso enviar data e hora da consulta. 
"""
# Descrever os passes manuais para depois transformar isso em código
# ler a planilha, guardar informações sobre nome, telefone e data_da_consulta
# criar links personalizados do whatsapp e enviar mensagens para cada cliente/paciente da planilha.
# https://web.whatsapp.com/send?phone=5511999999999&text=Olá%20este%20é%20um%20teste
# criar um validador ou formatador do número de telefone para ter certeza que ele está no padrão

import openpyxl
import webbrowser
from time import sleep
from urllib.parse import quote
import pyautogui

# Aguardar 30 seg para que dê tempo de o usuário autentique o whatsapp web.
# Caso já esteja logado aguardar os 30 seg que a automação continuará sem problemas.

webbrowser.open('https://www.whatsapp.com/?lang=pt_BR')
sleep(10)
pyautogui.click(1289, 158)
sleep(30)

workbook = openpyxl.load_workbook('pacientes.xlsx')
paginas_clientes = workbook['Sheet1']

for linha in paginas_clientes.iter_rows(min_row=2):
    # nome, telefone, data_consulta
    nome = linha[0].value
    telefone = linha[1].value
    data_consulta = linha[2].value
    mensagem = f'Olá {nome}, sua consulta com a Dra. Tereza está marcada para o dia {
        data_consulta.strftime('%d/%m/%Y, %H:%M:%S')}. Por favor, confirmar presença.'
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={
        telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
    try:
        seta = pyautogui.locateCenterOnScreen('setinha.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
