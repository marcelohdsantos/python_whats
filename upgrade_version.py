import webbrowser
import pyautogui
import openpyxl
import time
from urllib.parse import quote
import os


def validate_phone(phone):
    phone_formated = ''.join(filter(str.isdigit, phone))

    if len(phone_formated) >= 12:
        return phone_formated
    else:
        raise ValueError(f"Número de telefone inválido: {phone}")


def open_whatsapp_web():
    webbrowser.open('https://web.whatsapp.com/?lang=pt_BR')
    time.sleep(15)


def send_message(link):
    webbrowser.open(link)
    time.sleep(10)
    try:
        if os.path.exists('setinha.png'):
            arrow = None
            while arrow is None:
                arrow = pyautogui.locateCenterOnScreen('setinha.png')
                time.sleep(1)
            pyautogui.click(arrow[0], arrow[1])
            time.sleep(2)
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(5)
        else:
            raise FileNotFoundError("Imagem 'setinha.png' não encontrada")
    except Exception as e:
        print(f"Erro ao enviar mensagem: {e}")


def main():
    open_whatsapp_web()

    workbook = openpyxl.load_workbook('pacientes.xlsx')
    pages_clients = workbook['Sheet1']

    for line in pages_clients.iter_rows(min_row=2):
        name = line[0].value
        phone = line[1].value
        date_information = line[2].value
        try:
            phone_formated = validate_phone(phone)
            message = f'Olá {name}, sua consulta com a Dra. Tereza está marcada para o dia {
                date_information.strftime("%d/%m/%Y, %H:%M:%S")}. Por favor, confirmar presença.'
            link_message_whatsapp = f'https://web.whatsapp.com/send?phone={
                phone_formated}&text={quote(message)}'
            send_message(link_message_whatsapp)
        except Exception as e:
            print(f'Não foi possível enviar mensagem para o {name}: {e}')


if __name__ == "__main__":
    main()
