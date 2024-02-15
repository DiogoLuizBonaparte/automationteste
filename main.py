from datetime import time

import openpyxl
import pyperclip
import pyautogui
import pyscreeze
import time


pyautogui.click(501,752, duration=1)

workbook = openpyxl.load_workbook('clientes.xlsx')
sheet_clients = workbook['Clients']


for linha in sheet_clients:

    create_cliente = pyautogui.click(1073,333, duration=1)

    client = linha[0].value
    pyperclip.copy(client)
    pyautogui.click(371,375, duration=0.5)
    pyautogui.hotkey('ctrl','v')

    profissao =linha[1].value
    pyperclip.copy(profissao)
    pyautogui.click(367,431,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1000,487, duration=0.5)



