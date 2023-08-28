import openpyxl
import pyautogui
import keyboard
from time import sleep

planilha = openpyxl.load_workbook('ProdutosCafe.xlsx')
paginaProdutos = planilha['cafe']

sleep(2)

for linha in paginaProdutos.iter_rows(min_row=2):
    nome = linha[0].value
    tituloImg = nome.split()[-1]
    categoria = linha[1].value
    descricao = linha[2].value
    preco = linha[3].value

    pyautogui.click(270, 400, duration=0.5)
    keyboard.write(nome)
    pyautogui.press(['tab', 'enter'])
    sleep(1)
    keyboard.write(f'Type={tituloImg}.png')
    sleep(1)
    pyautogui.press('enter')
    pyautogui.press('tab')

    keyboard.write(str(preco))
    pyautogui.press('tab')

    keyboard.write(categoria)
    pyautogui.press('tab')

    keyboard.write(descricao)
    pyautogui.press(['tab','tab', 'enter'])
    sleep(1)

print("Concluido")

