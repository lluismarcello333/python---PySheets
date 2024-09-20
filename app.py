#ler dados da planilha
# inserir cada célula de cada linha em um campo do sistema

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('planilha_teste.xlsx')
teste_sheet = workbook['Sheet1']

for linha in teste_sheet.iter_rows(min_row=2):
    # Coluna 1
    pyautogui.click(1897,232,duration=1.5)
    pyautogui.write(str(linha[0].value))
    # Coluna 2
    pyautogui.click(1897,232,duration=1.5)
    pyautogui.write(linha[1].value)
    # Coluna 3
    pyautogui.click(1897,232,duration=1.5)
    pyautogui.write(linha[2].value)
    # Coluna 4
    pyautogui.click(1897,232,duration=1.5)
    pyautogui.write(str(linha[3].value))
    # Coluna 5
    pyautogui.click(1897,232,duration=1.5)
    pyautogui.write(str(linha[4].value))