import pyautogui
import time
from openpyxl import load_workbook

nomeCaminhoArquivo = 'C:\\Projetos Python\\Projeto Filtra Excel\\Lista Gestores.xlsx'
planilha_aberta = load_workbook(filename=nomeCaminhoArquivo)
sheet_selecionada = planilha_aberta["Arquivos"]  # Acessa a aba que estiver dentro de wb

for linha in range(2, len(sheet_selecionada['A']) + 1):
    nome_gestor = sheet_selecionada[f'A{linha}'].value

    pyautogui.PAUSE = 1
    pyautogui.press('win')
    pyautogui.typewrite('executar')
    pyautogui.press('enter')
    pyautogui.write("C:\\UiPath\\Envio de Emails\\Envio de Email RPA com anexo\\Envio de Email RPA RH Banco de Horas\\Bando de horas Outubro-Novembro.xlsx")
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.click(x=495, y=247)
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.write(nome_gestor)
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl', 't')
    pyautogui.hotkey('ctrl', 't')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('alt', 'shift', 'f1')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(8)
    pyautogui.press('alt')
    pyautogui.press('c')
    pyautogui.press('o')
    pyautogui.press('t')
    pyautogui.hotkey('ctrl', 'pgdown')
    pyautogui.click(x=246, y=989, button='right')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.press('alt')
    pyautogui.press('a')
    pyautogui.press('a')
    pyautogui.press('o')
    pyautogui.typewrite(nome_gestor)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
