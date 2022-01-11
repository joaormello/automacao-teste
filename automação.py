import pyautogui
import pyperclip
import time
import pandas as pd
import openpyxl

def achar_posicao():
    time.sleep(6)
    y = pyautogui.position()
    print(y)

#entrando no Chorme
pyautogui.press('win')
time.sleep(5)
pyautogui.write('Google C')
pyautogui.press('enter')
time.sleep(5)
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga') # por causa do caracter especial
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#cliks p baixar
time.sleep(5)
pyautogui.click(x=1722, y = 279, clicks = 2 ) # clicando em explorar
time.sleep(5)
pyautogui.rightClick(x= 1734, y = 360,) # clicando no arquivo
time.sleep(5)
pyautogui.click(x=1867, y=630)
time.sleep(10)


#pandas
tabela = pd.read_excel(r"C:\Users\joaop\Downloads\Vendas - Dez.xlsx")
print(tabela)

# calculando o faturamento e qtds de produtos vendidos
faturamento = tabela['Valor Final'].sum()
qtade_produtos = tabela['Quantidade'].sum()

# automatizando o envio de email

pyautogui.hotkey('ctrl', 't')
time.sleep(4)
pyautogui.write('https://mail.google.com/mail/u/0/#inbox')
pyautogui.press('enter')
time.sleep(6)
pyautogui.click(x=1437, y=179)
time.sleep(6)
pyautogui.write('joaopedrormello@hotmail.com ') # Destinatário do email
time.sleep(5)
pyautogui.press('tab')
pyperclip.copy('Email automático sobre o faturamento referente ao mês de Dezembro')
pyautogui.hotkey('ctrl', 'v')
time.sleep(5)
pyautogui.press('tab')
pyperclip.copy(f'Boa Tarde.\n\n\n O faturamento do mês de Dezembro foi de {faturamento} reais, com {qtade_produtos} de produtos vendidos.\n\n\n Abraços. \n\n\n Automação do João.'
                )
pyautogui.hotkey('ctrl', 'v')
time.sleep(5)
pyautogui.click(x=2201, y=635)