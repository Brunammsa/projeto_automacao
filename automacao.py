import pyperclip
import pyautogui
import time
import pandas as pd

pyautogui.PAUSE = 1

# iniciando automação com abertura do navegador

pyautogui.press('winleft')
time.sleep(3)
pyautogui.write('chrome')
time.sleep(3)
pyautogui.press('enter')
time.sleep(3)
pyautogui.hotkey('ctrl', 't')
time.sleep(3)

link = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'
pyperclip.copy(link) # necessá3rio utilizar o pyperclip copy para evitar erros com caracteres especiais
pyautogui.hotkey('ctrl', 'v') # hotkey utiliza ataplhos no teclado para otimizar controles
pyautogui.press('enter')
time.sleep(3)

# pyautogui.position() localizando a posição do mouse em tela
pyautogui.click(x=461, y=332, clicks=2)
time.sleep(3)
pyautogui.click(x=461, y=332, clicks=1)
time.sleep(5)
pyautogui.click(x=1155, y=208, clicks=1)
time.sleep(3)
pyautogui.click(x=998, y=650, clicks=1)
time.sleep(5)

caminho_de_leitura = pd.read_excel(r'/home/bruna/Downloads/Vendas - Dez.xlsx')
print(caminho_de_leitura)
time.sleep(3)

faturamento = caminho_de_leitura['Valor Final'].sum()
quantidade_produtos = caminho_de_leitura['Quantidade'].sum()

pyautogui.hotkey('ctrl', 't')
time.sleep(3)

link = 'https://mail.google.com/mail/u/0/#inbox'
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(7)

pyautogui.click(x=167, y=230, clicks=1)
time.sleep(3)

pyautogui.write('brunamelo.sa@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')

# aqui eu não tenho necessidade de escrever com cópia ou oculto
assunto = 'Relatório de vendas'

pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

corpo_email = f"""
Prezados, bom dia,

O faturamento de ontem foi de R$ {faturamento:,.2f}
a quantidade de produtos foi de {quantidade_produtos}

Abs.

Bruna M.
"""

pyperclip.copy(corpo_email)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')
