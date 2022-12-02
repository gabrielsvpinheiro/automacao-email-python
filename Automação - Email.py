get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')

import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

# Passo 1
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/u/0/folders/161aOo_wL-k2BBqqCb4b40PL2PvP_0rgA")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(2)

# Passo 2
pyautogui.click(x=406, y=297, clicks=2)
time.sleep(2)

# Passo 3
pyautogui.click(x=406, y=297)
time.sleep(1)
pyautogui.click(x=1159, y=184)
time.sleep(1)
pyautogui.click(x=1066, y=590)

# Passo 4
tabela = pd.read_excel(r"C://Users/gabri/Downloads/Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# Passo 5
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

#Passo 6
pyautogui.click(x=86, y=198)

pyperclip.copy("gabriel.silva.pinheiro2@gmail.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto = f"""Prezados, bom dia
          
O faturamento de ontem foi de: {faturamento}
A quantidade de produtos foi de: {quantidade}
          
Abs,
Bot Python"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")


# Para encontrar posição dos botões</h3>

#time.sleep(5)
#pyautogui.position()

