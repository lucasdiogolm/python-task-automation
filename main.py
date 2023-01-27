import pyautogui
import pyperclip
import time
import pandas as pd


# pyautogui.PAUSE = 2

# pyautogui.press('win')
# pyautogui.write('chrome')
# pyautogui.press('enter')
# pyautogui.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
# pyautogui.press('enter')
# time.sleep(5)

# pyautogui.click(x=414, y=270, clicks=2)
# time.sleep(3)
# pyautogui.click(x=453, y=382)
# pyautogui.click(x=1719, y=166)
# pyautogui.click(x=1534, y=556)


pyautogui.PAUSE = 2

df = pd.read_excel(r"C:workspace\automatizacao-sistemas-python\Vendas-Dez.xlsx") # Caminho do seu dataframe
print(df)
qtde = df["Quantidade"].sum()
faturamento = df["Valor Final"].sum()
print(qtde)
print(faturamento)

pyautogui.PAUSE = 2
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyautogui.write('https://mail.google.com/mail/u/0/')
pyautogui.press('enter')
time.sleep(10)
pyautogui.click(x=96, y=169)
pyautogui.click(x=1349, y=483)
pyautogui.write("exemplo@gmail.com")
pyautogui.press('tab')
pyautogui.press('tab')
pyperclip.copy("RELATÓRIO DE VENDAS + FATURAMENTO!")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

mensagem = f"""Prezados, bom dia!

O faturamento do dia anterior foi de: R${faturamento:,.2f};
A quantidade de produtos vendidos foi de: {qtde:,} unidades;

Atenciosamente, 

Lorem Impsum"""
pyperclip.copy(mensagem)
pyautogui.hotkey("ctrl", "v")

pyautogui.hotkey("ctrl", "enter")


# Use este código para cordenadas do cursor do mouse
#import pyautogui
#import time
#time.sleep(4)
#print(pyautogui.position())
