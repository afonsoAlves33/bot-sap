import pyautogui
import time

#1 Abrir o SAP

time.sleep(2)
pyautogui.click(x=867, y=1179)

#2 Entrar no PS0

time.sleep(5)
pyautogui.press("f10")
time.sleep(0.25)
pyautogui.press("down")
time.sleep(0.25)
pyautogui.press("down")
time.sleep(0.25)
pyautogui.press("down")
time.sleep(0.25)
pyautogui.press("down")
time.sleep(0.25)
pyautogui.press("down")
time.sleep(0.25)
pyautogui.press("enter")
time.sleep(0.25)


pyautogui.press("down")
time.sleep(1)
pyautogui.press("enter")

#3 Abrir a MB51

time.sleep(10)
pyautogui.click(x=220, y=156)
time.sleep(1)
pyautogui.write("mb51")
pyautogui.press("enter")

#4 Chamar a variante /BOTVSO 

time.sleep(8)
pyautogui.hotkey("shift","f5")
time.sleep(1)
pyautogui.write("BOT_VSO")
time.sleep(1)
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("del")
pyautogui.press("del")
pyautogui.press("del")
pyautogui.press("del")
pyautogui.press("del")
pyautogui.press("del")
time.sleep(1)
pyautogui.press("F8")
time.sleep(1)
pyautogui.press("F8")

#5 Selecionar como lista detalhada

time.sleep(15)
pyautogui.hotkey("ctrl","shift","F12")

#6 Baixar a planilha
time.sleep(10)
pyautogui.hotkey("shift","F4")
time.sleep(20)
pyautogui.write("real.xlsx")
time.sleep(1)
pyautogui.click(x=190, y=129)
time.sleep(1)
pyautogui.write("S:\PM\log\LOP2\Inter_Setor\RPD_35\RPD_35\DMPR_ED\Dados DMPR_ED\Atual\Dados\Dados LOP2")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
#pyautogui.press("enter")
#time.sleep(1)
pyautogui.press("tab")
time.sleep(1)
pyautogui.press("enter")

#10 Trocar a coluna "remessa" por "fornecimento"
    
time.sleep(20)
pyautogui.hotkey("ctrl","l")
time.sleep(1)

pyautogui.write("remessa")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("esc")
time.sleep(1)
pyautogui.write("Fornecimento")
time.sleep(1)
pyautogui.press("enter")

#11 Fechar Excel

time.sleep(1)
pyautogui.click(x=340,y=1035,button='right')
time.sleep(1)
pyautogui.click(x=296, y=985)
time.sleep(1)
pyautogui.press("enter")

#12 Fechar SAP

time.sleep(1)
pyautogui.click(x=131, y=1031)
time.sleep(1)
pyautogui.click(x=173, y=819)
time.sleep(1)
pyautogui.press("enter")

#13 fit to page (BI)

time.sleep(1)
pyautogui.click(x=1726, y=523)