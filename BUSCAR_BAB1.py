import pyautogui
import time


def save_excel_sheet(excel_sheet_name: str, variant: str, code: str = "mb51", saving_path: str = r"\\bosch.com\dfsrb\DfsBR\loc\Ca1\ED\mse\MOE1\Dados\Planejamento_Tecnico\19 - Dados_BI_online\03 - Base de dados\teste_bot"):
    """
    Saves a excel spreadsheet by following a default route on SAP
    """
    # time.sleep(5)
    # print(pyautogui.position())
    # time.sleep(3)
    # pyautogui.click(x=339, y=741)

    # # Maximizando a tela
    # time.sleep(10)
    # pyautogui.press("f10")
    # time.sleep(0.5)
    # pyautogui.press("down")
    # time.sleep(0.5)
    # pyautogui.press("down")
    # time.sleep(0.5)
    # pyautogui.press("down")
    # time.sleep(0.5)
    # pyautogui.press("down")
    # time.sleep(0.5)
    # pyautogui.press("down")
    # time.sleep(0.5)
    # pyautogui.press("enter")
    # time.sleep(1)
    # pyautogui.press("enter")
    # time.sleep(1)
    # pyautogui.press("enter")

    # # Abrir a MB51
    # time.sleep(10)
    # # pyautogui.click(x=125, y=148)
    # # time.sleep(1)
    # pyautogui.write(code)
    # time.sleep(1)
    # pyautogui.press("enter")

    # # Chamar a variante
    # time.sleep(8)
    # pyautogui.hotkey("shift","f5")
    # time.sleep(2)
    # pyautogui.write(variant)
    # time.sleep(1)
    # pyautogui.press("enter")
    # time.sleep(1)
    # pyautogui.press("F8")

    # # Selecionar como lista detalhada
    # time.sleep(150)
    time.sleep(9)
    pyautogui.hotkey("shift","F7")
    time.sleep(3)
    pyautogui.write("3")
    time.sleep(3)
    pyautogui.press("enter")


    # Clicando em crédito
    time.sleep(3)
    pyautogui.doubleClick(x=427, y=503)
    time.sleep(3)
    pyautogui.click(x=134, y=243)
    time.sleep(3)
    pyautogui.press("enter")

    time.sleep(1200)

    pyautogui.hotkey("ctrl","F8")
    time.sleep(3)
    pyautogui.click(x=514, y=310)
    time.sleep(3)
    pyautogui.hotkey("ctrl","a")
    time.sleep(3)
    pyautogui.press("F7")
    time.sleep(3)
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.click(x=641, y=654)
    time.sleep(3)
    pyautogui.hotkey("ctrl","shift", "F7")
    time.sleep(3)
    pyautogui.click(x=163, y=423)
    # time.sleep(3)
    # pyautogui.press("enter")
    time.sleep(3)

    #Escrever o nome da planilha
    pyautogui.write(excel_sheet_name.upper())
    time.sleep(4)
    pyautogui.hotkey('ctrl', 'l')
    time.sleep(2)
    pyautogui.write(saving_path)
    time.sleep(4)
    pyautogui.press("enter")
    time.sleep(6)
    # pyautogui.press("enter")
    pyautogui.click(x=524, y=630)
    time.sleep(1)
    pyautogui.press("left")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    print("done")
    time.sleep(9)
    pyautogui.hotkey("alt","F4")
    time.sleep(2)
    pyautogui.hotkey("alt","F4")
    time.sleep(1.5)
    pyautogui.click(x=233, y=444)
    time.sleep(2)
    pyautogui.hotkey("alt","F4")
    print("Sap Closed")


if __name__ == "__main__":
    save_excel_sheet(excel_sheet_name="EXPORT_BAB_2024_consolidado_PPC_BI.XLSX", variant="BAB2024", code="s_alr_87013625")
    # s_alr_87013625    
