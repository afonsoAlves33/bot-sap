import pyautogui
import time


def save_excel_sheet(excel_sheet_name: str, variant: str, code: str = "mb51", saving_path: str = r"\\bosch.com\dfsrb\DfsBR\loc\Ca1\ED\mse\MOE1\Dados\Planejamento_Tecnico\19 - Dados_BI_online\03 - Base de dados\teste_bot"):
    """
    Saves a excel spreadsheet by following a default route on SAP
    """
    pyautogui.click(x=339, y=741)

    # Maximizando a tela
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
    time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")

    # Abrir a MB51
    time.sleep(10)
    # pyautogui.click(x=125, y=148)
    # time.sleep(1)
    pyautogui.write(code)
    time.sleep(0.25)
    pyautogui.press("enter")

    # Chamar a variante
    time.sleep(8)
    pyautogui.hotkey("shift","f5")
    time.sleep(1)
    pyautogui.write(variant)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("F8")

    # Selecionar como lista detalhada
    time.sleep(15)
    pyautogui.hotkey("ctrl","shift","F12")

    # Baixar a planilha
    time.sleep(3)
    pyautogui.hotkey("shift","F4")
    time.sleep(3)
    pyautogui.press("enter")
    time.sleep(9)
    # pyautogui.hotkey('win', 'left')
    # time.sleep(1)
    pyautogui.write(excel_sheet_name)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'l')
    time.sleep(1)
    pyautogui.write(saving_path)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(5)
    pyautogui.press("enter")
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
    save_excel_sheet(excel_sheet_name="EXPORT_MB51_131_132.XLSX", variant="FPG_PROD")
    # s_alr_87013625