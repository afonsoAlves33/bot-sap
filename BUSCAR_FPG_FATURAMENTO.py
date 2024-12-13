import pyautogui
import time

from window_manager import WindowManager



def save_excel_sheet(excel_sheet_name: str, variant: str, code: str = "mb51", saving_path: str = r"\\bosch.com\dfsrb\DfsBR\loc\Ca1\ED\mse\MOE1\Dados\Planejamento_Tecnico\19 - Dados_BI_online\03 - Base de dados\teste_bot"):
    """
    Saves a excel spreadsheet by following a default route on SAP
    """
    pyautogui.click(x=339, y=741)

    # Maximizando a tela
    time.sleep(10)
    time.sleep(5)
    WindowManager.focus_window_by_title("SAP Logon 800")
    time.sleep(3)
    WindowManager.maximize_window_by_title("SAP Logon 800")
    time.sleep(5)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")

    time.sleep(10)

    WindowManager.focus_window_by_title("PS0(1)/011 SAP Easy Access")
    pyautogui.write(code)
    time.sleep(1)
    pyautogui.press("enter")


    # Chamar a variante
    time.sleep(8)
    WindowManager.focus_window_by_title("PS0(1)/011 Material Document List")
    pyautogui.hotkey("shift","f5")
    time.sleep(2)
    WindowManager.focus_window_by_title("PS0(1)/011 Find Variant")
    pyautogui.write(variant)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    WindowManager.focus_window_by_title("PS0(1)/011 Material Document List")
    pyautogui.press("F8")

    # Selecionar como lista detalhada
    time.sleep(15)
    WindowManager.focus_window_by_title("PS0(1)/011 Material Document List")
    pyautogui.hotkey("ctrl","shift","F12")

    # Baixar a planilha
    time.sleep(3)
    WindowManager.focus_window_by_title("PS0(1)/011 Material Document List")
    pyautogui.hotkey("shift","F4")
    time.sleep(7)

    # pyautogui.hotkey('win', 'left')
    # time.sleep(1)
    WindowManager.focus_window_by_title("Save As")
    pyautogui.click(x=163, y=423)
    time.sleep(7)

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
    WindowManager.focus_window_by_title("Confirm Save As")
    pyautogui.press("left")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(8)

    try: 
        WindowManager.focus_window_by_title("SAP GUI Security")
        pyautogui.click(x=311, y=430)
        time.sleep(3)
    except Exception as e:
        print("Probably the window SAP GUI Security did not open")
    print("done")

    WindowManager.close_window_by_title(f"{excel_sheet_name} - Excel")
    time.sleep(2)
    WindowManager.close_window_by_title("PS0(1)/011 Material Document List")
    time.sleep(2)
    WindowManager.close_window_by_title("SAP Logon 800")
    print("Sap Closed")


if __name__ == "__main__":
    save_excel_sheet(excel_sheet_name="EXPORT_MB51_601_602.XLSX", variant="FPG_FTM")
    # s_alr_87013625