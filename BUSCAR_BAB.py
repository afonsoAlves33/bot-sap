import pyautogui
import time
from window_manager import WindowManager

"""
    O nome das janelas do windows está hard coded então é possível que
    quebre caso eu mude o idioma do pc ou o sap atualiaze, etc.
"""


def save_excel_sheet(excel_sheet_name: str, variant: str, code: str = "mb51", saving_path: str = r"\\bosch.com\dfsrb\DfsBR\loc\Ca1\ED\mse\MOE1\Dados\Planejamento_Tecnico\19 - Dados_BI_online\03 - Base de dados\teste_bot"):
    """
    Saves a excel spreadsheet by following a default route on SAP
    """
    time.sleep(5)
    print(pyautogui.position())
    time.sleep(3)
    pyautogui.click(x=339, y=741)

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
    WindowManager.focus_window_by_title("PS0(1)/011 Cost Centers: Actl/Target/Variance: Selection")
    pyautogui.hotkey("shift","f5")
    time.sleep(2)
    WindowManager.focus_window_by_title("PS0(1)/011 Find Variant")
    pyautogui.write(variant)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    WindowManager.focus_window_by_title("PS0(1)/011 Cost Centers: Actl/Target/Variance: Selection")
    pyautogui.press("F8")

    # Selecionar como lista detalhada

    time.sleep(150)
    time.sleep(9)
    WindowManager.focus_window_by_title("PS0(1)/011 Cost Centers: Actual/Target/Varianc")
    pyautogui.hotkey("shift","F7")
    time.sleep(3)
    WindowManager.focus_window_by_title("PS0(1)/011 Summation Levels")
    pyautogui.write("3")
    time.sleep(3)
    pyautogui.press("enter")


    # Clicando em crédito

    time.sleep(3)
    WindowManager.focus_window_by_title("PS0(1)/011 Cost Centers: Actual/Target/Varianc")
    pyautogui.doubleClick(x=427, y=503)
    time.sleep(3)
    WindowManager.focus_window_by_title("PS0(1)/011 Select Report")
    pyautogui.click(x=134, y=243)
    time.sleep(3)
    pyautogui.press("enter")

    time.sleep(1200)

    WindowManager.focus_window_by_title("PS0(1)/011 Display Actual Cost Line Items for Cost Centers")
    pyautogui.hotkey("ctrl","F8")
    time.sleep(3)
    WindowManager.focus_window_by_title("PS0(1)/011 Change Layout")
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
    WindowManager.focus_window_by_title("PS0(1)/011 Display Actual Cost Line Items for Cost Centers")
    pyautogui.hotkey("ctrl","shift", "F7")
    time.sleep(3)
    WindowManager.focus_window_by_title("Save As")
    pyautogui.click(x=163, y=423)
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
    WindowManager.focus_window_by_title("Confirm Save As")
    pyautogui.press("left")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(5)

    try: 
        WindowManager.focus_window_by_title("SAP GUI Security")
        pyautogui.click(x=311, y=430)
        time.sleep(3)
    except Exception as e:
        print("Probably the window SAP GUI Security did not open")
    print("done")

    WindowManager.close_window_by_title(f"{excel_sheet_name} - Excel")
    time.sleep(2)
    WindowManager.close_window_by_title("PS0(1)/011 Display Actual Cost Line Items for Cost Centers")
    time.sleep(2)
    WindowManager.close_window_by_title("SAP Logon 800")
    print("Sap Closed")


if __name__ == "__main__":
    save_excel_sheet(excel_sheet_name="EXPORT_BAB_2024_consolidado_PPC_BI.XLSX", variant="BAB2024", code="s_alr_87013625")
    # s_alr_87013625   