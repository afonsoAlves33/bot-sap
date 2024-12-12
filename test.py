import pyautogui
import pygetwindow as gw
import time

# Listar todas as janelas abertas
windows = gw.getAllTitles()
print("Janelas abertas:")
for i, title in enumerate(windows):
    if title.strip():  # Ignorar títulos vazios
        print(f"{i + 1}. {title}")

# Escolha a janela pelo título
janela_titulo = "Nome da Janela"
janela = gw.getWindowsWithTitle("SAP Logon 800")

if janela:
    # Ativar a janela
    janela[0].activate()
    time.sleep(1)  # Esperar a janela ativar
    # Exemplo de ação: mover o mouse
    pyautogui.moveTo(100, 100)
else:
    print(f"Janela com o título '{janela_titulo}' não encontrada.")