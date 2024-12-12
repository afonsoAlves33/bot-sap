import pyautogui
import pygetwindow as gw
import time

class WindowManager:
    @staticmethod
    def list_open_windows():
        """Lista todas as janelas abertas com seus títulos."""
        windows = gw.getAllTitles()
        return [title for title in windows if title.strip()]

    @staticmethod
    def focus_window_by_title(title):
        """Foca em uma janela específica pelo título.
        Args:
            title (str): Título da janela a ser focada.
        Returns:
            bool: True se a janela foi encontrada e ativada, False caso contrário.
        """
        janela = gw.getWindowsWithTitle(title)
        if janela:
            janela[0].activate()
            time.sleep(1)  # Esperar a janela ativar
            return True
        return False

    @staticmethod
    def move_mouse(x, y):
        """Move o mouse para a posição especificada.
        Args:
            x (int): Coordenada x na tela.
            y (int): Coordenada y na tela.
        """
        pyautogui.moveTo(x, y)
        
    @staticmethod
    def maximize_window_by_title(title):
        """Maximiza uma janela específica pelo título.

        Args:
            title (str): Título da janela a ser maximizada.

        Returns:
            bool: True se a janela foi encontrada e maximizada, False caso contrário.
        """
        janela = gw.getWindowsWithTitle(title)
        if janela:
            janela[0].maximize()
            return True
        return False