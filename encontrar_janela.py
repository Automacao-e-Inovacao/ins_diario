

def encontrar_janela():
    import win32gui

    def enum_windows_callback(hwnd, window_list):
        window_text = win32gui.GetWindowText(hwnd)
        if window_text.startswith("Iniciar sess"):
            window_list.append(window_text)

    def check_open_windows():
        window_list = []
        win32gui.EnumWindows(enum_windows_callback, window_list)
        return window_list

    # Exemplo de uso
    open_windows = check_open_windows()
    if open_windows:
        print("As seguintes janelas estão abertas:")
        for window in open_windows:
            print(window)
        return True
    else:
        return False
