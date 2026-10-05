import subprocess
import time
from pathlib import Path
import pyautogui

EXE = Path(r"C:\Folha\folha.exe")
CRED = Path("data/credencial.txt")

def login():
    usuario, senha = CRED.read_text(encoding="utf-8").splitlines()[:2]
    
    subprocess.Popen([str(EXE)], cwd=EXE.parent)
    time.sleep(.5)

    pyautogui.write(usuario, interval=.05)
    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(senha, interval=.05)
    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.press("enter")
    time.sleep(.4)

    pyautogui.hotkey("alt")
    time.sleep(.02)

    pyautogui.hotkey("r")
    time.sleep(.02)

    pyautogui.press("down", presses=6, interval=.001)
    time.sleep(.02)

    pyautogui.press("right")
    time.sleep(.02)

    pyautogui.press("down", presses=9, interval=.001)
    time.sleep(.02)

    pyautogui.press("enter")
    time.sleep(.02)

    pyautogui.hotkey("win", "left")
    time.sleep(.02)

def main():
    login()

if __name__ == "__main__":
    main()