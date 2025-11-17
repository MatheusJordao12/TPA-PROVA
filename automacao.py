import pyautogui
import time
import psutil
import keyboard

tipo = input("Qual o tipo do cadastro? (excel ou web) ")

def parar():
    if keyboard.is_pressed("insert"):
        print("\nPrograma interrompido")
        exit()

if tipo.lower() == "excel":
    excel = "excel.exe"
    processos = []

    for proc in psutil.process_iter(['name']):
        nome = proc.info['name']
        processos.append(nome)
        parar()

    if excel in processos:
        pyautogui.hotkey('alt', 'tab')
        pyautogui.hotkey('enter')
    else: 
        time.sleep(1)
        parar()
        pyautogui.hotkey('win')
        time.sleep(1)
        pyautogui.write('excel')
        pyautogui.hotkey('enter')
        time.sleep(3)
        pyautogui.hotkey('enter')

    time.sleep(1)
    parar()
    pyautogui.write('Produto')
    pyautogui.hotkey('tab')
    pyautogui.write('Preco')
    pyautogui.hotkey('tab')
    pyautogui.write('Identificacao')
    pyautogui.hotkey('tab')
    pyautogui.write('Categoria')
    pyautogui.hotkey('enter')

    with open("dados.txt", "r", encoding="utf-8") as arquivo:
        linhas = arquivo.readlines()

    for linha in linhas:
        parar()
        partes = linha.strip().split(";")

        for parte in partes:
            parar()
            parte = parte.capitalize()
            pyautogui.write(parte)
            pyautogui.hotkey('tab')

        pyautogui.hotkey('enter')

    time.sleep(1)
    parar()
    pyautogui.hotkey('ctrl', 't')
    pyautogui.hotkey('alt', 'o')
    pyautogui.hotkey('alt', 'c')
    pyautogui.write('ac')
    pyautogui.hotkey('alt', 'c')
    pyautogui.write('o')
    pyautogui.write('t')

else:
    pyautogui.hotkey('alt', 'tab')

    with open("dados.txt", "r", encoding="utf-8") as arquivo:
        linhas = arquivo.readlines()

    for linha in linhas:
        parar()
        pyautogui.hotkey('enter')
        partes = linha.strip().split(";")
        pyautogui.hotkey('tab')

        for parte in partes:
            parar()
            parte = parte.capitalize()
            pyautogui.write(parte)
            pyautogui.hotkey('tab')

        pyautogui.hotkey('enter')
