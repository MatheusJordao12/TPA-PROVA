import pyautogui
import time
import psutil
import keyboard

tipo = input("Qual o tipo do cadastro? (excel ou web) ")

# Define uma função chamada parar(), que interrompe o programa quando aperta o botão Insert (Ins)
def parar():
    if keyboard.is_pressed("insert"):
        print("\nPrograma interrompido")
        exit()
# Cada chamada de parar() verifica se o Insert foi apertado

# Se a escolha for "excel"
if tipo.lower() == "excel":
    excel = "excel.exe"
    processos = []

    # Procura nos processos se o Excel está aberto
    for proc in psutil.process_iter(['name']):
        nome = proc.info['name']
        processos.append(nome)
        parar()

    # Se o Excel estiver aberto, então vai para ele
    if excel in processos:
        pyautogui.hotkey('alt', 'tab')
        pyautogui.hotkey('enter')

    # Se não estiver, abre o Excel
    else:
        time.sleep(1)
        parar()
        pyautogui.hotkey('win')
        time.sleep(1)
        pyautogui.write('excel')
        pyautogui.hotkey('enter')
        time.sleep(3)
        pyautogui.hotkey('enter')

    # Depois disso, começa a digitar a tabela
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

    # Lê o arquivo dados.txt
    with open("dados.txt", "r", encoding="utf-8") as arquivo:  # "r" significa modo leitura com o enconding em utf-8 como arquivo
        linhas = arquivo.readlines()  # Lê todas as linhas como uma lista

    for linha in linhas:
        parar()
        partes = linha.strip().split(";")  # strip() remove espaços e \n ; split divide pelo ponto e vírgula fazendo uma lista

        for parte in partes:
            parar()
            parte = parte.capitalize()  # Deixa a primeira letra maiúscula
            pyautogui.write(parte)     # Digita o texto
            pyautogui.hotkey('tab')    # Vai para a próxima célula

        pyautogui.hotkey('enter')  # Vai para a próxima linha

    # Ajusta o texto no Excel
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
    # Caso não for excel, vai para o modo web
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
