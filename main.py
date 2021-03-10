import datetime
import json
import os
import threading
import webbrowser

import colorama
import openpyxl


colorama.init()


def main():
    while A == 0:
        k = giorno()
        materia = ora(k)
        meet(materia, k)


def wait():
    i = 0
    while True:
        print(colorama.Fore.WHITE + animazione[i % len(animazione)], end="\r")
        i += 1
        timeout5 = threading.Event()
        timeout5.wait(timeout=0.4)


def giorno():
    for k in range(
        len(lista_giorni)
    ):  # scorro la lista dei giorni per trovare quello corrente
        if str(lista_giorni[k]) == datetime.datetime.now().strftime(
            "%A"
        ):  # giorno trovato = giorno corrente
            return k


def ora(k):
    for i in range(
        len(lista_ore)
    ):  # scorro la lista delle ore per trovare quella giusta
        if str(lista_ore[i]) == datetime.datetime.now().strftime(
            "%H:%M:%S"
        ):  # ora nel file = ora corrente
            x = int(
                lista_caselle[i] + 1
            )  # numero della riga (+1 perchè la prima è vuota)
            y = (
                k + 2
            )  # numero della colonna (+2 perchè la lista parte da 0 e il file
            # excel ha la colonna delle ore da non contare)
            materia = foglio1.cell(
                row=x, column=y
            ).value  # prendo la materia corrispondente all'ora e al giorno
            return materia


def meet(materia, k):
    try:
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + "\n\nMeet found")
        print(colorama.Fore.WHITE + "Connecting to: {}".format(materia))
        print(colorama.Fore.WHITE)
        webbrowser.open(
            dict_classi[materia], autoraise=True
        )  # apro il link alla classe di meet corrispondente all'ora e al
        # giorno
        timeout = threading.Event()
        timeout.wait(timeout=2)  # Aspetto 2 secondi
    except KeyError:
        ora(k)


# Cambio directory x non specificare il percorso dei file
os.chdir(os.getcwd())

# File JSON con link
classi = open("meet-link.json", "r").read()
dict_classi = json.loads(classi)

# File Excel con orario
orario = openpyxl.load_workbook("schedule.xlsx")
foglio1 = orario.get_sheet_by_name("Schedule")

# Giorni
lun = foglio1.cell(row=1, column=2).value
mar = foglio1.cell(row=1, column=3).value
mer = foglio1.cell(row=1, column=4).value
gio = foglio1.cell(row=1, column=5).value
ven = foglio1.cell(row=1, column=6).value

# Ore
ora1 = foglio1.cell(row=2, column=1).value
ora2 = foglio1.cell(row=3, column=1).value
ora3 = foglio1.cell(row=4, column=1).value
ora4 = foglio1.cell(row=5, column=1).value
ora5 = foglio1.cell(row=6, column=1).value
ora6 = foglio1.cell(row=7, column=1).value
ora7 = foglio1.cell(row=8, column=1).value
ora8 = foglio1.cell(row=9, column=1).value
ora9 = foglio1.cell(row=10, column=1).value

# Liste con ore, giorni e celle
lista_ore = [ora1, ora2, ora3, ora4, ora5, ora6, ora7, ora8, ora9]
lista_caselle = [1, 2, 3, 4, 5, 6, 7, 8, 9]
lista_giorni = [lun, mar, mer, gio, ven]

# Animazione
animazione = [
    "[°   ]",
    "[ .  ]",
    "[  ° ]",
    "[   .]",
    "[  ° ]",
    "[ .  ]",
]

print(
    colorama.Fore.YELLOW
    + colorama.Style.BRIGHT
    + "\n===== Meet attender =====\n"
)
print(
    colorama.Fore.WHITE
    + "The program is running.\nThe Meet will open at the indicated time.\n"
)

A = 0
i = 0

f = threading.Thread(target=wait, daemon=True)
f.start()

e = threading.Thread(target=main)
e.start()
