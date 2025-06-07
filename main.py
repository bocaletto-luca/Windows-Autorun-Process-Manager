# Software Name: Windows-Autorun-Process-Manager
# Author: Bocaletto Luca
# Language: Italian
# Import delle librerie tkinter e altre librerie necessarie

import tkinter as tk
import tkinter.ttk as ttk  # Importa il modulo ttk da tkinter
import os
import subprocess
import ctypes  # Importa ctypes per la gestione dei messaggi di errore
import winreg  # Importa winreg per la manipolazione del registro di sistema di Windows
from tkinter import filedialog  # Importa il modulo filedialog da tkinter per la finestra di dialogo di selezione file

# Funzione per ottenere le applicazioni all'avvio di Windows
def get_startup_apps():
    startup_apps = []
    # Apre la chiave del registro di sistema HKEY_CURRENT_USER per leggere le applicazioni di avvio
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ) as key:
        try:
            index = 0
            while True:
                # Enumera i valori nella chiave del registro (nome e percorso dell'applicazione)
                name, value, _ = winreg.EnumValue(key, index)
                startup_apps.append((name, value))
                index += 1
        except WindowsError:
            pass
    return startup_apps

# Funzione per aggiungere un'applicazione all'avvio di Windows
def aggiungi_startup_app():
    app_path = app_path_entry.get()  # Ottiene il percorso dell'applicazione dall'input dell'utente
    app_name = os.path.basename(app_path)  # Estrae il nome dell'applicazione dal percorso
    # Apre la chiave del registro di sistema per scrivere il nuovo valore (nome e percorso dell'applicazione)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE) as key:
        winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path)
    refresh_list()  # Aggiorna la lista delle applicazioni all'avvio

# Funzione per selezionare il percorso dell'applicazione utilizzando il pulsante "Sfoglia"
def seleziona_percorso():
    file_path = filedialog.askopenfilename()  # Apre una finestra di dialogo per la selezione del file
    app_path_entry.delete(0, tk.END)  # Cancella il contenuto dell'entry widget
    app_path_entry.insert(0, file_path)  # Inserisce il percorso del file nell'entry widget

# Funzione per rimuovere un'applicazione dall'avvio di Windows
def rimuovi_startup_app():
    selected_item = startup_apps_list.selection()  # Ottiene l'elemento selezionato nella lista
    if selected_item:
        app_name = startup_apps_list.item(selected_item, "values")[0]  # Ottiene il nome dell'applicazione dalla lista
        # Apre la chiave del registro di sistema per rimuovere il valore dell'applicazione dall'avvio
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE) as key:
            winreg.DeleteValue(key, app_name)
        refresh_list()  # Aggiorna la lista delle applicazioni all'avvio

# Funzione per avviare un'applicazione
def avvia_applicazione():
    selected_item = startup_apps_list.selection()  # Ottiene l'elemento selezionato nella lista
    if selected_item:
        app_path = startup_apps_list.item(selected_item, "values")[1]  # Ottiene il percorso dell'applicazione dalla lista
        try:
            subprocess.Popen([app_path], shell=True)  # Avvia l'applicazione tramite subprocess
        except Exception as e:
            # Mostra un messaggio di errore in caso di errore nell'avvio dell'applicazione
            ctypes.windll.user32.MessageBoxW(0, f"Impossibile avviare l'applicazione.\nErrore: {str(e)}", "Errore", 0)

# Funzione per aggiornare la lista delle applicazioni all'avvio
def refresh_list():
    startup_apps_list.delete(*startup_apps_list.get_children())  # Cancella tutti gli elementi nella lista attuale
    for app_name, app_path in get_startup_apps():
        # Aggiunge le applicazioni all'avvio attuali nella lista
        startup_apps_list.insert("", "end", values=(app_name, app_path))

# Creazione della finestra principale
root = tk.Tk()  # Crea una nuova finestra tkinter
root.title("Gestione Applicazioni all'Avvio di Windows")  # Imposta il titolo della finestra

# Etichetta e campo di inserimento per il percorso dell'applicazione
app_path_label = tk.Label(root, text="Percorso dell'Applicazione:")  # Crea un'etichetta
app_path_label.pack()  # Visualizza l'etichetta nella finestra
app_path_entry = tk.Entry(root)  # Crea un campo di inserimento per il percorso
app_path_entry.pack()  # Visualizza il campo di inserimento nella finestra

# Pulsante "Sfoglia" per selezionare il percorso dell'applicazione
sfoglia_button = tk.Button(root, text="Sfoglia", command=seleziona_percorso)  # Crea un pulsante
sfoglia_button.pack()  # Visualizza il pulsante nella finestra

# Pulsanti per aggiungere, rimuovere e avviare applicazioni
aggiungi_button = tk.Button(root, text="Aggiungi all'Avvio", command=aggiungi_startup_app)  # Crea un pulsante per l'aggiunta
rimuovi_button = tk.Button(root, text="Rimuovi dall'Avvio", command=rimuovi_startup_app)  # Crea un pulsante per la rimozione
avvia_button = tk.Button(root, text="Avvia Applicazione", command=avvia_applicazione)  # Crea un pulsante per l'avvio
aggiungi_button.pack()  # Visualizza il pulsante "Aggiungi" nella finestra
rimuovi_button.pack()  # Visualizza il pulsante "Rimuovi" nella finestra
avvia_button.pack()  # Visualizza il pulsante "Avvia" nella finestra

# Lista delle applicazioni all'avvio
startup_apps_list = ttk.Treeview(root, columns=("Nome", "Percorso"), show="headings")  # Crea una lista a due colonne
startup_apps_list.heading("Nome", text="Nome")  # Imposta l'intestazione della colonna "Nome"
startup_apps_list.heading("Percorso", text="Percorso")  # Imposta l'intestazione della colonna "Percorso"
startup_apps_list.pack()  # Visualizza la lista nella finestra
refresh_list()  # Aggiorna la lista delle applicazioni all'avvio

# Esecuzione del loop principale della GUI
root.mainloop()  # Avvia il ciclo principale per l'interfaccia grafica
