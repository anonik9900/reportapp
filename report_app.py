import tkinter as tk
from tkinter import ttk, messagebox
import csv
import os
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime
import ttkbootstrap as ttkb  # Importiamo la libreria ttkbootstrap

# Funzione per creare una nuova pagina PDF con i dati del report
def crea_nuova_pagina_pdf(dati, data_corrente):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Titolo del nuovo report
    pdf.cell(200, 10, txt="Nuovo Report Lavoro", ln=True, align="C")
    pdf.cell(200, 10, txt=f"Data di generazione: {data_corrente}", ln=True, align="C")
    pdf.ln(10)

    # Aggiungere i nuovi dati con il nome della risorsa (se presente)
    for chiave, valore in dati.items():
        if valore:  # Mostra solo se c'è un valore
            pdf.cell(200, 10, txt=f"{chiave}: {valore}", ln=True, align="L")
    
    # Salvare temporaneamente il PDF
    nuova_pagina_pdf = "temp_nuova_pagina.pdf"
    pdf.output(nuova_pagina_pdf)
    
    return nuova_pagina_pdf

# Funzione per unire il nuovo PDF con il PDF esistente (se presente)
def aggiorna_pdf(report_pdf, nuova_pagina_pdf):
    writer = PdfWriter()

    # Se il PDF esistente esiste, aggiungere le sue pagine al writer
    if os.path.exists(report_pdf):
        reader = PdfReader(report_pdf)
        for pagina in range(len(reader.pages)):
            writer.add_page(reader.pages[pagina])

    # Aggiungere la nuova pagina generata
    nuovo_pdf_reader = PdfReader(nuova_pagina_pdf)
    writer.add_page(nuovo_pdf_reader.pages[0])

    # Scrivere il PDF aggiornato
    with open(report_pdf, "wb") as file_output:
        writer.write(file_output)

    # Rimuovere il PDF temporaneo
    os.remove(nuova_pagina_pdf)

# Funzione per generare il report in CSV e PDF
def genera_report():
    # Raccogliere i dati inseriti
    dati = {
        'Nome Risorsa': nome_risorsa_entry.get(),
        'Numero SIM fatte': sim_entry.get(),
        'Numero Fissi fatti': fissi_entry.get(),
        'Numero MNP fatte': mnp_entry.get(),
        'Numero Rateizzazioni fatte': rateizzazioni_entry.get(),
        'Numero Intrattenimento fatti': intrattenimento_entry.get(),
        'Numero FWA ricaricabili fatte': fwa_entry.get(),
    }
    
    # Nome del file report CSV e PDF
    report_csv = "report_lavoro.csv"
    report_pdf = "report_lavoro.pdf"
    
    # Data corrente
    data_corrente = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Scrivere i dati in un file CSV
    with open(report_csv, mode='a', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=dati.keys())
        if os.stat(report_csv).st_size == 0:  # Se il file è vuoto, aggiungi l'intestazione
            writer.writeheader()
        writer.writerow(dati)
    
    # Creare una nuova pagina PDF con i nuovi dati
    nuova_pagina_pdf = crea_nuova_pagina_pdf(dati, data_corrente)
    
    # Unire il nuovo PDF al PDF esistente
    aggiorna_pdf(report_pdf, nuova_pagina_pdf)

    # Mostrare un messaggio di successo
    messagebox.showinfo("Successo", f"Report aggiornato con successo in {os.path.abspath(report_csv)} e {os.path.abspath(report_pdf)}")

# Creazione della finestra principale con tema moderno "superhero"
root = ttkb.Window(themename="superhero")
root.title("Genera Report Lavoro")

# Imposta la finestra a schermo intero
root.state('zoomed')

import sys
import os

# ...

# Imposta l'icona dell'applicazione
if getattr(sys, 'frozen', False):
    # Se l'app è in esecuzione come eseguibile
    icon_path = os.path.join(sys._MEIPASS, 'app_icon.ico')
else:
    # Se è in esecuzione come script
    icon_path = 'app_icon.ico'

root.iconbitmap(icon_path)

# Imposta l'icona dell'applicazione (sostituisci 'app_icon.ico' con il percorso del tuo file .ico)
#root.iconbitmap('app_icon.ico')

# Layout generale della finestra
root.configure(bg='#2b3e50')  # Impostiamo uno sfondo scuro per un look moderno

# Stile personalizzato per ttk
style = ttkb.Style()

# Frame superiore per i campi di input
frame_campi = ttkb.Frame(root, padding="20 10 20 10", bootstyle="dark")
frame_campi.grid(row=0, column=0, padx=20, pady=20, sticky=tk.N+tk.E+tk.W)

# Rendiamo il layout adattabile alla finestra
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)

# Etichette e campi di inserimento personalizzati
def aggiungi_campo(testo, row):
    ttkb.Label(frame_campi, text=testo, bootstyle="info").grid(row=row, column=0, padx=10, pady=5, sticky=tk.W)
    entry = ttkb.Entry(frame_campi, width=25)
    entry.grid(row=row, column=1, padx=10, pady=5, sticky=tk.E+tk.W)
    return entry

# Campo per il nome della risorsa
nome_risorsa_entry = aggiungi_campo("Nome Risorsa:", 0)

# Campi di input per i vari dati
sim_entry = aggiungi_campo("Numero SIM fatte:", 1)
fissi_entry = aggiungi_campo("Numero Fissi fatti:", 2)
mnp_entry = aggiungi_campo("Numero MNP fatte:", 3)
rateizzazioni_entry = aggiungi_campo("Numero Rateizzazioni fatte:", 4)
intrattenimento_entry = aggiungi_campo("Numero Intrattenimento fatti:", 5)
fwa_entry = aggiungi_campo("Numero FWA ricaricabili fatte:", 6)

# Bottone per generare il report con effetto hover
genera_button = ttkb.Button(root, text="Genera Report", command=genera_report, bootstyle="success outline")
genera_button.grid(row=1, column=0, pady=20, sticky=tk.E+tk.W)

# Footer con istruzioni o copyright
footer = ttkb.Label(root, text="Applicazione per il report del lavoro Nicholas Impieri © 2024", anchor=tk.CENTER, bootstyle="light")
footer.grid(row=2, column=0, pady=10, sticky=tk.E+tk.W)

# Avviare il loop principale
root.mainloop()
