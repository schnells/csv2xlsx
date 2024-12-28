import os
import pandas as pd
from tkinter import Tk, Label, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD


def convert_csv_to_excel(file_path):
    try:
        # CSV-Datei einlesen (UTF-8)
        df = pd.read_csv(file_path, encoding='utf-8', sep=';')

        # Zieldatei erstellen
        output_file = os.path.splitext(file_path)[0] + ".xlsx"

        # Speichern im Excel-Format
        df.to_excel(output_file, index=False, engine='openpyxl')

        # Erfolgsmeldung
        messagebox.showinfo("Erfolg", f"Datei wurde erfolgreich konvertiert:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Konvertieren:\n{str(e)}")


def on_drop(event):
    # Dateipfad bereinigen (kann geschweifte Klammern enthalten)
    file_path = event.data.strip('{}')
    if file_path.lower().endswith('.csv'):
        convert_csv_to_excel(file_path)
    else:
        messagebox.showerror("Fehler", "Bitte nur CSV-Dateien verwenden.")


# Hauptfenster erstellen
root = TkinterDnD.Tk()  # Drag-and-Drop-fähiges Fenster
root.title("CSV zu Excel Konverter")
root.geometry("500x200")

# Label für Drag-and-Drop-Bereich
label = Label(root, text="Ziehen Sie Ihre CSV-Datei hierher", padx=20, pady=20, relief="groove")
label.pack(expand=True, fill="both")

# Drag-and-Drop aktivieren
root.drop_target_register(DND_FILES)
label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', on_drop)

# GUI starten
root.mainloop()
