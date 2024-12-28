import os
import pandas as pd
from tkinter import Tk, Label, messagebox, Checkbutton, IntVar
from tkinterdnd2 import DND_FILES, TkinterDnD
from openpyxl import load_workbook


def apply_auto_width(file_path):
    """Wendet automatische Spaltenbreite auf die Excel-Datei an."""
    wb = load_workbook(file_path)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Spaltenbuchstabe
        for cell in col:
            try:  # Maximale Zeichenlänge in der Spalte finden
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2  # Extra Platz hinzufügen

    wb.save(file_path)


def convert_csv_to_excel(file_path):
    try:
        # CSV-Datei einlesen (mit Semikolon als Trennzeichen)
        df = pd.read_csv(file_path, encoding='utf-8', sep=';')

        # Excel-Datei speichern
        output_file = os.path.splitext(file_path)[0] + ".xlsx"
        df.to_excel(output_file, index=False, engine='openpyxl')

        # Falls Checkbox aktiv ist, automatische Spaltenbreite anwenden
        if auto_width_var.get():
            apply_auto_width(output_file)

        # Erfolgsmeldung
        messagebox.showinfo("Erfolg", f"Datei wurde erfolgreich konvertiert:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Konvertieren:\n{str(e)}")


def on_drop(event):
    # Dateipfad bereinigen
    file_path = event.data.strip('{}')
    if file_path.lower().endswith('.csv'):
        convert_csv_to_excel(file_path)
    else:
        messagebox.showerror("Fehler", "Bitte nur CSV-Dateien verwenden.")


# GUI erstellen
root = TkinterDnD.Tk()
root.title("CSV zu Excel Konverter")
root.geometry("500x200")

# Label für Drag-and-Drop
label = Label(root, text="Ziehen Sie Ihre CSV-Datei hierher", padx=20, pady=20, relief="groove")
label.pack(expand=True, fill="both")

# Checkbox für automatische Spaltenbreite (Standard: Aktiviert)
auto_width_var = IntVar(value=1)  # Standardmäßig aktiviert
Checkbutton(root, text="Automatische Spaltenbreite", variable=auto_width_var).pack(anchor='w', padx=20)

# Drag-and-Drop aktivieren
root.drop_target_register(DND_FILES)
label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', on_drop)

# GUI starten
root.mainloop()
