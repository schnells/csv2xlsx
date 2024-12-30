import os
import pandas as pd
from tkinter import Tk, Label, messagebox, Checkbutton, IntVar, Frame
from tkinterdnd2 import DND_FILES, TkinterDnD
from openpyxl import load_workbook
# from https://stackoverflow.com/a/70529292/6615718
from PyInstaller.utils.hooks import collect_data_files
datas = collect_data_files('tkinterdnd2')

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


def open_file(file_path):
    """Öffnet die Datei mit dem Standardprogramm."""
    try:
        if os.name == 'nt':  # Windows
            os.startfile(file_path)
        elif os.name == 'posix':  # macOS / Linux
            os.system(f'open "{file_path}"')
    except Exception as e:
        messagebox.showerror("Fehler", f"Datei konnte nicht geöffnet werden:\n{str(e)}")


def convert_csv_to_excel(file_path):
    try:
        # CSV-Datei einlesen (Semikolon als Trennzeichen)
        df = pd.read_csv(file_path, encoding='utf-8', sep=';')

        # Excel-Datei speichern
        output_file = os.path.splitext(file_path)[0] + ".xlsx"
        df.to_excel(output_file, index=False, engine='openpyxl')

        # Automatische Spaltenbreite anwenden, falls Checkbox aktiv ist
        if auto_width_var.get():
            apply_auto_width(output_file)

        # Erfolgsnachricht im Label anzeigen
        success_label.config(text="Datei erfolgreich konvertiert")
        open_link.config(text="öffnen", fg="blue")
        open_link.bind("<Button-1>", lambda e: open_file(output_file))

        # Falls Checkbox für Löschen aktiv ist, lösche die CSV-Datei
        if delete_csv_var.get():
            try:
                os.remove(file_path)
            except Exception as e:
                messagebox.showerror("Fehler", f"CSV-Datei konnte nicht gelöscht werden:\n{str(e)}")

    except Exception as e:
        # Fehler-MessageBox anzeigen
        messagebox.showerror("Fehler", f"Fehler beim Konvertieren:\n{str(e)}")


def on_drop(event):
    # Vorherige Meldungen und Links zurücksetzen
    success_label.config(text="")
    open_link.config(text="", fg="black")
    open_link.unbind("<Button-1>")

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

# Frame für die Checkboxen (nebeneinander)
checkbox_frame = Frame(root)
checkbox_frame.pack(pady=10)

# Checkbox für automatische Spaltenbreite (Standard: Aktiviert)
auto_width_var = IntVar(value=1)  # Standardmäßig aktiviert
Checkbutton(checkbox_frame, text="Automatische Spaltenbreite", variable=auto_width_var).pack(side='left', padx=10)

# Checkbox für Löschen der CSV-Datei (Standard: Deaktiviert)
delete_csv_var = IntVar(value=0)  # Standardmäßig deaktiviert
Checkbutton(checkbox_frame, text="CSV nach Konvertierung löschen", variable=delete_csv_var).pack(side='left', padx=10)

# Erfolgsnachricht und Link zum Öffnen
success_label = Label(root, text="", fg="black")
success_label.pack(pady=5)

open_link = Label(root, text="", fg="blue")
open_link.pack()

# Drag-and-Drop aktivieren
root.drop_target_register(DND_FILES)
label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', on_drop)

# GUI starten
root.mainloop()
