import os
import difflib
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def extract_text_from_docx(file_path):
    """
    Extrahiert den gesamten Text aus einem Word-Dokument (.docx).
    """
    document = Document(file_path)
    text = []
    for paragraph in document.paragraphs:
        text.append(paragraph.text)
    return "\n".join(text)

def calculate_similarity(file1, file2):
    """
    Berechnet die Ähnlichkeit zwischen zwei Dokumenten.
    """
    content1 = extract_text_from_docx(file1)
    content2 = extract_text_from_docx(file2)
    similarity = difflib.SequenceMatcher(None, content1, content2).ratio()
    return similarity * 100  # Prozentualer Wert

def compare_with_folder(main_file, folder_path):
    """
    Vergleicht ein Haupt-Dokument mit allen .docx-Dateien in einem Ordner.
    """
    results = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx') and filename != os.path.basename(main_file):
            other_file = os.path.join(folder_path, filename)
            similarity = calculate_similarity(main_file, other_file)
            results.append((filename, round(similarity, 2)))
    return results

def select_main_file():
    """
    Öffnet einen Dialog zum Auswählen der Hauptdatei.
    """
    filepath = filedialog.askopenfilename(
        title="Hauptdatei auswählen",
        filetypes=[("Word-Dokumente", "*.docx")])
    if filepath:
        main_file_entry.delete(0, tk.END)
        main_file_entry.insert(0, filepath)

def select_folder():
    """
    Öffnet einen Dialog zum Auswählen eines Ordners.
    """
    folderpath = filedialog.askdirectory(title="Ordner auswählen")
    if folderpath:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folderpath)

def start_comparison():
    """
    Startet den Vergleich und zeigt die Ergebnisse an.
    """
    main_file = main_file_entry.get()
    folder_path = folder_entry.get()

    if not os.path.isfile(main_file):
        messagebox.showerror("Fehler", "Die angegebene Hauptdatei existiert nicht.")
        return
    if not os.path.isdir(folder_path):
        messagebox.showerror("Fehler", "Der angegebene Ordner existiert nicht.")
        return

    results = compare_with_folder(main_file, folder_path)

    # Ergebnisse anzeigen
    result_text.delete(*result_text.get_children())  # Löscht alte Einträge
    for filename, similarity in results:
        # Tag auswählen basierend auf dem Ähnlichkeitswert
        if similarity < 30:
            tag = "low"
        elif 30 <= similarity < 60:
            tag = "medium"
        else:
            tag = "high"
        result_text.insert("", "end", values=(filename, f"{similarity}%"), tags=(tag,))

# GUI erstellen
root = tk.Tk()
root.title("RACompare")

# Hauptdatei-Auswahl
main_file_label = tk.Label(root, text="Hauptdatei:")
main_file_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
main_file_entry = tk.Entry(root, width=50)
main_file_entry.grid(row=0, column=1, padx=10, pady=5)
main_file_button = tk.Button(root, text="Durchsuchen", command=select_main_file)
main_file_button.grid(row=0, column=2, padx=10, pady=5)

# Ordner-Auswahl
folder_label = tk.Label(root, text="Ordner mit Regieanweisungen (.docx-Dateien):")
folder_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=1, column=1, padx=10, pady=5)
folder_button = tk.Button(root, text="Durchsuchen", command=select_folder)
folder_button.grid(row=1, column=2, padx=10, pady=5)

# Vergleich starten
compare_button = tk.Button(root, text="Vergleich starten", command=start_comparison)
compare_button.grid(row=2, column=0, columnspan=3, pady=10)

# Ergebnis-Tabelle
columns = ("Datei", "Ähnlichkeit")
result_text = ttk.Treeview(root, columns=columns, show="headings", height=10)
result_text.heading("Datei", text="Datei")
result_text.heading("Ähnlichkeit", text="Ähnlichkeit (%)")
result_text.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

# Tags für Farbgebung hinzufügen
result_text.tag_configure("low", background="red", foreground="white")
result_text.tag_configure("medium", background="yellow", foreground="black")
result_text.tag_configure("high", background="green", foreground="white")

# GUI starten
root.mainloop()