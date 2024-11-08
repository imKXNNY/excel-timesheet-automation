# src/arbeitszeiterfassung_gui.py
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
import tkinter.ttk as ttk

from src.logger import setup_logging
setup_logging()
import logging
from src import adjust_hours, analyze_adjusted_data, validate_data, import_employee_data

class ArbeitszeiterfassungGUI:
    def __init__(self, master):
        self.master = master
        master.title("Arbeitszeiterfassung System")
        master.geometry("400x500")  # Erweitert für Progressbar

        self.label = tk.Label(master, text="Arbeitszeiterfassung System", font=("Helvetica", 16))
        self.label.pack(pady=10)

        self.select_file_button = tk.Button(master, text="Zeiterfassungsdatei auswählen", command=self.select_file, width=30)
        self.select_file_button.pack(pady=5)

        self.process_button = tk.Button(master, text="Daten verarbeiten", command=self.process_data, state=tk.DISABLED, width=30)
        self.process_button.pack(pady=5)

        self.validate_button = tk.Button(master, text="Daten validieren", command=self.validate_data, state=tk.DISABLED, width=30)
        self.validate_button.pack(pady=5)

        self.analyze_button = tk.Button(master, text="Daten analysieren", command=self.analyze_data, state=tk.DISABLED, width=30)
        self.analyze_button.pack(pady=5)

        self.import_button = tk.Button(master, text="Mitarbeiterdaten importieren", command=self.import_data, state=tk.NORMAL, width=30)
        self.import_button.pack(pady=5)

        self.quit_button = tk.Button(master, text="Beenden", command=master.quit, width=30)
        self.quit_button.pack(pady=20)

        # Progressbar hinzufügen
        self.progress = ttk.Progressbar(master, orient='horizontal', length=300, mode='determinate')
        self.progress.pack(pady=10)

    def select_file(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if self.filename:
            self.process_button['state'] = tk.NORMAL
            messagebox.showinfo("Datei ausgewählt", f"Ausgewählte Datei: {self.filename}")
            logging.info(f"Datei ausgewählt: {self.filename}")

    def process_data(self):
        if hasattr(self, 'filename'):
            try:
                output_filename = self.filename.replace('.xlsx', '_angepasst.xlsx')
                self.progress['value'] = 20
                self.master.update_idletasks()

                adjust_hours(self.filename, output_filename)
                self.progress['value'] = 60
                self.master.update_idletasks()

                self.validate_button['state'] = tk.NORMAL
                self.analyze_button['state'] = tk.NORMAL
                messagebox.showinfo("Verarbeitung abgeschlossen", f"Angepasste Daten wurden in {output_filename} gespeichert.")
                logging.info(f"Daten verarbeitet und gespeichert als: {output_filename}")
                self.processed_file = output_filename
                self.progress['value'] = 100
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler bei der Datenverarbeitung: {e}")
                logging.error(f"Fehler bei der Datenverarbeitung: {e}")
                self.progress['value'] = 0
        else:
            messagebox.showerror("Fehler", "Bitte wählen Sie zuerst eine Datei aus.")

    def validate_data(self):
        if hasattr(self, 'processed_file'):
            try:
                validate_data(self.processed_file)
                messagebox.showinfo("Validierung abgeschlossen", "Datenvalidierung erfolgreich abgeschlossen.")
                logging.info("Datenvalidierung erfolgreich abgeschlossen.")
            except Exception as e:
                messagebox.showerror("Validierungsfehler", f"Datenvalidierung fehlgeschlagen: {e}")
                logging.error(f"Datenvalidierung fehlgeschlagen: {e}")
        else:
            messagebox.showerror("Fehler", "Bitte verarbeiten Sie zuerst die Daten.")

    def analyze_data(self):
        if hasattr(self, 'processed_file'):
            try:
                analysis_result = analyze_adjusted_data(self.processed_file)
                self.show_analysis(analysis_result)
                logging.info("Datenanalyse durchgeführt.")
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler bei der Datenanalyse: {e}")
                logging.error(f"Fehler bei der Datenanalyse: {e}")
        else:
            messagebox.showerror("Fehler", "Bitte verarbeiten Sie zuerst die Daten.")

    def show_analysis(self, analysis_result):
        analysis_window = tk.Toplevel(self.master)
        analysis_window.title("Datenanalyse")
        analysis_window.geometry("600x400")

        text_widget = tk.Text(analysis_window, wrap='word')
        text_widget.pack(expand=True, fill='both')
        text_widget.insert(tk.END, analysis_result)

    def import_data(self):
        try:
            admin_file = filedialog.askopenfilename(title="Admin Dashboard auswählen", filetypes=[("Excel Files", "*.xlsx;*.xls")])
            if not admin_file:
                messagebox.showerror("Fehler", "Bitte wählen Sie eine Admin Dashboard Datei aus.")
                return

            employee_files_dir = filedialog.askdirectory(title="Verzeichnis der Mitarbeiter-Dateien auswählen")
            if not employee_files_dir:
                messagebox.showerror("Fehler", "Bitte wählen Sie ein Verzeichnis aus.")
                return

            import_employee_data(admin_file, employee_files_dir)
            messagebox.showinfo("Import abgeschlossen", f"Admin-Dashboard wurde aktualisiert: {admin_file}")
            logging.info(f"Mitarbeiterdaten aus {employee_files_dir} in {admin_file} importiert.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Importieren der Daten: {e}")
            logging.error(f"Fehler beim Importieren der Daten: {e}")
