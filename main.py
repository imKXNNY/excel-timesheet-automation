# main.py
import tkinter as tk

from src.logger import setup_logging
setup_logging()
import logging
from src import ArbeitszeiterfassungGUI


def start_gui():
    try:
        root = tk.Tk()
        gui = ArbeitszeiterfassungGUI(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"Fehler beim Starten der GUI: {e}")

if __name__ == '__main__':
    start_gui()
