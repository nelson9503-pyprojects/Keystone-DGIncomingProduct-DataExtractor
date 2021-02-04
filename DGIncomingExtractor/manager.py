import tkinter as tk
from tkinter import ttk
from tkinter import font
from .file_selector import FileSelector
from .column_selector import ColumnSelector
from .extractor import Extractor


class Manager:

    def __init__(self):
        self.win = tk.Tk()
        self.win.title("Keystone DG Incoming Production Extractor")
        defaultFontObj = font.nametofont("TkDefaultFont")
        defaultFontObj.config(size=18)

        self.file_selector = FileSelector(self.win)
        self.file_selector.main.pack(fill="both", pady=30)

        self.column_selector = ColumnSelector(self.win)
        self.column_selector.main.pack(fill="both", pady=30)

        self.btn_extract = ttk.Button(
            self.win, text="Extract", command=self.extract_dg_product)
        self.btn_extract.pack(pady=30)

    def extract_dg_product(self):
        self.btn_extract.config(state="disable")
        fileconfig = self.file_selector.get_config()
        colsconfig = self.column_selector.get_config()
        Extractor(fileconfig, colsconfig)
        self.btn_extract.config(state="normal")

    def run(self):
        self.win.mainloop()
