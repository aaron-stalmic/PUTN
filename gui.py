import tkinter as tk
from tkinter import filedialog
import pyodbc
from dbfunctions import *


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()
        self.getfile()

    def create_widgets(self):
        self.filename = ""
        self.picktoroute = tk.Button(self, width=20)
        self.picktoroute["text"] = "Pick List → Route Notes"
        self.picktoroute["command"] = lambda: createpick(self.filename)
        self.openfile = tk.Button(self, width=20)
        self.openfile["text"] = "Select Excel File. . ."
        self.openfile["command"] = self.getfile
        self.openfile.grid(row=0, column=0)
        self.picktoroute.grid(row=0, column=1)

    def getfile(self):
        self.filename = filedialog.askopenfilename(
            filetypes=[('Excel Spreadsheets', '.xls, .xlsx'),
                       ('All Files', '.*')],
            title="Please select the Pick Up List spreadsheet."
        )
