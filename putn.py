import tkinter as tk
from tkinter import filedialog
from os import getcwd
from gui import *

root = tk.Tk()
root.wm_title("Pick Ups To Notes")
root.minsize(width=250, height=10)
root.iconbitmap(os.getcwd() + '\putn.ico')
app = Application(master=root)
app.mainloop()
