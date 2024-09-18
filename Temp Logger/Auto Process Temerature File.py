# read temp logger .data file
import tkinter as tk
import matplotlib as pt
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
tempearture_log = open(file_path)
tempearture_log.read
print(tempearture_log)