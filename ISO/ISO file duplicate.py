import pandas as pd
import openpyxl as xl
import tkinter as tk
from tkinter import Image, filedialog
from tkinter import messagebox


root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename() #tkinter method
path = open(file_path, "rb") #rb is reading in binary
print(file_path)
