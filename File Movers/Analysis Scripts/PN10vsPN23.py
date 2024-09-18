from tkinter import *
from tkinter import filedialog
import tkinter as tk
from dash import Dash, html, dcc, callback, Output, Input
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import os
import sys
import base64
from PIL import Image

def browseFiles():
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select a File",
                                          filetypes = (("Text files",
                                                        "*.xls*"),
                                                       ("all files",
                                                        "*.*")))
      
    # Change label contents
    label_file_explorer.configure(text="File Opened: "+filename,wraplength=300, justify="center")
 
    #Analysis
    df = pd.read_excel(filename, sheet_name='ContinuousData')
    Testdata = df[['PlotTime','DilutePNCumulativeCount', 'DilutePNRate', 'DilutePNCumulativeCountSub23nm', 'DilutePNRateSub23nm']]
    #print(PNdata.head())
    image = 'MPTLogo.png'
    logo = base64.b64encode(open(image, 'rb').read())

    app = Dash()

    fig = make_subplots(rows=2, cols=1,shared_xaxes=True, subplot_titles=("PN Cumulative Count PN10 Vs PN23", "PN Rate PN10 Vs PN23"))

    fig.append_trace(go.Scatter(x=Testdata.PlotTime, y=Testdata.DilutePNCumulativeCount, name="PN PN23", line=dict(color="#0526FA")),1,1)
    fig.append_trace(go.Scatter(x=Testdata.PlotTime, y=Testdata.DilutePNCumulativeCountSub23nm, name="PN PN10", line=dict(color="#397A04")),1,1)

    fig.append_trace(go.Scatter(x=Testdata.PlotTime, y=Testdata.DilutePNRate, name="PN PN23", line=dict(color="#0526FA")),2,1)
    fig.append_trace(go.Scatter(x=Testdata.PlotTime, y=Testdata.DilutePNRateSub23nm, name="PN PN10" ,line=dict(color="#397A04")),2,1)

    fig.update_yaxes(showexponent='all', exponentformat = 'e') #comment out if not PN
    fig.layout.images = [dict(source='data:image/png;base64,{}'.format(logo.decode()),xref="paper", yref="paper",x=0.15, y=1.018,sizex=0.2, sizey=0.08,xanchor="right", yanchor="bottom")]

    fig.show()      
 
def exit_program():
    sys.exit(0)
                                                                                                      
# Create the root window
window = tk.Tk()
  
# Set window title
window.title('PN 10nm Vs PN 23nm')
  
# Set window size
window.geometry("500x125")
  
#Set window background color
window.config(background = "grey")
  
# Create a File Explorer label
label_file_explorer = Label(window, 
                            text = "PN 10nm Vs PN 23nm", 
                            fg = "blue",
                            height=5)
  
      
button_explore = Button(window, 
                        text = "Browse Files",
                        command = browseFiles) 
  
button_exit = Button(window, 
                     text = "Exit",
                     command = exit_program) 
  
label_file_explorer.pack(pady=(0,0), fill=BOTH) 
button_explore.pack(pady=(0,0), side=tk.LEFT) 
button_exit.pack(pady=0, side=tk.RIGHT) 

window.mainloop()