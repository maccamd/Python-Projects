import os, time
import pandas as pd

def analysis(filename):
    cols = [0,2,3,4,8,9,10,11,12,13,14,15,16,17,18]
    rows = [15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33]
    df = pd.read_excel(filename, 'Result Force', engine='openpyxl')
    df.drop(df.columns[cols], axis=1, inplace=True)
    df.drop(df.index[rows], inplace=True)
    df.rename(columns={'Unnamed: 1':'A', 'Unnamed: 5':'B', 'Unnamed: 6':'C', 'Unnamed: 7':'D'}, inplace=True )
    date = df.at[2, 'B']
    setpoint = df.at[13,'B']
    result = df.at[14, 'B']
    maxlimit = df.at[37,'B']
    testError = df.at[39,'B']

    return

path_to_watch = r"W:\EDC Schedules\RDEC Weekly Checks Results\VDC 2\FEDERAL CERT"
before = dict ([(f, None) for f in os.listdir (path_to_watch)])
while 1:
  time.sleep (10)
  after = dict ([(f, None) for f in os.listdir (path_to_watch)])
  added = [f for f in after if not f in before]
  removed = [f for f in before if not f in after]
  if added: print("Added: ", ", ".join (added))
  if removed: print("Removed: ", ", ".join (removed))
  filename = path_to_watch + added
  analysis(filename)
  before = after