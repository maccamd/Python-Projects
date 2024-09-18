import pandas as pd

path = r'C:\\Users\\m0082668\\Documents\\Python Projects\\File Movers\\TEST-002.xlsx'
cols = [0,2,3,4,8,9,10,11,12,13,14,15,16,17,18]
rows = [15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33]
df = pd.read_excel(path, 'Result Force')
df.drop(df.columns[cols], axis=1, inplace=True)
df.drop(df.index[rows], inplace=True)
df.rename(columns={'Unnamed: 1':'A', 'Unnamed: 5':'B', 'Unnamed: 6':'C', 'Unnamed: 7':'D'}, inplace=True )
date = df.at[2, 'B']
setpoint = df.at[13,'B']
result = df.at[14, 'B']
maxlimit = df.at[37,'B']
testError = df.at[39,'B']