import pandas as pd
import xlsxwriter
import sys
import time
import os

def analysis(file):
    cols = [0,2,3,4,8,9,10,11,12,13,14,15,16,17,18]
    rows = [15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33]
    df = pd.read_excel(file, 'Result Force',engine='xlrd')
    df.drop(df.columns[cols], axis=1, inplace=True)
    df.drop(df.index[rows], inplace=True)
    df.rename(columns={'Unnamed: 1':'A', 'Unnamed: 5':'B', 'Unnamed: 6':'C', 'Unnamed: 7':'D'}, inplace=True )
    date = df.at[2, 'B']
    setpoint = df.at[13,'B']
    result = df.at[14, 'B']
    maxlimit = df.at[37,'B']
    testError = df.at[39,'B']
    print(setpoint)
    del df
    if testError < maxlimit:
        testresultdf = pd.DataFrame([[date, setpoint, result, testError, maxlimit]],
                                    columns=['Date & Time', 'Setpoint', 'Test Result', 'Test Error', 'Maximum Error Allowed'])
        file = r"C:\Users\m0082668\Documents\Python Projects\File Movers\Federal Cert Checker.xlsx"
        dffile = pd.read_excel(r"C:\Users\m0082668\Documents\Python Projects\File Movers\Federal Cert Checker.xlsx")
        dffile = pd.concat([dffile, testresultdf], ignore_index=True)
        print(dffile)
        writer = pd.ExcelWriter(file, engine="openpyxl")
        dffile.to_excel(writer, sheet_name='Sheet1', index=False)
        writer._save()
        del writer
        #sys.exit()
    else:
        return

        #sheet_name = 'Sheet1'
        #worksheet = writer.sheets[sheet_name]    
        #data1_end = len(dffile)
        #workbook = xlsxwriter.Workbook('Federal Cert Checker.xlsx')
        #chart = workbook.add_chart({'type': 'line'})
        #chart.set_y_axis({'name': 'lbf'})
        #chart.set_x_axis({'name': 'Run No'})
        #chart.set_title({'name': 'Setpoint vs Result'})
        #chart.add_series({
        #    'name': ['Sheet1', 0,1],
        #   'values': ['Sheet1', 1, 1, data1_end, 2]
        #})
        #chart.add_series({
        #    'values':['Sheet1',1,2, data1_end, 2],
        #    'name': ['Sheet1',0,2]
        #})
        #worksheet.insert_chart('I1', chart)
        #writer.save()
            
    