import pandas as pd
import os
import time
import watchdog.events
import watchdog.observers
import datetime
import time
import openpyxl as xl
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import matplotlib
import matplotlib.pyplot as plt
import win32com.client as win32
import pythoncom
import warnings
 
class Handler(watchdog.events.PatternMatchingEventHandler):
    def __init__(self):
        # Set the patterns for PatternMatchingEventHandler
        watchdog.events.PatternMatchingEventHandler.__init__(self, patterns=['*.xls'],
                                                             ignore_directories=True, case_sensitive=False)
        self.last_created = None
        
    def wait_till_file_is_created(self, source_path):
        time.sleep(5)
        historicalSize = -1
        while (historicalSize != os.path.getsize(source_path)):
            historicalSize = os.path.getsize(source_path)
            time.sleep(1) # Wait for one second
            
        
    def on_created(self, event):
        #print("Watchdog received created event - % s." % event.src_path)
        # Event is created, you can process it now
        file = event.src_path
        if file != self.last_created:
            print(str(datetime.datetime.now()) + " " + str(event))
            self.last_created = file
            daytest = file[63:71]
            daytest = daytest.replace("_", "/")
            #print(daytest)
            warnings.filterwarnings('ignore')
            cols = [0,2,3,4,8,9,10,11,12,13,14,15,16,17,18]
            rows = [15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33]
            self.wait_till_file_is_created(event.src_path)
            df = pd.read_excel(file, 'Result Force',engine='xlrd')
            df.drop(df.columns[cols], axis=1, inplace=True)
            df.drop(df.index[rows], inplace=True)
            df.rename(columns={'Unnamed: 1':'A', 'Unnamed: 5':'B', 'Unnamed: 6':'C', 'Unnamed: 7':'D'}, inplace=True )
            setpoint = df.at[13,'B']
            result = df.at[14, 'B']
            maxlimit = df.at[37,'B']
            testError = df.at[39,'B']
            testresultdf = pd.DataFrame([[daytest, setpoint, result, testError, maxlimit]],
                                            columns=['Date & Time', 'Setpoint', 'Test Result', 'Test Error', 'Maximum Error Allowed'])
            file = r"W:\EDC Schedules\RDEC Weekly Checks Results\VDC 2\Federal Cert Checker.xlsx"
            dffile = pd.read_excel(r"W:\EDC Schedules\RDEC Weekly Checks Results\VDC 2\Federal Cert Checker.xlsx")
            dffile = pd.concat([dffile, testresultdf], ignore_index=True)
            dffile.set_index('Date & Time')
            #print(dffile)
            writer = pd.ExcelWriter(file, engine="openpyxl")
            dffile.to_excel(writer, sheet_name='Sheet1', index=False)
            writer._save()
            matplotlib.use('Tkagg')
            fig, (ax1, ax2) = plt.subplots(2, sharex=True)
            fig.suptitle('Results Vs Error')
            ax1.plot(dffile['Date & Time'], dffile['Test Result'],color='r', label = 'Test Result')
            ax1.plot(dffile['Date & Time'], dffile['Setpoint'], color='k', label = 'Setpoint')
            ax1.set_ylim(190, 194)
            ax2.plot(dffile['Date & Time'], dffile['Test Error'],color='b', label = 'Test Error')
            ax2.plot(dffile['Date & Time'], dffile['Maximum Error Allowed'],color='g', label = 'Maximum Error Allowed')
            ax2.set_ylim(-1.2, 1.5)
            ax1.set_xticklabels(dffile['Date & Time'], rotation=30)
            ax2.set_xticklabels(dffile['Date & Time'], rotation=30)
            ax1.set_ylabel('Force lbf')
            ax2.set_ylabel('Deviation %')
            ax2.set_xlabel('Date')
            ax1.legend(loc='lower right')
            ax2.legend(loc='center left')
            plt.savefig(r'C:\Users\m0082668\Documents\results.png')
                
            wb = xl.load_workbook(file)
            sheet_obj = wb.active
            img = Image(r'C:\Users\m0082668\Documents\results.png')
            sheet_obj.add_image(img, 'I2')
            wb.save(file)
            if testError < maxlimit:
                print("Test Passed")
            else:
                print("Test Failed Criteria Sending Email")
                pythoncom.CoInitialize()
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = 'sam.peter.mcdonald@mahle.com; glenn.lawes@mahle.com; paul.heald@mahle.com; ciaran.durkin@mahle.com'
                mail.Subject = 'Wathchdog Seen Failed Daily Federal Dyno Readiness Calibration'
                mail.Body = ''
                mail.HTMLBody = '<h3>Hello Watchdog here, Today the dyno readiness calibration on VDC2 failed, as a matter of urgency check the data and rectify.</h3>' #this field is optional
                # To attach a file to the email (optional):
                #attachment  = "Path to the attachment"
                #mail.Attachments.Add(attachment)
                mail.Send()
                print("Email Sent") 
        
#"W:\EDC Schedules\RDEC Weekly Checks Results\VDC 2\FEDERAL CERT" 
#"C:\Users\m0082668\Documents\Python Projects\File Movers\Federal Cert Checker.xlsx"
      
if __name__ == "__main__":
    src_path = r"W:\EDC Schedules\RDEC Weekly Checks Results\VDC 2\FEDERAL CERT"
    event_handler = Handler()
    observer = watchdog.observers.Observer()
    observer.schedule(event_handler, path=src_path, recursive=True, )
    observer.start()
    print("Watchdog Running")
    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    
    