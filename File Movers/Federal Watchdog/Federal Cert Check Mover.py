import pandas as pd
import os
import time
import watchdog.events
import watchdog.observers
import time
import xlsxwriter
import openpyxl as xl
import analysis
 
class Handler(watchdog.events.PatternMatchingEventHandler):
    def __init__(self):
        # Set the patterns for PatternMatchingEventHandler
        watchdog.events.PatternMatchingEventHandler.__init__(self, patterns=['*.xls'],
                                                             ignore_directories=True, case_sensitive=False)
        
    def on_created(self, event):
        print("Watchdog received created event - % s." % event.src_path)
        # Event is created, you can process it now
        file = event.src_path
        print(event.src_path)
        analysis.analysis(file)
        
#"W:\EDC Schedules\RDEC Weekly Checks Results\VDC 2\FEDERAL CERT"       
if __name__ == "__main__":
    src_path = r"C:\Users\m0082668\Desktop\To Folder"
    event_handler = Handler()
    observer = watchdog.observers.Observer()
    observer.schedule(event_handler, path=src_path, recursive=True, )
    observer.start()
    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    
    