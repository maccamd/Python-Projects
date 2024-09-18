import os
import schedule
import shutil
import time
import datetime
##--File Moving Program--##
"""
Program is designed to move files from one folder to another at a set time every day
Only lines of code that need altering are:
Line 22 - Folder that stores the original files
Line 23 - Folder where files will go to
Line 34 - Can choose if you want to cut or copy the files depending if you use .move or .copy
Line 37 - Time the program runs the backup 

Program checks every second and runs as the minute changes
Program can be made into an .exe using auto-py-to-exe, ensure terminal is selected in conversion settings
"""
print("Daily back up of VDC2 files program is running! Backup set for 01:00am every day")
print("To end program press ctrl+c at any point inside the terminal")

def job():
    source_folder = r'W:\\Temp VDC Backup\\'    
    destination_folder = r'P:\\VDC Project\\Backup\\'  
    now = datetime.datetime.now()
    for src_dir, dirs, files in os.walk(source_folder):
        dst_dir = src_dir.replace(source_folder, destination_folder, 1)
        if not os.path.exists(dst_dir):
            os.makedirs(dst_dir)
        for file_ in files:
            src_file = os.path.join(src_dir, file_)
            dst_file = os.path.join(dst_dir, file_)
            if os.path.exists(dst_file):
                os.remove(dst_file)
            shutil.move(src_file, dst_dir) #.copy for copy paste .move for cut paste
    print("Daily backup done on " + now.strftime("%d-%m-%Y %H:%M:%S"))
     
schedule.every().day.at("01:00").do(job)

while 1:
    schedule.run_pending()
    time.sleep(1)