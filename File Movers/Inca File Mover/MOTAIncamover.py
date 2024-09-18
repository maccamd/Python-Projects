
import os
import schedule
import shutil
import time
import datetime
import fnmatch
##--File Moving Program V2--##
#V1 - copy
#V2 - copy2 to try and keep metadata, Current VDC1 (5:45am) & VDC2 (5:10am) Version
"""
Program is designed to move files from one folder to another at a set time every day
Only lines of code that need altering are:
Line 27 - Folder that stores the original files
Line 28 - Folder where files will go to
Line 35 - File format your looking for to move
Line 40 - Can choose if you want to cut or copy the files depending if you use .move or .copy
Line 43 - Time the program runs the backup 

Program checks every second and runs as the minute changes
Program can be made into an .exe using auto-py-to-exe, ensure terminal is selected in conversion settings
"""
print("VDC2 MOTA INCA Files Daily move program running! Backup set for 05:10am every day")
print("To end program press ctrl+c at any point inside the terminal")

def job():
    source_folder = r'D:\\External Customer dat files\\JLR\\'    
    destination_folder = r'W:\\MOTA Inca Files\\'  
    now = datetime.datetime.now()
    for src_dir, dirs, files in os.walk(source_folder):
        dst_dir = src_dir.replace(source_folder, destination_folder, 1)
        if not os.path.exists(dst_dir):
            os.makedirs(dst_dir)
        for file_ in files:
            if fnmatch.fnmatch(file_, '*.mf4'):
                src_file = os.path.join(src_dir, file_)
                dst_file = os.path.join(dst_dir, file_)
                if os.path.exists(dst_file):
                    os.remove(dst_file)
                shutil.copy2(src_file, dst_dir) #.copy for copy paste .move for cut paste
    print("Daily move done on " + now.strftime("%d-%m-%Y %H:%M:%S"))
     
schedule.every().day.at("05:10").do(job)

while 1:
    schedule.run_pending()
    time.sleep(1)