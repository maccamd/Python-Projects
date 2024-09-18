import os, shutil

path_to_watch = 'C:\\Users\m0082668\Desktop\Test\\'
print('Your folder path is"',path_to_watch,'"')
print('The destination path is P:\VDC Project\\')

old = os.listdir(path_to_watch)
print(old)

while True:
    new = os.listdir(path_to_watch)
    if len(new) > len(old):
        newfile = list(set(new) - set(old))
        print(newfile[0])
        old = new
        extension = os.path.splitext(path_to_watch + "/" + newfile[0])[1]
        if extension == ".xlsx":
            folderTo =  (os.path.expanduser("P:\\VDC Project\\"))
            files = os.listdir(path_to_watch)
            mover = newfile[0]
            shutil.move(path_to_watch + mover, folderTo)
            continue            
    else:
        continue