import filecmp

direc1 = r'W:\\Temp VDC Backup\\'
direc2 = r'P:\\VDC Project\\Backup\\'

match, mismatch, errors = filecmp.cmpfiles(direc1, direc2, shallow=False)

print("Not There: ", mismatch)
print("Error: ", errors)