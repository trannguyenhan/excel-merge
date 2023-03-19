# Import libraries
import numpy as np
import pandas as pd
import os
import os.path
import datetime as dt

# Print output to be displayed in terminal
print("\nBegin merge script.\nScanning for files.\n\nImporting files:")

# Get all .xlsx files from project root recursively
targetF = []
errFiles = ""
validFiles = ""
validCnt = 0
errCnt = 0

# extract header from list file excel
def get_header(lst_target_excel):
    # get header from each file excel
    header = pd.read_excel(lst_target_excel[0], header=None)
    header = header.loc[0:1]

    header2 = pd.read_excel(lst_target_excel[len(lst_target_excel) - 1], header=None)
    header2 = header2.loc[0:1]

    header[4][0] = header[4][0].replace("Page", "Page From")
    header2[4][0] = header2[4][0].replace("Page", "Page To")

    header[6][0] = header2[4][0]

    return header

for root, subdirs, files in os.walk(os.curdir):
    for f in files:
        if f.endswith('.xls') and not (f == "master_table.xls" or f == "~$master_table.xls"):
            try:
                checkfile = open(os.path.join(root, f))
                checkfile.close()
                validCnt += 1
                print("  " + os.path.join(root, f))
                validFiles += "  " + os.path.join(root, f) + "\n"
                targetF.append(os.path.join(root, f))
            except:
                errCnt += 1
                errFiles += "  " + os.path.join(root, f[2:]) + "\n"
                pass

# Output a log file from the merge process
if errCnt == 0:
    logMsg = "Process executed at: {}\n\nFiles imported:\n".format(str(dt.datetime.now()).split('.')[0]) + validFiles + "\nNumber of files imported = {}\n\nImport successful!".format(validCnt)
    log_file = open('log.txt', 'w')
elif errCnt > 0:
    logMsg = "Process executed at: {}\n\nFiles imported:\n".format(str(dt.datetime.now()).split('.')[0]) + validFiles + "{} file(s) imported.\n\nWarning!!! Some files were in use during the time of import.".format(validCnt) \
    + "\nRun the script again in case there are changes from the affected files.\n\n{} file(s) affected:\n".format(errCnt) \
    + errFiles
    log_file = open('log.txt', 'w')
    print("\n{} file(s) imported.\nWarning!!! {} file(s) in use during runtime:\n".format(validCnt, errCnt) + \
          errFiles + "There may have been changes to the affected files.")

log_file.write(logMsg)
log_file.close()

# Load all .xlsx to dataframes and concatenate into master dataframe
dataframes = []

try:
    # get header from each file excel
    header = get_header(targetF)
    dataframes.append(header)
except:
    print("No header")

for t in targetF:
    dataframe = pd.read_excel(t, skiprows=[0,1], header=None)
    dataframes.append(dataframe)

# dataframes = [pd.read_excel(t, skiprows=[1,2], header=None) for t in targetF]
# df_master = pd.concat(header)
df_master = pd.concat(dataframes)

# Exports a consolidated excel file 
if os.path.exists('master_table.xls'):
    try:
        mstFile = open('master_table.xls','r')
        mstFile.close()
        df_master.to_excel('master_table.xls', index = False, header=False)
        print("\nMerge complete!")
    except PermissionError:
        errMsg = "\nERROR!!! UPDATE FAILED! Please close the master_table.xlsx file and run the script again."
        print(errMsg)
        # open('log.txt', 'a').write(errMsg)   
        open('log.txt', 'w').write(errMsg)
else:
    df_master.to_excel('master_table.xls', index = False)
    print("\nMerge complete!")
            
