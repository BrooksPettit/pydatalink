#Brooks Pettit
#832-580-7346
#brooks.pettit@live.com

"""
This program is designed to open a list of PI pages in order; I have a VBA macro that can be run on each page to create a .csv file containing
all of the PI tags on the page. Currently there is no way to apply the macro to multiple pages at once. Therefore, a user must copy and paste the
macro into a module on each PI page. This program will search the G drive for whitelisted PI pages and open them for the user in sequence to streamline
the VBA code insertion.
"""


import os #Module to navigate windows file system
import subprocess #module to execute batch files written by the program
from tqdm import * #module to display a progress bar
import pandas as pd #module to read from the PI page whitelist file already created




os.chdir(r'G:\Local Initiatives\Public Share\2020 PINCH\PI pages') #working directory contains PI page file system


df=pd.read_excel(r"C:\Users\bpettit1\Documents\pythonProjects\pi_pages_selected.xlsx")
whitelist=list(df['page']) #create list of PI pages to search for

pi_pages=[]
paths=[]
count=0
for root, dirs, files in os.walk(".", topdown=False): #walk the \procbook directory tree
   for name in files:
      if name[-3:] in ('PDI','pdi'): #only include .pdi filetypes and filenames which are in the whitelist 'and name in whitelist'
        pi_pages.append(name)
        paths.append(os.path.join(os.getcwd(),root[2:],name)) #create an absolute path to the selected file
        count+=1



#this preamble will cause each batch file to wait until PI ProcessBook is no longer running before returning
#prevents this program from opening 400 PI pages at once
#exit PI to continue to the next file
batch_loop=""":LoopStart
timeout /t 1 /nobreak > nul
tasklist /FI "IMAGENAME eq Procbook.exe" 2>NUL | find /I /N "Procbook.exe">NUL
if "%ERRORLEVEL%"=="1" GOTO LoopEnd
GOTO LoopStart
:LoopEnd"""



print("Program to iterate through PI pages in a file system.")
input("Press the enter key to begin:")
print("\n")

t = tqdm(total=len(pi_pages)) #instantiate a progress bar

for page in paths:
    #print(r'"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PI System\PI ProcessBook.lnk"'+' '+page+'"\n')
    batch_file=open(r"C:\users\bpettit1\documents\temp.bat",'w') #overwrite the temporary file
    batch_file.write("@echo off\n\n")
    batch_file.write(r'"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PI System\PI ProcessBook.lnk"'+' "'+page+'"\n') #call PI from the shortcut and pass the absolute path to a PI page (.pdi file)
    batch_file.write(batch_loop) #wait for PI to close before returning
    batch_file.close()
    subp_i=subprocess.Popen(r'C:\users\bpettit1\documents\temp.bat') #execute the batch file
    subp_i.wait() #this program waits until the batch file returns (i.e. when PI is closed by the user)
    t.update(1) #update the progress bar





