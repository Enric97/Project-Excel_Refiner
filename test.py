from base64 import encode
import pandas as pd # Open Excel
import tkinter as tk    #Window dialog
from tkinter import filedialog
import os



def selectFileWindow():

    root = tk.Tk()
    root.withdraw()

    currentDir=os.getcwd()  #Pillem el directori des de on estem executant
    #Obrim unicament xlsx amb el directori inicial d'asobre
    file_path = filedialog.askopenfilename(initialdir=currentDir, filetypes=(('xlsx files','*.xlsx'),))

    return file_path


fileDirectory=selectFileWindow()



#Llegim el document del termCat que s'ens indica el directori
#La part del index solventa el problema de les cel.les mergeades
termcatDoc = pd.read_excel(fileDirectory, index_col=[0])


#Creem un document d'output
outputNameFile="testingOutput.xlsx"
outputDoc = pd.ExcelWriter(outputNameFile, engine="xlsxwriter")
outputDoc.save()

#Obrim el document d'ouput en una variable
outputDoc = pd.read_excel(outputNameFile)

#-----------------------




print(termcatDoc)











