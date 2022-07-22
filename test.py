from base64 import encode
import pandas as pd # Open Excel
import tkinter as tk    #Window dialog
from tkinter import filedialog
import os
from openpyxl import load_workbook
import xlrd


def selectFileWindow():

    root = tk.Tk()
    root.withdraw()

    currentDir=os.getcwd()  #Pillem el directori des de on estem executant
    #Obrim unicament xlsx amb el directori inicial d'asobre
    file_path = filedialog.askopenfilename(initialdir=currentDir, filetypes=(('xlsx files','*.xlsx'),))

    return file_path


# fileDirectory=selectFileWindow()
# print(fileDirectory)


doc = pd.read_excel(r'D:\Github\personal\Project-TE2T\folder_in\provaS11.xlsx')
# wb = xlrd.open_workbook('folder_in\provaS11.xlsx', formatting_info=True)

# sheet = wb.sheet_by_name("Hoja1")
# f = open("testing.txt", 'w')
# f.write(doc.iloc[0][0])
print(doc)














