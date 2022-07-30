import pandas as pd # Open Excel
import tkinter as tk    #Window dialog
from tkinter import filedialog
import os


#definim variables
termcatDoc = "" # Variable on guardarem el DataFrame (l'excel) del TermCat
newDataFrame = "" # DataFrame (estructura) del nou excel que crearem
diccionary = {} # Diccionary on guardarem f_principal(key) i llista de complementaries (values)
diccionary_arreglat = {} # Diccionary on tindrem escrit adequadament la llista de complementaris (com un string amb |||)
formaPrincipal = "" # Variable utilitzada per iterar sobre el doc del termCat sobre el valor de forma principal
formaComplement = "" # Variable utilitzada per iterar sobre el doc del termCat sobre el valor de les formes complementaries
paraula = "" # Placeholder d'iteracions sobre les diferents formes complementaries
alelex_desc = [] # Llista on guardarem el valor que li toca a aquesta columna del excel que generem (principal + complementaries)
idioma = [] # Llista on tindrem tots els valors de la columna d'idioma del excel que generem
data_inici = [] # Llista on tindrem tots els valors de la comuna de data d'inici del excel que generem
fileDirectory = "" # Directori on tenim el fitxer del termCat


#HARDCODED VARIABLES (WIP)
outputFileName = "" + ".xlsx" # Inidicar el nom de sortida del nou document
language_value = "" # Indicar el que es vol posar en la columna d'IDIOMA del nou Excel
initial_date_value = "" # Indicar que es vol posar en la columna de DATA_INICI

#   ---------- Es poden codificar més en cas de que fos necessari -------


#Finestreta per seleccionar l'arxiu del TermCat
def selectFileWindow():

    root = tk.Tk()
    root.withdraw()

    currentDir=os.getcwd()  #Pillem el directori des de on estem executant
    #Obrim unicament xlsx amb el directori inicial d'asobre
    file_path = filedialog.askopenfilename(initialdir=currentDir, filetypes=(('xlsx files','*.xlsx'),))

    return file_path


def loadExcel(fileDirectory):
    global termcatDoc
    #Llegim el document del termCat que s'ens indica el directori
    #La segona part serveix per arreglar les cel.les combinades
    #La tercera linea definirà que els valors NaN pasin a ser strings buits
    termcatDoc = pd.read_excel(fileDirectory)
    termcatDoc['Forma principal'] = termcatDoc['Forma principal'].fillna(method="ffill") #Arreglem cel.les combinades de la primera filera
    termcatDoc = termcatDoc.fillna('')     #Definim que els valors buits seran ''

 #Creem un DataFrame amb com volem que siguin les coses
 #Basicament estem definint les cel.les de l'excel
def createNewDataFrame():
    global newDataFrame

    newDataFrame = pd.DataFrame(columns=[
            'ALELEX_ID', 
            'ALELEX_DESC',
            'IDIOMA',
            'DATA_INI',
            'DATA_FIN',
            'DESC_CA',
            'DESC_ES',
            'DESC_EN',
            'F_PRINCIPAL',
            'F_COMPLEMENT'])

# Metode que afegeix a un diccionary keys com forma principal i values com formes complemetaries
def createDictionary():

    global diccionary
    global diccionary_arreglat

    # Eliminar ----------------------------
    # diccionary = {} #diccionary on elvalor sera una llista dels diferents opcions
    # diccionary_arreglat = {} #Diccionary on el valor sera la cocatenacio amb format correcte
    # Eliminar ----------------------------

    #Creem el diccionari malo i las keys del diccionary bo
    for row in range(len(termcatDoc.index)): #iterem sobre el numero de filas del doc del Termcat
        
        formaPrincipal= termcatDoc.iloc[row, 0]
        formaComplement = termcatDoc.iloc[row,1]

        if(formaPrincipal not in diccionary):
            diccionary[formaPrincipal] = list()
            diccionary_arreglat[formaPrincipal] =""

        diccionary[formaPrincipal].append(formaComplement)

# Mètode en que posem les formes complementaries correctament (amb |||) al diccionary_arreglat
# Creem els values en el diccionary bo concatenant els diferents values del diccionary malo
def refineValidDictionaryValues():
  
    global alelex_desc
    paraula=""

    for key, value in diccionary.items(): # Iterem per cada item del diccionari
        
        for item in value:  # Iterem per cada forma complementaria
            if(item != ""):
                paraula += item
            if(item != value[-1]):
                paraula += "||| "
        
        diccionary_arreglat[key]=paraula    # Afegim la definició ben escrita al diccionari bó

        if(paraula!=""):        # Aprofitem la iteració per crear el contingut del alalex_desc, que es la suma dels anteriors                
            alelex_desc.append(key +"||| "+ paraula)
        else:
            alelex_desc.append(key)

        paraula=""            # Reiniciem variable  


def firstHalfDataFrame():
    # Afegim el diccionary arreglat al dataframe
    global newDataFrame
    newDataFrame["F_PRINCIPAL"]=diccionary_arreglat.keys()
    newDataFrame["F_COMPLEMENT"]=diccionary_arreglat.values()
    newDataFrame["ALELEX_DESC"]= alelex_desc

# Afegim els items necesaris a las columnes d'idioma i data_inici
def completeLanguageAndInitialDateColumns():
    global idioma
    global data_inici
    global newDataFrame

    for row in range(0,len(newDataFrame.index)): # Totes son iguals, i n'hi ha tantes com files...
        idioma.append("ca")
        data_inici.append("01012021")


# Afegim l'idioma i data d'inici al dataFrame
def secondHalfDataFrame():
    global newDataFrame
    newDataFrame["IDIOMA"]= idioma
    newDataFrame["DATA_INI"]= data_inici


# Exportem a Excel
def exportingToExcel():
    global newDataFrame
    newDataFrame.to_excel("really.xlsx", index=False)





fileDirectory=selectFileWindow()

loadExcel(fileDirectory)
createNewDataFrame()
createDictionary()
refineValidDictionaryValues()
firstHalfDataFrame()
completeLanguageAndInitialDateColumns()
secondHalfDataFrame()
exportingToExcel()

print(newDataFrame)











