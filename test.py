import pandas as pd # Open Excel
import tkinter as tk    # Window dialog
from tkinter import filedialog
import os   # Directoris


# <----------------- DEFINICIÓ DE LES VARIABLES ----------------->

termcatDoc = "" # Variable on guardarem el DataFrame (l'excel) del TermCat
newDataFrame = "" # DataFrame (estructura) del nou excel que crearem
diccionary = {} # Diccionary on guardarem f_principal(key) i llista de complementaries (values)
diccionary_arreglat = {} # Diccionary on tindrem escrit adequadament la llista de complementaris (com un string amb |||)
formaPrincipal = "" # Variable utilitzada per iterar sobre el doc del termCat sobre el valor de forma principal
formaComplement = "" # Variable utilitzada per iterar sobre el doc del termCat sobre el valor de les formes complementaries
paraula = "" # Placeholder d'iteracions sobre les diferents formes complementaries
_descColumn = [] # Llista on guardarem el valor que li toca a aquesta columna del excel que generem (principal + complementaries)
idioma = [] # Llista on tindrem tots els valors de la columna d'idioma del excel que generem
data_inici = [] # Llista on tindrem tots els valors de la comuna de data d'inici del excel que generem
fileDirectory = "" # Directori on tenim el fitxer del termCat
outputFileDirectory = os.getcwd()+"/folder_out" # Indicar el directory del arxiu de sortida
outputFileName = "" # Inidicar el nom de sortida del nou document
outputDirectoryFile = "" # Suma dels directori de sortida i nom del arxiu
columnName = "" # Nom de la columna demanada, tipus "ALELEX" o "MENLEX"
index_column = []

#---------- NEw Vars
definitionAndNotes = False
definitionColumn = []
notesColumn = []

#HARDCODED VARIABLES (WIP)
language_value = "" # Indicar el que es vol posar en la columna d'IDIOMA del nou Excel
initial_date_value = "" # Indicar que es vol posar en la columna de DATA_INICI

#   ---------- Es poden codificar més en cas de que fos necessari -------


# <----------------- COMENCEN ELS DIFERENTS MÈTODES (ORDENATS) ----------------->

# Finestreta per seleccionar l'arxiu del TermCat
def selectFileWindow():
    root = tk.Tk()
    root.withdraw()

    currentDir=os.getcwd()  #Pillem el directori des de on estem executant
    # Obrim unicament xlsx amb el directori inicial d'asobre, e indiquem titol de la finestreta
    file_path = filedialog.askopenfilename(
        initialdir=currentDir, 
        title="Selecciona l'arxiu del TermCat",
        filetypes=(('xlsx files','*.xlsx'),)
    )
    return file_path

# Demanem al usuari que ens indiqui el nom de la priemra i segona columna, que també utilitzem com nom de arxiu
def setColumnName():
    global columnName
    global outputFileName

    columnName = input("Inica el nom de la columna que fa referencia al document (ex. ALELEX o MENLEX)\n\t")
    columnName = columnName.upper()
    outputFileName = columnName



# Llegim el document del termCat que s'ens indica el directori
def loadExcel(fileDirectory):
    global termcatDoc
    
    #La segona part serveix per arreglar les cel.les combinades
    #La tercera linea definirà que els valors NaN pasin a ser strings buits
    termcatDoc = pd.read_excel(fileDirectory)
    termcatDoc['Forma principal'] = termcatDoc['Forma principal'].fillna(method="ffill") #Arreglem cel.les combinades de la primera filera
    termcatDoc = termcatDoc.fillna('')     #Definim que els valors buits seran ''


# Creem un DataFrame amb com volem que siguin les coses
# Basicament estem definint les cel.les de l'excel
def createNewDataFrame():
    global newDataFrame
    global columnName

    newDataFrame = pd.DataFrame(columns=[
            columnName+'_ID', 
            columnName+'_DESC',
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

    #Creem el diccionari malo i las keys del diccionary bo
    for row in range(len(termcatDoc.index)): #iterem sobre el numero de filas del doc del Termcat
        
        formaPrincipal= termcatDoc.iloc[row, 0]
        formaPrincipal = formaPrincipal.strip()     #Fem que la key tampoc tingui salts de linia ni espais en blanc al principi i final del text (important de cara al diccionary_arreglat)
        formaPrincipal = formaPrincipal.replace("\n","")
        formaComplement = termcatDoc.iloc[row,1]

        if(formaPrincipal not in diccionary.keys()):
            diccionary[formaPrincipal] = list()
            diccionary_arreglat[formaPrincipal] =""

        diccionary[formaPrincipal].append(formaComplement)


# Mètode en que posem les formes complementaries correctament (amb |||) al diccionary_arreglat
# Creem els values en el diccionary bo concatenant els diferents values del diccionary malo
def refineValidDictionaryValues():
    global _descColumn
    global diccionary
    global diccionary_arreglat

    paraula=""

    for key, value in diccionary.items(): # Iterem per cada item del diccionari
        
        for item in value:  # Iterem per cada forma complementaria
            if(item==""):
                break
            else:              
                paraula += item.strip()     # Mètode que elimina els whitespaces del principi i final del string --> així no quedan desquadrats els espais
                if(item != value[-1]):
                    paraula += " ||| "
        
        paraula = paraula.replace("\n","")  # Per eliminar salts de línia que hi han randoms en el document del Termcat
        diccionary_arreglat[key]=paraula    # Afegim la definició ben escrita al diccionari bó

        if(paraula!=""):        # Aprofitem la iteració per crear el contingut del alalex_desc, que es la suma dels anteriors                
            key=key.strip()     # Igual que amb paraula, eliminem els espais que pougui tenir, així aconseguim un format uniforme
            key = key.replace("\n","")
            _descColumn.append(key +" ||| "+ paraula)
        else:
            _descColumn.append(key)

        paraula=""            # Reiniciem variable  


# Afegim el diccionary arreglat al dataframe, al menys les 3 columnes que tenim fins ara
def firstHalfDataFrame():
    global newDataFrame
    global _descColumn
    global columnName

    newDataFrame["F_PRINCIPAL"]=diccionary_arreglat.keys()
    newDataFrame["F_COMPLEMENT"]=diccionary_arreglat.values()
    newDataFrame[columnName+"_DESC"]= _descColumn
    newDataFrame["DESC_CA"]= _descColumn


# Afegim els items necesaris a las columnes d'idioma i data_inici
def completeLanguageAndInitialDateColumns():
    global idioma
    global data_inici
    global newDataFrame
    global index_column

    for row in range(0,len(newDataFrame.index)): # Totes son iguals, i n'hi ha tantes com files...
        idioma.append("ca")
        data_inici.append("01012021")
        index_column.append(row+1)



# Afegim l'idioma i data d'inici al dataFrame
def secondHalfDataFrame():
    global newDataFrame
    global index_column

    newDataFrame["IDIOMA"]= idioma
    newDataFrame["DATA_INI"]= data_inici
    newDataFrame[columnName+"_ID"]=index_column


# Exportem a Excel i a txt (tindrem dos arxius iguals, amb diferent format)
def exportingToExcel():
    global newDataFrame
    global outputFileName
    global outputFileDirectory
    

    # newDataFrame.to_excel("really.xlsx", index=False)
    # newDataFrame.to_csv("really.txt", sep="\t", index=False)
    outputDirectoryFile = outputFileDirectory + "/catàleg_REF_" + outputFileName

    newDataFrame.to_excel(outputDirectoryFile+".xlsx", index=False)
    newDataFrame.to_csv(outputDirectoryFile+".txt", sep="\t", index=False)


# -----------Other coloumns --------

def askForAdditionalColumns():
    global definitionAndNotes

    answer = input("Existeixen les columnes de definicio i notes? (y/n)\n\t")
    if(answer.casefold =="y"):
        definitionAndNotes = True
    
    return definitionAndNotes

def loadAdditionalColumns():
    global termcatDoc

    termcatDoc['Definició'] = termcatDoc['Definició'].fillna(method="ffill")
    termcatDoc['Notes'] = termcatDoc['Notes'].fillna(method="ffill")
















# <----------------- MAIN ----------------->


fileDirectory=selectFileWindow()

setColumnName()

askForAdditionalColumns()

loadExcel(fileDirectory)

# if(definitionAndNotes):
#     loadAdditionalColumns()

# print(termcatDoc)

createNewDataFrame()
createDictionary()
refineValidDictionaryValues()
firstHalfDataFrame()
completeLanguageAndInitialDateColumns()
secondHalfDataFrame()
exportingToExcel()

print(newDataFrame)











