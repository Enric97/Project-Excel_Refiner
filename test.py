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
alelex_desc = [] # Llista on guardarem el valor que li toca a aquesta columna del excel que generem (principal + complementaries)
idioma = [] # Llista on tindrem tots els valors de la columna d'idioma del excel que generem
data_inici = [] # Llista on tindrem tots els valors de la comuna de data d'inici del excel que generem
fileDirectory = "" # Directori on tenim el fitxer del termCat


#HARDCODED VARIABLES (WIP)
outputFileName = "" # Inidicar el nom de sortida del nou document
language_value = "" # Indicar el que es vol posar en la columna d'IDIOMA del nou Excel
initial_date_value = "" # Indicar que es vol posar en la columna de DATA_INICI
outputFileDirectory = os.getcwd()+"/folder_out" # Indicar el directory del arxiu de sortida

#   ---------- Es poden codificar més en cas de que fos necessari -------


# <----------------- COMENCEN ELS DIFERENTS MÈTODES (ORDENATS) ----------------->

# Finestreta per seleccionar l'arxiu del TermCat
def selectFileWindow():
    global outputFileName
    root = tk.Tk()
    root.withdraw()

    currentDir=os.getcwd()  #Pillem el directori des de on estem executant
    # Obrim unicament xlsx amb el directori inicial d'asobre, e indiquem titol de la finestreta
    file_path = filedialog.askopenfilename(
        initialdir=currentDir, 
        title="Selecciona l'arxiu del TermCat",
        filetypes=(('xlsx files','*.xlsx'),)
    )

    # Per que el arxiu de sortida contingui el nom del arxiu d'entrada
    outputFileName= os.path.splitext(os.path.basename(file_path))[0]

    return file_path


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

    #Creem el diccionari malo i las keys del diccionary bo
    for row in range(len(termcatDoc.index)): #iterem sobre el numero de filas del doc del Termcat
        
        formaPrincipal= termcatDoc.iloc[row, 0]
        formaComplement = termcatDoc.iloc[row,1]

        if(formaPrincipal not in diccionary):
            formaPrincipal = formaPrincipal.strip()     #Fem que la key tampoc tingui salts de linia ni espais en blanc al principi i final del text (important de cara al diccionary_arreglat)
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
            alelex_desc.append(key +" ||| "+ paraula)
        else:
            alelex_desc.append(key)

        paraula=""            # Reiniciem variable  


# Afegim el diccionary arreglat al dataframe, al menys les 3 columnes que tenim fins ara
def firstHalfDataFrame():
    global newDataFrame
    global alelex_desc

    newDataFrame["F_PRINCIPAL"]=diccionary_arreglat.keys()
    newDataFrame["F_COMPLEMENT"]=diccionary_arreglat.values()
    newDataFrame["ALELEX_DESC"]= alelex_desc
    newDataFrame["DESC_CA"]= alelex_desc


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


# Exportem a Excel i a txt (tindrem dos arxius iguals, amb diferent format)
def exportingToExcel():
    global newDataFrame
    global outputFileName
    global outputFileDirectory

    # newDataFrame.to_excel("really.xlsx", index=False)
    # newDataFrame.to_csv("really.txt", sep="\t", index=False)
    outputDirectoryFile = outputFileDirectory + "/" + outputFileName

    newDataFrame.to_excel(outputDirectoryFile+"_refined.xlsx", index=False)
    newDataFrame.to_csv(outputDirectoryFile+"_refined.txt", sep="\t", index=False)


# <----------------- MAIN ----------------->


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











