from asyncio.windows_events import NULL
from multiprocessing.reduction import duplicate
import pandas as pd
from datetime import datetime
from tkinter import *
from tkinter import ttk
import tkinter.filedialog as dlg
from tkinter.messagebox import showwarning

def filter_For_Current_Year(dataFrame, column) :
    filterActualYear = dataFrame[dataFrame[column].dt.year == currentTime.year]
    return filterActualYear

def filter_For_Month(dataFrame, column, numMonth) : 
    filterMonth = dataFrame[dataFrame[column].dt.month == numMonth]
    return filterMonth

def create_File_Excel(dataFrame, filePathTxt) : 
    # change TXT to xlsx
    filePathExcel = filePathTxt.replace('TXT', 'xlsx')
    # create file excel
    dataFrame.to_excel(filePathExcel, index=False)

def get_Path_Arqtxt() :
    global filePath
    filePath = dlg.askopenfilename()

def get_Values_By_Arqtxt() :
    global arqTexto
    try :
        chooseMonth = int(boxChooseMonth.get())
        with open(filePath) as op:
            arqTexto = pd.read_csv(op, sep='$', header=0, encoding='ANSI', low_memory=False)
    except :
        showwarning(title = "Aviso", message = "Erro de arquivo: \n Verifique se selecionou um arquivo com extensão TXT \n Verifique se selecionou o mês desejado")
    # change the type of colunms
    arqTexto['ACTUAL DATE'] = convert_To_Date(arqTexto, 'ACTUAL DATE')
    arqTexto['AVAILABLE DATE'] = arqTexto['AVAILABLE DATE'].astype('datetime64[ns]')
    # drop rows duplicates 
    print(len(arqTexto.index))
    arqTexto = drop_Rows_Duplicates(arqTexto)
    print(len(arqTexto.index))
    # drop column
    #arqTexto = arqTexto.drop(columns='ORDER NUMBER')
    # replace "? for NULL
    #arqTexto["ORDER NUMBER"] = arqTexto["ORDER NUMBER"].replace("?", NULL)
    arqTexto["ORDER NUMBER"] = arqTexto["ORDER NUMBER"].astype(str, copy = False, errors = 'ignore')

    # make the filter in the file
    filterReady = filter_Rows(arqTexto,'TASK STATUS', 'Ready')
    infoValues = "Ready = " + str(len(filterReady.index)) + "\n"
    # filter to get the complete of the year and month
    filterYear = filter_For_Current_Year(arqTexto, 'ACTUAL DATE')
    infoValues += "Concluidos do ano = " + str(len(filterYear.index)) + "\n"
    filterMonth = filter_For_Month(filterYear, 'ACTUAL DATE', chooseMonth)
    infoValues += "Concluidos do mes " + str(chooseMonth) + " = " + str(len(filterMonth.index)) + "\n"
    # filter to get the pending of year and month
    filterOnlyYear = filter_For_Current_Year(filterYear, 'AVAILABLE DATE')
    infoValues += "Entrantes do ano = " + str(len(filterOnlyYear.index)) + "\n"
    filterAvailableYear = filter_For_Current_Year(filterMonth, 'AVAILABLE DATE')
    filterAvailableMonth = filter_For_Month(filterAvailableYear, 'AVAILABLE DATE', chooseMonth)
    infoValues += "Entrantes do mes " + str(chooseMonth) + " = " + str(len(filterAvailableMonth.index)) + "\n"
    # print on the window
    textValues['text'] = infoValues

# remover duplicatas
def drop_Rows_Duplicates(dataFrame) :
    duplicatesOut = dataFrame.drop_duplicates() # linhas totalmente iguais
    return duplicatesOut
    
def filter_Rows(dataFrame, column, cell) : 
    filter = dataFrame[(dataFrame[column]==cell)]
    return filter

def count_Amount_Rows(dataFrame) : 
    amountRows = len(dataFrame.index)
    return amountRows

def convert_To_Date(dataFrame, column) :
    dataFrame[column] = pd.to_datetime(dataFrame[column], dayfirst = True, format='%d/%m/%Y')
    return dataFrame[column] 

def exeProgram() :
    get_Values_By_Arqtxt()
    if (result) : 
        create_File_Excel(arqTexto, filePath)

def checkBox() :
    global result
    result = 0
    if (buildExcel.get()) : 
        result = 1
    else : result = 0

#main
currentTime = datetime.now()

windows = Tk()
windows.title("Consulta de valores")
labelFile = Label(windows, text="Deve ser escolhido um arquivo (TXT)")
labelFile.grid(column = 1, row = 0, pady = 2)

buttonCreateExcelTest = Button(windows, text="Importar arquivo", command=get_Path_Arqtxt)
buttonCreateExcelTest.grid(column = 1, row = 1, pady = 2)

buildExcel = IntVar() 
boxBuildExcel = Checkbutton(windows, text="Criar arquivo excel", variable=buildExcel, onvalue=1, offvalue=0, command=checkBox)
boxBuildExcel.grid(column = 1, row = 2, pady = 2)

labelBoxMonth = Label(windows, text="Escolha o mês")
labelBoxMonth.grid(column = 1, row = 3, pady = 2)
listMonth = [1,2,3,4,5,6,7,8,9,10,11,12]
boxChooseMonth = ttk.Combobox(windows, values=listMonth)
boxChooseMonth.grid(column = 1, row = 4, pady = 2)

labelValues = Label(windows, text = "Valores:") 
labelValues.grid(column = 1, row = 5, pady = 2)

textValues = Label(windows, text="")
textValues.grid(column = 1, row = 6, pady = 2)

buttonConfirm = Button(windows, text="Confirmar", command=exeProgram)
buttonConfirm.grid(column = 2, row = 7, pady = 2)
windows.mainloop()
