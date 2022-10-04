from openpyxl import *
#lendo a planilha
planilha = load_workbook(
    filename="E:\Programas em python\Python com Excel\TimeSheet-ThacioLopes-2022-08.xlsx")

# Mexendo na aba Time-Sheet
timeSheet = planilha['Time-Sheet']

lista= ['1', '2', '3']

data= '1/1/22  2/2/22  '

newdata= data + data

print(lista)

list(newdata)

print(newdata)

newdata = newdata.split()





print(len(lista))




print(newdata)

print(type(newdata))

print(len(newdata))

