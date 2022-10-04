from openpyxl import *

dias= 1
fim= dias + 3
mes= 8
ano = 22
# print(fim)

planilha = load_workbook(
    filename="E:\Programas em python\Python com Excel\TimeSheet-ThacioLopes-2022-08.xlsx")

# Mexendo na aba Time-Sheet
timeSheet = planilha['Time-Sheet']

datas= [ '1/1/22', '1/1/22', '2/1/22', '2/1/22', '3/1/22', '3/1/22', '4/1/22', '4/1/22', '5/1/22', '5/1/22']


# while (dias<=len(datas)):
#     print(dias)
#     dias= dias + 1
#     print('--------')


linha=1


    # print(lista_compras[1])

















# datas = [dias]
#
# for data in datas:
#     print(data)
#
#
# def inserir_dia(dia):
#     dia =1
#     ano = 22
#     dias = f'{dia}/{mes}/{ano}'
#     return dias
#
# print(inserir_dia(5))