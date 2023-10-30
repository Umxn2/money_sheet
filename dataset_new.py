import openpyxl
import matplotlib
from matplotlib import pyplot
import numpy as np
open_workbook = 'dataset.xlsx'
data_workbook = 'moneyball.xlsx'
workbook = openpyxl.load_workbook(open_workbook)
workbookdata = openpyxl.load_workbook(data_workbook)
ws = workbook.active
ws.cell(row=1, column=7, value = 'MONTH')
ws.cell(row=1, column=8, value = 'MONEY_PER_DAY')
ws.cell(row=1, column=9, value = 'MONEY_PER_MONTH')
workbook.save(open_workbook)

def remove_emptyrows():
    for rows in range(1,1700):
        if(ws.cell(row = rows, column = 4).value == None):
            ws.delete_rows(rows)
            print(ws.cell(row = rows, column = 4).value)
            print("deleting row ", rows)
        else:
            continue
    workbook.save(open_workbook)
def absurd_entries():
    for rows in range(2,1700):
        if(ws.cell(row = rows, column = 4).value!=None):
            if(float(ws.cell(row = rows, column = 4).value) > 1300.0):
                print(ws.cell(row = rows, column = 4).value)
                ws.delete_rows(rows)

        
   
def check_empty_row():
    i = 1
    while (ws.cell(row = i, column = 4).value!=None):
        i=i+1
    return i
def convert_str_to_float():
    i = check_empty_row()
    for rows in range(2,i):
        ws.cell(row = rows, column = 4).value=float(ws.cell(row = rows, column = 4).value)
    workbook.save(open_workbook)
    return

def sum_of_day():
    i = check_empty_row()
    sum = 0
    for rows in range(2, i):
        if(ws.cell(row = rows, column = 2).value==ws.cell(row = rows+1, column = 2).value):
            sum = sum +  ws.cell(row = rows, column = 4).value
        else:
            sum=sum+ws.cell(row = rows, column = 4).value
            ws.cell(row = rows, column = 8).value=sum
            sum = 0
    workbook.save(open_workbook)
def clear_row():
    i = check_empty_row()
    for rows in range(2,i):
        ws.cell(row = rows, column = 6).value=None
    workbook.save(open_workbook)
def get_month():
    i = check_empty_row()
    for rows in range(2,i):
        date= ws.cell(row = rows, column = 2).value
        month = date[3:8]
        ws.cell(row = rows, column = 7).value=month
    workbook.save(open_workbook)
def money_per_month():
    i = check_empty_row()
    sum = 0
    for rows in range(2, i):
        if(ws.cell(row = rows, column = 7).value==ws.cell(row = rows+1, column = 7).value):
            sum = sum +  ws.cell(row = rows, column = 4).value
        else:
            sum=sum+ws.cell(row = rows, column = 4).value
            ws.cell(row = rows, column = 9).value=sum
            sum = 0
    workbook.save(open_workbook)

def grapher():
    i = check_empty_row()
    money_month= []
    month = []
    for rows in range(2,i):
        if(ws.cell(row = rows, column = 9).value!=None):
            
            money_month.append(ws.cell(row = rows, column = 9).value)
            month.append(ws.cell(row = rows, column = 7).value)
    #pyplot.plot(month, money_month)
    return month
    
def graph_per_day():
    i = check_empty_row()
    money_date= []
    date = []
    for rows in range(2,i):
        if(ws.cell(row = rows, column = 8).value!=None):
            
            money_date.append(ws.cell(row = rows, column = 8).value)
            date.append(ws.cell(row = rows, column = 2).value)

    pyplot.plot(date, money_date)
    
    return money_date, date
def clip_data():
    i = check_empty_row()
    for rows in range(2,i):
        if(ws.cell(row = rows, column = 7).value=="03/22" or ws.cell(row = rows, column = 7).value=="10/23" ):
            print(ws.cell(row = rows, column = 7).value)
            ws.delete_rows(rows)
    workbook.save(open_workbook)
def copy_data():

            





# def sum_of_month():
#     i = check_empty_row()
#     sum = 0
#     date =
#     month = ws.cell(row = rows, column = 2).value
    
#     for rows in range(1, i):




        
    
    

             
            
    #workbook.save(open_workbook)

#absurd_entries()
#workbook.save(open_workbook)
#remove_emptyrows()
#convert_str_to_float()
#sum_of_day()
#clear_row()
#get_month()
#money_per_month()


money, date = graph_per_day()
#clip_data()
