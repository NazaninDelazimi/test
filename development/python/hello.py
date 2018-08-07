import xlsxwriter
import datetime
import time
from datetime import datetime

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()
now  = datetime.now()

times = {}
def add_time ():
    num = raw_input(" Number works you have done : ")
    col = raw_input(" enter the num of empty column : ")
    row  = raw_input(" enter the num of empty row : ")
    for i  in range(0,int(num)):
        t = raw_input("please enter time : ")
        w = raw_input("please enter your work : ")
        times[t]  = w
    return times

def add_time_to_excel(times,col,row):
    worksheet.write(row,col ,str (now.month)+" / " + str(now.day)+ " / "+ str(now.year) )
    col+=1
    for time in times:
        print time
        worksheet.write(row, col,time)
        worksheet.write(row+1 , col ,times[time])        
        col+=1

row = 0 
col = 0

add_time_to_excel(add_time(),col,row)


workbook.close()