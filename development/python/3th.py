from xlrd import open_workbook
from xlutils.copy import copy
import time
from datetime import datetime


rb = open_workbook("names.xlsx")
wb = copy(rb)
now  = datetime.now()

s = wb.get_sheet(0)






times = {}
def add_time ():
    num = raw_input(" Number works you have done : ")
    if num==0:
        return
    for i  in range(0,int(num)):
        t = raw_input("please enter time : ")
        w = raw_input("please enter your work : ")
        times[t]  = w
    
    col = raw_input(" enter the num of empty column : ")
    row  = raw_input(" enter the num of empty row : ")
    int_row = int(row)
    int_col= int(col)
    
    s.write(int(row),int(col) ,str (now.month)+" / " + str(now.day)+ " / "+ str(now.year) )
    int_col  = int_col +1
    for time in times:
        print time
        s.write(int_row,int_col,time)
        s.write(int_row+1 , int_col ,times[time])        
        int_col+=1


    

int_row = 0 
int_col = 0

add_time()


wb.save('names.xlsx')